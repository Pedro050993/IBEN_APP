"""Monta o e-mail de fechamento mensal da analise de custo.

Saida: .eml que abre no Outlook como mensagem nova (X-Unsent: 1), previa.html e
conferencia.json. O app nunca envia: quem confere e clica em Enviar e o engenheiro.
"""
import datetime as dt
import html
import json
import mimetypes
import os
import re
from email.message import EmailMessage
from email.utils import make_msgid

from iben_app.comum import br, norm
from iben_app.fechamento import leitura

MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO",
         "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
DOMINIO_INTERNO = "iben.eng.br"


class ObraParada(Exception):
    """Problema que para a obra (ex.: mes da planilha diferente do fechamento)."""


def mes_extenso(mes):
    a, m = int(mes[:4]), int(mes[5:7])
    return f"{MESES[m - 1]}/{str(a)[2:]}"


SIGLAS = {"incc", "inss", "fgts", "crea", "avcb", "sinapi", "totvs", "pbqp", "siac", "igpm", "ipca", "opr",
          "igile", "spe", "epc", "nbr", "art", "rrt", "rdo", "fvs", "eap", "sm", "oc", "pc", "mo", "nf", "cnpj",
          "ltda", "bdi", "vgv", "cub", "pls"}
PALAVRAS_CURTAS = {"de", "da", "do", "das", "dos", "e", "em", "com", "para", "por", "na", "no", "nas", "nos", "ao",
                   "aos", "a", "o", "as", "os", "um", "uma", "sem", "mes", "que", "fim", "dia", "uso", "via", "ser",
                   "tem", "mais", "pela", "pelo", "sob", "ate", "mao", "obra", "item", "lote", "fase", "piso", "muro",
                   "novo", "nova", "teto", "rede", "agua", "gas", "luz", "vao", "ou", "se", "foi", "nao", "sim"}


def caixa_normal(s):
    """Justificativa em caixa normal, preservando siglas e codigos (palavras com digito)."""
    if not s:
        return ""
    s = " ".join(str(s).split())
    if s != s.upper():
        return s
    out = []
    for w in s.split(" "):
        letras = norm(re.sub(r"[^A-Za-zÀ-ÿ]", "", w))
        sigla = letras in SIGLAS or (len(letras) <= 3 and letras and letras not in PALAVRAS_CURTAS)
        out.append(w if re.search(r"\d", w) or sigla else w.lower())
    t = " ".join(out)
    return t[:1].upper() + t[1:]


def _rs(v):
    return ("-R$ " if v < 0 else "R$ ") + br(abs(v))


def resumo_obra(af, nota_fixa=None):
    """Texto (lista de paragrafos) e alertas de uma obra, so com fatos da planilha."""
    al = list(af["alertas"])
    tot = af["total"]
    no_total = [l for l in af["linhas"] if l["no_total"]]
    soma_var = round(sum(l["var"] for l in no_total), 2)
    soma_incc = round(sum(l["incc"] for l in no_total), 2)
    if tot:
        var_total = tot["var"] or 0.0
        if abs(soma_var - var_total) > 1:
            al.append(f"soma das variações por EAP ({br(soma_var)}) não fecha com o total ({br(var_total)})")
        if tot["dif"] is not None and tot["ant"] is not None and abs((tot["dif"] - tot["ant"]) - var_total) > 1:
            al.append(f"diferença atual − mês anterior ({br(tot['dif'] - tot['ant'])}) não fecha com a variação ({br(var_total)})")
        if tot["incc"] is not None and abs((tot["incc"] or 0) - soma_incc) > 1:
            al.append(f"total da coluna INCC ({br(tot['incc'])}) difere da soma por EAP ({br(soma_incc)})")
        if tot["formula_incc"] is not None and not str(tot["formula_incc"]).startswith("="):
            al.append("total da coluna INCC está digitado, não é fórmula")
        dif, ant = tot["dif"] or 0.0, tot["ant"] or 0.0
    else:
        var_total, dif, ant = soma_var, None, None
    fora_incc = []
    for l in no_total:
        resto = round(l["var"] - l["incc"], 2)
        just = norm(l["just"])
        if abs(resto) >= leitura.LIMITE_EAP:
            fora_incc.append((l, resto))
            if not just or just in ("ganho incc", "variacao de incc", "variacao incc", "incc"):
                al.append(f"EAP {l['eap']}: {br(resto)} fora do INCC com justificativa só de INCC ou vazia")
        elif abs(l["var"]) >= 0.01 and not just:
            al.append(f"EAP {l['eap']} variou {br(l['var'])} sem justificativa")
    fora_incc.sort(key=lambda x: -abs(x[1]))
    mudancas = round(var_total - soma_incc, 2)
    par = []
    if abs(var_total) < 0.01:
        par.append("No mês, a análise não apresentou variação no resultado"
                   + (f", que se manteve em {_rs(dif)}." if dif is not None else "."))
    else:
        sentido = "positiva" if var_total > 0 else "negativa"
        frase = f"No mês, a análise apresentou variação {sentido} de {_rs(var_total)}"
        if dif is not None:
            frase += f", com o resultado passando de {_rs(ant)} para {_rs(dif)}"
        par.append(frase + ".")
        par.append(f"O efeito do INCC no mês foi de {_rs(soma_incc)}.")
        if not fora_incc:
            par.append("Não houve mudança de previsão nem pedido novo com efeito relevante no resultado"
                       + (f" (demais itens: {_rs(mudancas)})." if abs(mudancas) >= 0.01 else "."))
        else:
            itens = []
            for l, resto in fora_incc[:6]:
                j = caixa_normal(l["just"])
                itens.append(f"EAP {l['eap']} {l['descricao'] or ''}: {_rs(resto)}" + (f" ({j})" if j else ""))
            citado = sum(r for _, r in fora_incc[:6])
            demais = round(mudancas - citado, 2)
            par.append(f"As mudanças de previsão e pedidos somaram {_rs(mudancas)}:")
            par.append(itens + ([f"Demais itens: {_rs(demais)}"] if abs(demais) >= 0.01 else []))
    fora = [l for l in af["linhas"] if not l["no_total"] and not l["subtotal"] and abs(l["var"]) >= 0.01]
    if fora:
        if nota_fixa:
            par.append(nota_fixa)
        for l in fora:
            par.append(f"Item fora do resultado, EAP {l['eap']} {l['descricao'] or ''}: variação do mês de {_rs(l['var'])}.")
    return par, al, {"variacao": var_total, "incc": soma_incc, "mudancas": mudancas, "diferenca": dif,
                     "diferenca_anterior": ant}


def html_desvios(dv, largura=980):
    def cel(v, fundo, cab=False):
        txt = br(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else html.escape(str(v or ""))
        cor = "#FFFFFF" if cab and fundo else "#1F1F1F"
        alin = "right" if isinstance(v, (int, float)) else "left"
        estilo = (f"border:1px solid #BFBFBF;padding:3px 6px;text-align:{alin};color:{cor};"
                  f"{'font-weight:bold;' if cab else ''}{f'background:{fundo};' if fundo else ''}")
        return f"<td style=\"{estilo}\">{txt}</td>"
    t = [f"<table style=\"border-collapse:collapse;max-width:{largura}px;font-family:Aptos,Calibri,sans-serif;font-size:10pt\">",
         f"<tr><td colspan=\"{len(dv['cabecalho'])}\" style=\"font-weight:bold;padding:4px 6px\">{html.escape(str(dv['titulo']))}</td></tr>",
         "<tr>" + "".join(cel(h["titulo"], h["fundo"], True) for h in dv["cabecalho"]) + "</tr>"]
    for ln in dv["linhas"]:
        t.append("<tr>" + "".join(cel(v, f) for v, f in zip(ln["valores"], ln["fundos"])) + "</tr>")
    f = dv["final"]
    t.append("<tr>" + "".join(cel(v, fd, False).replace("<td style=\"", "<td style=\"font-weight:bold;")
                              for v, fd in zip(f["valores"], f["fundos"])) + "</tr>")
    t.append("</table>")
    return "\n".join(t)


def _paragrafos_html(par):
    out = []
    for p in par:
        if isinstance(p, list):
            out.append("<ul style=\"margin-top:0\">" + "".join(f"<li>{html.escape(x)}</li>" for x in p) + "</ul>")
        else:
            out.append(f"<p>{html.escape(p)}</p>")
    return "\n".join(out)


def montar(cfg_email, planilhas, mes, saudacao=None, abertura=None, assinatura=None):
    """cfg_email: bloco do destinatarios.json; planilhas: {obra: caminho do .xlsm}."""
    if cfg_email.get("pendente"):
        raise ObraParada(f"destinatários de '{cfg_email.get('empreendimento')}' ainda pendentes no destinatarios.json")
    alertas, blocos, conf, anexos = [], [], {}, []
    for obra, path in planilhas.items():
        af = leitura.analise_financeira(path)
        if af["mes"] != mes:
            raise ObraParada(f"{obra}: MÊS ANÁLISE da planilha é {af['mes'] or 'vazio'} ({af['celula_mes']}), "
                             f"o fechamento é {mes}. A revisão do mês não está na pasta.")
        dv = leitura.desvios(path)
        nota = (cfg_email.get("obras", {}).get(obra) or {}).get("nota")
        par, al, nums = resumo_obra(af, nota)
        if af["total"] and dv["total_diferenca"] is not None and af["total"]["dif"] is not None \
                and abs(dv["total_diferenca"] - af["total"]["dif"]) > 1:
            al.append(f"Total da aba Desvios difere ({br(dv['total_diferenca'])} × {br(af['total']['dif'])}): "
                      "coluna lida errada, parar e investigar")
        nome, renomeado = leitura.nome_anexo(path)
        if renomeado:
            al.append(f"anexo renomeado: tirei 'Cópia de' do nome ({nome})")
        anexos.append((path, nome))
        alertas += [f"{obra}: {a}" for a in al]
        blocos.append((obra, par, html_desvios(dv)))
        conf[obra] = {"arquivo": path, "mes": af["mes"], "colunas": af["colunas"], "numeros": nums,
                      "desvios_total": dv["total_diferenca"], "alertas": al}
    para = list(dict.fromkeys(cfg_email.get("para", [])))
    cc = [x for x in dict.fromkeys(cfg_email.get("cc", [])) if x.lower() not in {p.lower() for p in para}]
    externos = [x for x in para + cc if not x.lower().endswith("@" + DOMINIO_INTERNO)]
    if externos:
        alertas.append("destinatários fora da Iben: " + ", ".join(externos) + " (o texto só traz fatos da planilha)")
    empreendimento = cfg_email.get("empreendimento", "")
    assunto = f"{empreendimento} | ANÁLISE CUSTO (FECHAMENTO) - {mes_extenso(mes)}"
    sauda = saudacao or ("bom dia" if dt.datetime.now().hour < 12 else "boa tarde")
    obras_txt = ", ".join(o.title() for o in planilhas) if len(planilhas) > 1 else ""
    ref = f"Segue anexa a análise de custos do {empreendimento.title()}" + (f", {obras_txt}" if obras_txt else "") + \
          f", referente a {mes_extenso(mes).lower()}."
    corpo = [f"<p>Prezados, {html.escape(sauda)}!</p>"]
    if abertura:
        corpo.append(f"<p>{html.escape(abertura)}</p>")
    corpo.append(f"<p>{html.escape(ref)}</p>")
    for obra, par, tab in blocos:
        if len(blocos) > 1:
            corpo.append(f"<p><b>{html.escape(obra)}</b></p>")
        corpo.append(_paragrafos_html(par))
        corpo.append(tab)
    corpo.append("<p>Fico à disposição para esclarecer quaisquer dúvidas.</p>")
    imagens = []
    if assinatura:
        linhas = [html.escape(x) for x in assinatura.get("linhas", [])]
        corpo.append("<div style=\"font-family:Arial,sans-serif;font-size:13.5pt;color:#1F1F1F\">" +
                     "<br>".join(linhas) + "</div>")
        for chave in ("logo", "selos"):
            p = assinatura.get(chave)
            if p and os.path.exists(p):
                cid = make_msgid(domain="iben.eng.br")[1:-1]
                imagens.append((p, cid))
                corpo.append(f"<img src=\"cid:{cid}\" alt=\"{chave}\" style=\"max-height:70px;margin-top:6px\"><br>")
            elif p:
                alertas.append(f"imagem da assinatura não encontrada: {p}")
    html_corpo = ("<html><body style=\"font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:#1F1F1F\">"
                  + "\n".join(corpo) + "</body></html>")
    return {"assunto": assunto, "para": para, "cc": cc, "html": html_corpo, "anexos": anexos, "imagens": imagens,
            "alertas": alertas, "conferencia": conf, "mes": mes}


def salvar(email, pasta):
    """Grava rascunho .eml (abre no Outlook como mensagem nova), previa.html e conferencia.json."""
    os.makedirs(pasta, exist_ok=True)
    msg = EmailMessage()
    msg["Subject"] = email["assunto"]
    if email["para"]:
        msg["To"] = ", ".join(email["para"])
    if email["cc"]:
        msg["Cc"] = ", ".join(email["cc"])
    msg["X-Unsent"] = "1"
    msg.set_content("Abra no Outlook para ver a versão formatada.")
    msg.add_alternative(email["html"], subtype="html")
    corpo_html = msg.get_payload()[1]
    for p, cid in email["imagens"]:
        tipo = (mimetypes.guess_type(p)[0] or "image/png").split("/")
        with open(p, "rb") as f:
            corpo_html.add_related(f.read(), maintype=tipo[0], subtype=tipo[1], cid=f"<{cid}>")
    for p, nome in email["anexos"]:
        with open(p, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=nome)
    base = re.sub(r"[^\w\- ]+", "", email["assunto"]).strip()
    eml = os.path.join(pasta, base + ".eml")
    with open(eml, "wb") as f:
        f.write(bytes(msg))
    prev = os.path.join(pasta, "previa.html")
    cab = (f"<p style=\"font-family:Calibri;color:#555\"><b>Assunto:</b> {html.escape(email['assunto'])}<br>"
           f"<b>Para:</b> {html.escape('; '.join(email['para']))}<br><b>Cc:</b> {html.escape('; '.join(email['cc']))}<br>"
           f"<b>Anexos:</b> {html.escape('; '.join(n for _, n in email['anexos']))}</p><hr>")
    with open(prev, "w", encoding="utf-8") as f:
        f.write(re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + cab, email["html"], count=1))
    conf = os.path.join(pasta, "conferencia.json")
    with open(conf, "w", encoding="utf-8") as f:
        json.dump({"assunto": email["assunto"], "mes": email["mes"], "alertas": email["alertas"],
                   "obras": email["conferencia"]}, f, ensure_ascii=False, indent=1, default=str)
    return {"eml": eml, "previa": prev, "conferencia": conf}
