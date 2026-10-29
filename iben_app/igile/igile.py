#!/usr/bin/env python3
"""Atualizacao mensal do iGile (Iben): extracao, ajustes, JavaScript do navegador e log.

Versao 2.1 (29/10/2026). Correcao do Pedro:
  - o periodo do iGile e o mes que esta NO NOME da planilha mais recente da pasta da obra,
    na maior revisao (comando `ultima`). Nada de calcular "mes anterior ao atual";
  - o Incorrido da planilha do mes M e sempre N-1: o custo real do periodo M sai da propria
    planilha M, somando MES = M-1 (a planilha de outubro traz incorrido ate setembro);
  - confere o MES ANALISE da aba Analise Financeira contra o nome (vale o nome, sai alerta);
  - Curva Fisica com duas obras no mesmo arquivo (Kronos): aba "Avanco Custo - <obra>",
    secao da obra na aba e-mail e, se a aba e-mail estiver parada, bloco da obra no Comparativo.
Versao 2.0 (28/10/2026): trava de layout, Curva Fisica da Altus, avanco exato, ajuste --json,
  js-nivel8 e js-valor.

Comandos:
  ultima      python3 -I igile.py ultima --tipo custo "<nome 1>" "<nome 2>" ...
              (escolhe o ultimo mes pelo nome e a maior revisao; sai com codigo 1 se houver empate)
  extrair (padrao)
    python3 -I igile.py --obra "Viva Parque Unique" --mes 2026-10 \
        --custo-mes <planilha do mes>.xlsm [--curva CURVA.xlsx] \
        [--pls PLS.xlsx] [--efetivo EFETIVO.xlsx | --efetivo-manual 73] --saida obra.json
  ajuste
    python3 -I igile.py ajuste --etapa previsto --json obra.json --gerado -5697157.39 \
        --estoque-atual X --adiantamento-atual Y [--agregado-atual Z] [--estoque-fixo V]
    python3 -I igile.py ajuste --etapa previsao --json obra.json --gerado 8072192.20 --previsao-atual Z
  js-nivel8   python3 -I igile.py js-nivel8 obra.json --modo ler|aplicar [--salvar]
  js-valor    python3 -I igile.py js-valor 3043351,44 [--periodo 16] [--salvar]
  log         python3 -I igile.py log dados/*.json --saida "Log iGile 2026-09.xlsx"
  b64         python3 -I igile.py b64 <tool-result salvo> <pasta>

Cada valor sai com a origem (arquivo, aba, celula/regra) e alertas.
Nunca inventa valor: campo sem fonte sai como null com o motivo.
"""
import argparse, datetime as dt, json, os, re, sys, unicodedata, warnings
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.utils import get_column_letter as L

VERSAO = "2.4"

# campos que cada arquivo tem que entregar; se o arquivo foi informado e o campo
# saiu vazio, e mudanca de layout: o script para (codigo 2) em vez de seguir.
OBRIGATORIOS = {
    "custo_mes": ["contingencia", "diferenca", "retencao", "a_pagar", "custo_real"],
    "curva": ["medicao", "avanco_fisico_acumulado", "medicao_exclusiva_custo"],
    "pls": ["realizado_banco"],
    "efetivo": ["realizado_recursos"],
}


def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def sheet(wb, *nomes, exato=False):
    alvo = [norm(n) for n in nomes]
    for ws in wb.worksheets:
        t = norm(ws.title)
        if (t in alvo) if exato else any(a == t or (a in t and len(t) - len(a) <= 2) for a in alvo):
            return ws
    raise KeyError(f"aba nao encontrada: {nomes} (abas: {[w.title for w in wb.worksheets]})")


def mes_de(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m")
    return None


def mes_anterior(mes):
    a, m = int(mes[:4]), int(mes[5:7])
    return f"{a - 1}-12" if m == 1 else f"{a}-{m - 1:02d}"


# ---------------- escolha do arquivo pelo nome (regra do Pedro, 29/10/2026) ----------------

_MESES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
          "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
          "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8,
          "set": 9, "out": 10, "nov": 11, "dez": 12}
_RX_NOME = re.compile(r"(?<![a-z])(" + "|".join(sorted(_MESES, key=len, reverse=True)) +
                      r")[\s_\-\.]*(20\d{2}|\d{2})(?!\d)")
_TIPOS = {"custo": "analise de custos", "curva": "curva fisica", "pls": "pls", "efetivo": "efetivo"}


def _ano(s):
    a = int(s)
    return a if a > 100 else 2000 + a


def mes_do_nome(nome):
    """AAAA-MM pelo nome do arquivo. Ordem: nome do mes + ano (Outubro26, Out_26, Set26,
    Setembro 2026); depois MM.AAAA / MM-AAAA (ex. EFETIVO 22.09.2026); depois codigo de 4
    digitos, MMAA (1026) ou AAMM (2609). Retorna (mes, alerta)."""
    s = norm(os.path.splitext(os.path.basename(nome))[0])
    por_nome = [f"{_ano(a)}-{_MESES[m]:02d}" for m, a in _RX_NOME.findall(s)]
    por_data = [f"{a}-{m}" for m, a in re.findall(r"(?<!\d)(0[1-9]|1[0-2])[\.\-_/](20\d{2})(?!\d)", s)]
    por_cod = []
    for x in re.findall(r"(?<![\d.,])(\d{4})(?![\d.,])", s):
        p, q = int(x[:2]), int(x[2:])
        if 1 <= p <= 12 and 20 <= q <= 40:
            por_cod.append(f"20{q:02d}-{p:02d}")
        elif 20 <= p <= 40 and 1 <= q <= 12:
            por_cod.append(f"20{p:02d}-{q:02d}")
    achados = por_nome or por_data or por_cod
    if not achados:
        return None, "sem mês no nome"
    todos = set(por_nome + por_data + por_cod)
    alerta = f"nome com mais de um mês ({', '.join(sorted(todos))}); usei {achados[-1]}" if len(todos) > 1 else None
    return achados[-1], alerta


def revisao_do_nome(nome):
    s = norm(os.path.splitext(os.path.basename(nome))[0])
    r = re.findall(r"(?<![a-z])r(?:ev)?[\s_\-\.]*(\d{1,2})(?!\d)", s)
    return int(r[-1]) if r else 0


def escolher(nomes, tipo=None):
    """Ultimo mes pelo nome e, dentro dele, a maior revisao. Retorna dict com a escolha,
    a tabela de candidatos e alertas (empate = bloqueia)."""
    filtro = _TIPOS.get(tipo)
    cands = []
    for n in nomes:
        base = os.path.basename(n)
        if filtro and filtro not in norm(base):
            continue
        if base.startswith("~$"):
            continue
        mes, al = mes_do_nome(base)
        cands.append({"arquivo": n, "mes": mes, "revisao": revisao_do_nome(base), "alerta": al,
                      "copia": norm(base).startswith("copia de")})
    validos = [c for c in cands if c["mes"]]
    if not validos:
        return {"escolha": None, "candidatos": cands, "alerta": "nenhum arquivo com mês no nome"}
    topo = max((c["mes"], c["revisao"]) for c in validos)
    iguais = [c for c in validos if (c["mes"], c["revisao"]) == topo]
    out = {"escolha": iguais[0], "candidatos": sorted(cands, key=lambda c: (c["mes"] or "", c["revisao"]), reverse=True)}
    if len(iguais) > 1:
        out["escolha"] = None
        out["alerta"] = "empate no mesmo mês e revisão: " + " | ".join(os.path.basename(c["arquivo"]) for c in iguais)
    return out


def campo(valor, origem, alerta=None, **extra):
    d = {"valor": valor, "origem": origem}
    if alerta:
        d["alerta"] = alerta
    d.update(extra)
    return d


def achar_header(ws, rotulos, max_row=6):
    """Retorna (linha, {rotulo_normalizado: coluna}) da primeira linha que contem algum rotulo."""
    for r in range(1, max_row + 1):
        cols = {}
        for c in range(1, ws.max_column + 1):
            t = norm(ws.cell(r, c).value)
            if t:
                cols.setdefault(t, c)
        if any(norm(x) in cols for x in rotulos):
            return r, cols
    return None, {}


def soma_coluna(ws, col, r0):
    tot = 0.0
    for (v,) in ws.iter_rows(min_row=r0, min_col=col, max_col=col, values_only=True):
        if num(v):
            tot += v
    return tot


def total_coluna(ws, hdr_row, col, nome_arq, aba_rotulo):
    proprio = soma_coluna(ws, col, hdr_row + 1)
    cache = ws.cell(1, col).value
    ref = f"{nome_arq} | {ws.title} | coluna {L(col)} ({aba_rotulo})"
    if num(cache) and abs(cache - proprio) > 0.01:
        return campo(round(proprio, 2), ref,
                     f"total da linha 1 ({cache:,.2f}) difere da soma da coluna ({proprio:,.2f}); "
                     "provavel filtro salvo no arquivo. Usei a soma completa.")
    return campo(round(proprio, 2), ref)


# ---------------- planilha de analise de custos ----------------

def custo_mes(path, mes):
    nome = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}

    # Contingencia atual: aba Memoria de Calculo - Previsao, rotulo ATUAL
    try:
        ws = sheet(wb, "Memória de Cálculo - Previsão")
        v = None
        for r in range(1, 15):
            for c in range(1, 8):
                if norm(ws.cell(r, c).value) == "atual":
                    for c2 in range(c + 1, c + 4):
                        if num(ws.cell(r, c2).value):
                            v, cel = ws.cell(r, c2).value, f"{L(c2)}{r}"
                            break
                if v is not None:
                    break
            if v is not None:
                break
        out["contingencia"] = campo(round(v, 2), f"{nome} | {ws.title} | {cel} (Contingência ATUAL)") if v is not None \
            else campo(None, f"{nome} | {ws.title}", "rótulo ATUAL não encontrado")
    except KeyError as e:
        out["contingencia"] = campo(None, nome, str(e))

    # Diferenca (resultado final da analise da PREVISAO): aba Desvios
    try:
        ws = sheet(wb, "Desvios", exato=True)
        sec = None
        for r in range(1, ws.max_row + 1):
            t = " ".join(norm(ws.cell(r, c).value) for c in range(1, 5))
            if "resultado da analise do previsao" in t:
                sec = r
                break
        dif = None
        if sec:
            col_dif = None
            for c in range(1, 12):
                if norm(ws.cell(sec + 1, c).value) == "diferenca":
                    col_dif = c
            for r in range(sec + 1, ws.max_row + 1):
                if norm(ws.cell(r, 2).value).startswith("resultado final"):
                    if col_dif and num(ws.cell(r, col_dif).value):
                        dif, cel = ws.cell(r, col_dif).value, f"{L(col_dif)}{r}"
                    break
        out["diferenca"] = campo(round(dif, 2), f"{nome} | Desvios | {cel} (Resultado final da previsão, coluna DIFERENÇA)") \
            if dif is not None else campo(None, f"{nome} | Desvios", "linha RESULTADO FINAL não encontrada")
    except KeyError as e:
        out["diferenca"] = campo(None, nome, str(e))

    # Retencao atual (saldo ainda retido)
    try:
        ws = sheet(wb, "Retenções")
        rotulos = ["RETENÇÃO ATUAL", "SALDO RETIDO", "VALOR RETENÇÃO"]
        hr, cols = achar_header(ws, rotulos)
        rot = next((x for x in rotulos if norm(x) in cols), None)
        col = cols.get(norm(rot)) if rot else None
        out["retencao"] = total_coluna(ws, hr, col, nome, rot) if col else campo(None, f"{nome} | {ws.title}", "coluna de saldo retido não encontrada")
        if col and rot == "VALOR RETENÇÃO" and not out["retencao"].get("alerta"):
            # layout Kronos: so o valor retido por NF, sem coluna de saldo/devolucao
            out["retencao"]["alerta"] = "layout Kronos: soma de VALOR RETENÇÃO (sem coluna de saldo; conferir se houve devolução)"
    except KeyError as e:
        out["retencao"] = campo(None, nome, str(e))

    # A pagar: total da coluna VALOR BAIXADO (ou VALOR)
    try:
        ws = sheet(wb, "A pagar")
        hr, cols = achar_header(ws, ["VALOR BAIXADO", "VALOR"])
        col = cols.get(norm("VALOR BAIXADO")) or cols.get("valor")
        rot = "VALOR BAIXADO" if norm("VALOR BAIXADO") in cols else "VALOR"
        out["a_pagar"] = total_coluna(ws, hr, col, nome, rot) if col else campo(None, f"{nome} | {ws.title}", "coluna de valor não encontrada")
        fo = cols.get(norm("CUSTO QUE NÃO É DA OBRA")) or cols.get(norm("CUSTO FORA OBRA"))
        if fo and out["a_pagar"]["valor"] is not None:
            out["a_pagar"]["custo_fora_obra_info"] = round(soma_coluna(ws, fo, hr + 1), 2)
    except KeyError as e:
        out["a_pagar"] = campo(None, nome, str(e))

    # MES ANALISE (aba Analise Financeira): so confere com o mes do nome. Vale o nome.
    try:
        ws = sheet(wb, "Análise Financeira")
        ma = cel = None
        for r in range(1, 12):
            for c in range(1, 16):
                if norm(ws.cell(r, c).value) in ("mes analise", "mes de analise"):
                    for rr, cc in ((r + 1, c), (r, c + 1)):
                        ma = mes_de(ws.cell(rr, cc).value)
                        if ma:
                            cel = f"{L(cc)}{rr}"
                            break
                if ma:
                    break
            if ma:
                break
        if ma:
            out["mes_analise"] = campo(ma, f"{nome} | {ws.title} | {cel} (MÊS ANÁLISE)",
                                       None if ma == mes else f"MÊS ANÁLISE da planilha é {ma}, o nome diz {mes} (vale o nome)")
    except KeyError:
        pass
    return out


def incorrido_do_mes(path, mes):
    """Soma do Incorrido com MES = mes. Desde a v2.1 o extrair chama com a propria planilha do
    periodo M e mes = M-1 (o Incorrido da analise e sempre N-1)."""
    nome = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = sheet(wb, "Incorrido", exato=True)
    except KeyError as e:
        return campo(None, nome, str(e))
    rows = ws.iter_rows(values_only=True)
    hdr = None
    for row in rows:
        if any(norm(v) == "mes" for v in row):
            hdr = [norm(v) for v in row]
            break
    if hdr is None:
        return campo(None, f"{nome} | Incorrido", "cabeçalho com coluna MÊS não encontrado")
    c_mes = hdr.index("mes")
    if "valor baixado" in hdr:
        c_val = hdr.index("valor baixado")
    elif "valor" in hdr:
        c_val = hdr.index("valor")
    else:
        return campo(None, f"{nome} | Incorrido", "coluna VALOR BAIXADO / VALOR não encontrada")
    c_fo = next((i for i, h in enumerate(hdr) if "nao e da obra" in h or "fora obra" in h), None)
    c_deb = hdr.index("debito") if "debito" in hdr else None
    val = fo = deb = 0.0
    n = 0
    depois = {}
    for row in rows:
        if len(row) > c_mes and mes_de(row[c_mes]) and mes_de(row[c_mes]) > mes:
            depois[mes_de(row[c_mes])] = depois.get(mes_de(row[c_mes]), 0) + 1
        if len(row) <= c_val or mes_de(row[c_mes]) != mes:
            continue
        n += 1
        val += row[c_val] if num(row[c_val]) else 0
        if c_fo is not None and num(row[c_fo]):
            fo += row[c_fo]
        if c_deb is not None and num(row[c_deb]):
            deb += row[c_deb]
    if n == 0:
        return campo(None, f"{nome} | Incorrido", f"nenhum lançamento com MÊS = {mes}")
    al = []
    if depois:
        al.append("Incorrido tem lançamentos depois do N-1, não somados: " +
                  ", ".join(f"{k} ({v} linha{'s' if v > 1 else ''})" for k, v in sorted(depois.items())))
    return campo(round(val - fo, 2),
                 f"{nome} | Incorrido | MÊS = {mes} (N-1): soma de '{hdr[c_val].upper()}' ({n} linhas) menos 'custo fora da obra'",
                 "; ".join(al) or None,
                 componentes={"valor": round(val, 2), "custo_fora_obra": round(fo, 2),
                              "debito": round(deb, 2) if c_deb is not None else None, "linhas": n,
                              "mes_incorrido": mes, "depois_do_n1": depois})


# ---------------- curva fisica ----------------

def _secao_email(ws, sufixo):
    """Linha onde comeca a secao da obra na aba e-mail (ex. 'LOTEAMENTO:' no Kronos). 1 se nao houver."""
    if sufixo:
        for r in range(1, ws.max_row + 1):
            if norm(ws.cell(r, 1).value).rstrip(":").strip() == sufixo:
                return r
    return 1


def _mes_aba_email(ws, r0=1):
    """Mes da aba e-mail: data na linha 'Mes Atual de Analise' (Unique: H1; Altus: H2)."""
    for r in range(r0, r0 + 6):
        if norm(ws.cell(r, 1).value).startswith("mes atual de analise"):
            for c in range(2, 12):
                m = mes_de(ws.cell(r, c).value)
                if m:
                    return m
    if r0 == 1:
        for r in range(1, 4):
            m = mes_de(ws.cell(r, 8).value)
            if m:
                return m
    return None


def _comparativo(wb, obra, mes):
    """Bloco da obra na aba Comparativo (Kronos): linhas REALIZADO > Evolucao Fisica / Fisico Acumulado,
    coluna do mes. Usado so quando a aba e-mail nao esta no mes."""
    if not obra:
        return None
    try:
        ws = sheet(wb, "Comparativo", exato=True)
    except KeyError:
        return None
    alvo = norm(obra)
    ini = c_lab = None
    for r in range(1, ws.max_row + 1):
        c_lab = next((c for c in (1, 2, 3) if norm(ws.cell(r, c).value) == alvo), None)
        if c_lab:
            ini = r
            break
    if not ini:
        return None
    col = None
    for r in range(ini, min(ini + 4, ws.max_row) + 1):
        col = next((c for c in range(1, ws.max_column + 1) if mes_de(ws.cell(r, c).value) == mes), None)
        if col:
            break
    rr = next((r for r in range(ini + 1, min(ini + 12, ws.max_row) + 1) if norm(ws.cell(r, c_lab).value) == "realizado"), None)
    if not col or not rr:
        return None
    per = acum = None
    for r in range(rr + 1, rr + 4):
        t = norm(ws.cell(r, c_lab).value)
        if t.startswith("evolucao fisica"):
            per = (ws.cell(r, col).value, f"{L(col)}{r}")
        elif t.startswith("fisico acumulado"):
            acum = (ws.cell(r, col).value, f"{L(col)}{r}")
    if per and acum and num(per[0]) and num(acum[0]):
        return {"per": per, "acum": acum, "aba": ws.title}
    return None


def _layout_avanco(ws, mes):
    """Acha linha de cabecalho, coluna do mes e colunas de descricao/EAP na aba Avanco Custo.
    Unique e Vila Flora: datas na linha 1, EAP em A, descricao em B, dados a partir da linha 2.
    Altus: 'MES n' na linha 1, datas e rotulos na linha 2, EAP em B, descricao em E."""
    hr = col = None
    for r in range(1, 5):
        for c in range(1, ws.max_column + 1):
            if mes_de(ws.cell(r, c).value) == mes:
                hr, col = r, c
                break
        if col:
            break
    if not col:
        raise KeyError(f"coluna do mês {mes} não encontrada nas linhas 1 a 4 da aba Avanço Custo")
    rot = {}
    for r in range(1, hr + 1):
        for c in range(1, min(ws.max_column, 20) + 1):
            t = norm(ws.cell(r, c).value)
            if t:
                rot.setdefault(t, c)
    c_desc = next((c for t, c in rot.items() if t.startswith("discriminacao")), 2)
    c_eap = next((c for t, c in rot.items() if t.startswith("codigo eap")), 1)
    c_niv = next((c for t, c in rot.items() if t.startswith("nivel eap")), None)
    return hr, col, c_desc, c_eap, c_niv


def curva(path, mes, obra=None):
    nome = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    sufixo = norm(obra.split(" - ")[-1]) if obra and " - " in obra else None
    # resumo da aba e-mail (oculta): % realizado no periodo e acumulado (secao da obra, se houver)
    try:
        ws = sheet(wb, "e-mail", exato=True)
        r0 = _secao_email(ws, sufixo)
        mes_aba = _mes_aba_email(ws, r0)
        achados = {}
        for r in range(r0, ws.max_row + 1):
            t = norm(ws.cell(r, 1).value)
            if t.startswith("% realizado acumulado") and "acum" not in achados:
                achados["acum"] = (ws.cell(r, 2).value, f"B{r}")
            elif t.startswith("% realizado no periodo") and "per" not in achados:
                achados["per"] = (ws.cell(r, 2).value, f"B{r}")
        aba = "e-mail" + (f" (seção {ws.cell(r0, 1).value})" if r0 > 1 else "")
        al = None if mes_aba == mes else f"aba e-mail está no mês {mes_aba}, não em {mes}"
        if mes_aba != mes:
            comp = _comparativo(wb, obra, mes)
            if comp:
                achados = {"per": comp["per"], "acum": comp["acum"]}
                aba = f"{comp['aba']} (bloco {obra}, REALIZADO)"
                al = f"aba e-mail parada em {mes_aba}; usei o bloco REALIZADO da obra no Comparativo (conferir com o iGile)"
        if "per" in achados and num(achados["per"][0]):
            v, cel = achados["per"]
            out["medicao"] = campo(round(v * 100, 2), f"{nome} | {aba} | {cel} (% realizado no período)", al, unidade="%")
        else:
            out["medicao"] = campo(None, f"{nome} | e-mail", "rótulo '% REALIZADO NO PERÍODO' sem valor numérico")
        if "acum" in achados and num(achados["acum"][0]):
            v, cel = achados["acum"]
            out["avanco_fisico_acumulado"] = campo(round(v, 4), f"{nome} | {aba} | {cel} (% realizado acumulado)", al, exato=v)
        else:
            out["avanco_fisico_acumulado"] = campo(None, f"{nome} | e-mail", "rótulo '% REALIZADO ACUMULADO' sem valor numérico")
    except KeyError as e:
        out["medicao"] = campo(None, nome, str(e))
        out["avanco_fisico_acumulado"] = campo(None, nome, str(e))

    # Avanco Custo: coluna do mes, da primeira linha com descricao ate a ultima -> nivel 8
    try:
        avs = [w for w in wb.worksheets if norm(w.title).startswith("avanco custo")]
        if sufixo and len(avs) > 1:
            ws = next((w for w in avs if norm(w.title).endswith(sufixo)), None)
            if ws is None:
                raise KeyError(f"nenhuma aba Avanço Custo da obra '{sufixo}' entre {[w.title for w in avs]}")
        else:
            ws = sheet(wb, "Avanço Custo")
        hr, col, c_desc, c_eap, c_niv = _layout_avanco(ws, mes)
        com_desc = [r for r in range(hr + 1, ws.max_row + 1) if ws.cell(r, c_desc).value not in (None, "")]
        if not com_desc:
            raise KeyError("nenhuma linha com descrição abaixo do cabeçalho da aba Avanço Custo")
        ini, ult = com_desc[0], com_desc[-1]
        itens = []
        # linhas vazias no meio entram: o iGile casa pela posicao (ex.: Vila Flora, linha 43)
        for r in range(ini, ult + 1):
            v = ws.cell(r, col).value
            d = ws.cell(r, c_desc).value
            itens.append({"linha": r, "eap": ws.cell(r, c_eap).value,
                          "nivel": ws.cell(r, c_niv).value if c_niv else None,
                          "descricao": str(d).strip() if d is not None else None,
                          "valor_pct": round(v * 100, 2) if num(v) else None})
        tot = next((ws.cell(r, col).value for r in range(ult + 1, ws.max_row + 1) if num(ws.cell(r, col).value)), None)
        al = None if num(tot) else "total abaixo da última linha não encontrado"
        if not num(tot):
            # Kronos: sem linha de total. O iGile pondera pelo custo: total = soma(custo x %) / soma(custo)
            c_custo = next((c for r in range(1, hr + 1) for c in range(1, min(ws.max_column, 12) + 1)
                            if norm(ws.cell(r, c).value) == "custo"), None)
            if c_custo:
                pares = [(ws.cell(r, c_custo).value, ws.cell(r, col).value) for r in range(ini, ult + 1)]
                pares = [(c, v) for c, v in pares if num(c) and num(v)]
                den = sum(c for c, _ in pares)
                if den:
                    tot = sum(c * v for c, v in pares) / den
                    al = (f"sem linha de total: calculei ponderado pelo custo (coluna {L(c_custo)}), "
                          "conferir com o nível 1 do iGile")
        out["medicao_exclusiva_custo"] = campo(
            round(tot * 100, 2) if num(tot) else None,
            f"{nome} | {ws.title} | {L(col)}{ini}:{L(col)}{ult} ({len(itens)} linhas, colar no nível 8)",
            al, unidade="%", itens=itens)
    except KeyError as e:
        out["medicao_exclusiva_custo"] = campo(None, nome, str(e))
    return out


# ---------------- PLS (realizado banco) ----------------

def pls(path):
    nome = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = sheet(wb, "PLS", exato=True)
    except KeyError as e:
        return campo(None, nome, str(e))
    etapa = None
    for row in ws.iter_rows(min_row=1, max_row=12):
        for c in row:
            if norm(c.value).startswith("etapa"):
                for k in range(1, 4):
                    v = ws.cell(c.row, c.column + k).value
                    if v not in (None, ""):
                        etapa = v
                        break
    for row in ws.iter_rows():
        for c in row:
            if norm(c.value).startswith("no periodo"):
                cand = [ws.cell(c.row, c.column + k) for k in range(1, 6)] + \
                       [ws.cell(c.row + d, c.column + k) for d in (1, 2) for k in range(0, 4)]
                for x in cand:
                    if num(x.value) and 0 < x.value < 1:
                        return campo(round(x.value * 100, 2), f"{nome} | PLS | {x.coordinate} (No período)",
                                     "usar só a PLS aceita pela Caixa", unidade="%", etapa=etapa)
    return campo(None, f"{nome} | PLS", "rótulo 'No período' com percentual não encontrado")


# ---------------- efetivo ----------------

def efetivo(path, mes):
    """Media dos dias com apontamento, sem domingos. Procura linha TOTAL e colunas de dias."""
    nome = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ano, m = map(int, mes.split("-"))
    ndias = (dt.date(ano + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).day
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if norm(c.value) == "total":
                    dias = [ws.cell(c.row, col).value for col in range(c.column + 1, ws.max_column + 1)]
                    seq = [v if num(v) else 0 for v in dias[:ndias]]
                    if sum(1 for v in seq if v) < 5:
                        continue
                    usados = [v for d, v in enumerate(seq, 1)
                              if dt.date(ano, m, d).weekday() != 6 and v > 0]
                    media = sum(usados) / len(usados)
                    return campo(int(media), f"{nome} | {ws.title} | linha {c.row} (TOTAL), {len(usados)} dias usados",
                                 "conferir: assume que a 1ª coluna após TOTAL é o dia 1",
                                 media_exata=round(media, 2), media_com_domingos=round(sum(seq) / ndias, 2))
    return campo(None, nome, "linha TOTAL com dias do mês não encontrada")


# ---------------- trava de layout ----------------

def validar(r, informados):
    """Lista as paradas: arquivo informado cujo campo obrigatorio saiu vazio, ou curva de outro mes."""
    paradas = []
    for fonte, campos in OBRIGATORIOS.items():
        if fonte not in informados:
            continue
        for k in campos:
            v = r["campos"].get(k)
            if v is None or v.get("valor") is None:
                paradas.append(f"{k}: {(v or {}).get('alerta', 'sem leitura')} [{(v or {}).get('origem', fonte)}]")
            elif "não em" in (v.get("alerta") or ""):
                paradas.append(f"{k}: {v['alerta']}")
    return paradas


# ---------------- ajustes de equilibrio ----------------

def ajustes(a):
    """Formulas da planilha auxiliar do Raphael (PLANILHA_IGILE_OBRAS), em modo delta.
    Previsto: alvo = (diferenca + contingencia) x avanco acumulado; delta = alvo - gerado.
      delta >= 0 (melhorar): soma em Estoque e Adiantamento (padrao metade/metade).
      delta < 0 (piorar): soma |delta| em Agregado nao Pago; Estoque e Adiantamento ficam como estao.
    Previsao: delta = gerado - diferenca; Previsao de Despesas nova = atual + delta.
    |delta| < tolerancia: nada a lancar (OPR ja bate)."""
    if a.json:
        d = json.load(open(a.json, encoding="utf-8"))["campos"]
        if a.diferenca is None:
            a.diferenca = d["diferenca"]["valor"]
        if a.contingencia is None:
            a.contingencia = d["contingencia"]["valor"]
        if a.acumulado is None:
            av = d.get("avanco_fisico_acumulado", {})
            a.acumulado = av.get("exato", av.get("valor"))
    faltam = [n for n in ("diferenca", "gerado") if getattr(a, n) is None]
    if a.etapa == "previsto":
        faltam += [n for n in ("contingencia", "acumulado") if getattr(a, n) is None]
    if faltam:
        sys.exit(f"PARADA: faltam {faltam} (informe direto ou via --json)")
    tol = 1000.0
    if a.etapa == "previsto":
        alvo = (a.diferenca + a.contingencia) * a.acumulado
        delta = alvo - a.gerado
        r = {"alvo_variacao_previsto": round(alvo, 2), "gerado": a.gerado, "delta": round(delta, 2),
             "entradas": {"diferenca": a.diferenca, "contingencia": a.contingencia, "acumulado": a.acumulado}}
        if abs(delta) < tol:
            r["acao"] = "nada a lançar: OPR previsto já bate com o alvo"
        elif delta > 0:
            p_est = delta * a.parcela_estoque if a.estoque_fixo is None else a.estoque_fixo
            r.update(acao="melhorar: somar em Estoque e Adiantamento",
                     estoque_novo=round(a.estoque_atual + p_est, 2),
                     adiantamento_novo=round(a.adiantamento_atual + delta - p_est, 2))
        else:
            r.update(acao="piorar: somar em Agregado não Pago",
                     agregado_nao_pago_novo=round(a.agregado_atual - delta, 2))
    else:
        delta = a.gerado - a.diferenca
        r = {"alvo_variacao_previsao": a.diferenca, "gerado": a.gerado, "delta": round(delta, 2)}
        if abs(delta) < tol:
            r["acao"] = "nada a lançar: OPR previsão já bate com a diferença"
        else:
            r.update(acao="somar delta em Previsão de Despesas",
                     previsao_despesas_nova=round(a.previsao_atual + delta, 2))
    print(json.dumps(r, ensure_ascii=False, indent=1))
    return r


# ---------------- JavaScript para o iGile ----------------

JS_SET = ("const _set=(i,v)=>{const s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;"
          "i.focus();s.call(i,v);['input','change','keyup','blur'].forEach(t=>i.dispatchEvent(new Event(t,{bubbles:true})));};")
JS_NUM = "const _n=s=>{s=String(s||'').trim();if(!s)return null;if(s.includes(','))s=s.replace(/\\./g,'').replace(',','.');const x=parseFloat(s);return isNaN(x)?null:x;};"
JS_NORM = "const _t=s=>String(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase().replace(/\\s+/g,' ').trim();"
JS_ROWS = "const _rows=()=>[...document.querySelectorAll('tr')].filter(tr=>[...tr.querySelectorAll('input[type=text]')].some(i=>i.offsetParent));"
# v2.1 (29/10/2026, validado no iGile real): para gravar, o iGile exige Calcular ANTES de Salvar.
# Sequencia: preencher o campo -> botao Calcular (#botaoCalcular, no painel lateral que abre pelo
# icone no canto direito, acima de NIVEL) -> Salvar do painel (#salvar) -> aviso "Operacao realizada
# com sucesso" -> ao Voltar, o icone da tela fica branco (btn-white) em vez de vermelho (btn-danger).
# A pagina tem outros "Salvar" ocultos (formSettings / buttonChangeApiKey, configuracoes da conta):
# NUNCA clicar neles. So #botaoCalcular e #salvar, pelo id.
JS_GRAVAR = ("const _sleep=ms=>new Promise(r=>setTimeout(r,ms));"
             "const _okToasts=()=>[...document.querySelectorAll('#toast-container .toast')].filter(t=>/sucesso/i.test(t.innerText)).length;"
             "const _gravar=async()=>{const c=document.getElementById('botaoCalcular'),s=document.getElementById('salvar');"
             "if(!c||!s||c.closest('form')||s.closest('form'))return 'PARADA: Calcular/Salvar do painel não encontrados';"
             "if(!c.offsetParent){const tg=document.getElementById('toggleAdvanced');if(tg){tg.click();await _sleep(1200);}}"
             "if(c.hasAttribute('disabled')){const i=[...document.querySelectorAll('tr input[type=text]')].find(x=>x.offsetParent);"
             "if(i){['keyup','input','change','blur'].forEach(t=>i.dispatchEvent(new Event(t,{bubbles:true})));await _sleep(800);}}"
             "if(c.hasAttribute('disabled'))return 'PARADA: Calcular desabilitado (painel '+(c.offsetParent?'aberto':'fechado')+')';"
             "c.click();await _sleep(2000);"
             "if(s.hasAttribute('disabled'))return 'PARADA: Salvar continua desabilitado depois do Calcular';"
             "const n0=_okToasts();s.click();"
             "for(let k=0;k<16;k++){await _sleep(500);"
             "const er=[...document.querySelectorAll('#toast-container .toast-error')].map(t=>t.innerText.trim()).join(' ');"
             "if(er)return 'PARADA: iGile respondeu erro: '+er.slice(0,120);"
             "if(_okToasts()>n0)return 'GRAVADO (Operação realizada com sucesso)';}"
             "return 'PARADA: sem aviso de sucesso em 8 s (reabrir a tela e conferir)';};")


def br(v, casas=2):
    return f"{v:.{casas}f}".replace(".", ",")


def js_nivel8(a):
    """Gera o JS do nivel 8. modo ler: compara iGile x planilha e devolve so as divergencias.
    modo aplicar: confere de novo e preenche as linhas com valor (as de grupo o iGile calcula)."""
    d = json.load(open(a.json, encoding="utf-8"))
    c = d["campos"].get("medicao_exclusiva_custo") or {}
    itens = c.get("itens")
    if not itens:
        sys.exit("PARADA: JSON sem itens de medicao_exclusiva_custo")
    esp = [[norm(i.get("descricao")), br(i["valor_pct"]) if i.get("valor_pct") is not None else None] for i in itens]
    aplicar = "true" if a.modo == "aplicar" else "false"
    salvar = "true" if (a.salvar and a.modo == "aplicar") else "false"
    js = ("await (async()=>{" + JS_SET + JS_NUM + JS_NORM + JS_ROWS + JS_GRAVAR +
          f"const E={json.dumps(esp, ensure_ascii=False)};const APLICAR={aplicar},SALVAR={salvar},MAXNOME={a.max_divergencias};"
          "const R=_rows();"
          "if(R.length!==E.length)return 'PARADA: iGile tem '+R.length+' linhas no nivel 8 e a planilha tem '+E.length;"
          "if(APLICAR&&R.some(tr=>[...tr.querySelectorAll('input[type=text]')].some(i=>i.offsetParent&&i.disabled)))return 'PARADA: período fechado no iGile (campos desabilitados)';"
          "const nomes=[],valores=[],multi=[];"
          "R.forEach((tr,k)=>{const ins=[...tr.querySelectorAll('input[type=text]')].filter(i=>i.offsetParent);"
          "if(ins.length!==1){multi.push(k+1);return;}"
          "const txt=_t(tr.innerText);if(E[k][0]&&!txt.includes(E[k][0]))nomes.push((k+1)+': '+tr.innerText.trim().replace(/\\s+/g,' ').slice(0,45)+' x '+E[k][0]);"
          "if(E[k][1]!==null){const at=_n(ins[0].value),nv=_n(E[k][1]);if(at===null||Math.abs(at-nv)>0.005)valores.push((k+1)+': '+(ins[0].value||'vazio')+' -> '+E[k][1]);}});"
          "if(multi.length)return 'PARADA: linhas com mais de um campo: '+multi.slice(0,10).join(',');"
          "if(nomes.length>MAXNOME)return 'PARADA: '+nomes.length+' nomes diferentes: '+nomes.slice(0,6).join(' | ');"
          "if(!APLICAR)return 'LEITURA OK '+R.length+' linhas | nomes diferentes: '+(nomes.join(' | ')||'nenhum')+' | valores diferentes: '+valores.length+(valores.length?' ('+valores.slice(0,15).join('; ')+')':'');"
          "let n=0;R.forEach((tr,k)=>{if(E[k][1]===null)return;const i=[...tr.querySelectorAll('input[type=text]')].find(x=>x.offsetParent);_set(i,E[k][1]);n++;});"
          "const g=SALVAR?await _gravar():'sem salvar';"
          "return 'PREENCHIDO '+n+' campos | nomes diferentes: '+(nomes.join(' | ')||'nenhum')+' | '+g;"
          "})()")
    if a.saida:
        open(a.saida, "w", encoding="utf-8").write(js)
        print(a.saida)
    else:
        print(js)
    return js


def _js_valor(v, per=None, salvar=False):
    """JS de uma tela: preenche o unico campo visivel (ou da linha do periodo N) e, com salvar,
    faz Calcular + Salvar do painel e espera o aviso de sucesso."""
    v = str(v).replace(" ", "")
    if "." in v and "," not in v:
        v = v.replace(".", ",")
    p = "null" if per is None else str(int(per))
    return ("await (async()=>{" + JS_SET + JS_ROWS + JS_GRAVAR +
            f"const V={json.dumps(v)},P={p},SALVAR={'true' if salvar else 'false'};"
            "if(location.pathname.indexOf('TabMeses')>=0)return 'PARADA: tela de lista (Contingência/Índice): a máscara corta valor preenchido por JavaScript; digitar com o teclado (seção 4)';"
            "let R=_rows();if(P!==null)R=R.filter(tr=>new RegExp('^\\\\s*'+P+'\\\\s').test(tr.innerText));"
            "if(R.length!==1)return 'PARADA: '+R.length+' linhas candidatas (esperado 1)';"
            "const ins=[...R[0].querySelectorAll('input[type=text]')].filter(i=>i.offsetParent);"
            "if(ins.length!==1)return 'PARADA: '+ins.length+' campos visiveis na linha (esperado 1)';"
            "if(ins[0].disabled)return 'PARADA: período fechado no iGile (campo desabilitado). Só o período do mês corrente aceita lançamento';"
            "if(ins[0].readOnly)return 'PARADA: campo somente leitura (tela de lista: sequência não validada)';"
            "const antes=ins[0].value;_set(ins[0],V);"
            "if(!SALVAR)return 'anterior='+antes+' | campo='+ins[0].value+' | preenchido sem salvar';"
            "const g=await _gravar();"
            "return 'anterior='+antes+' | novo='+ins[0].value+' | '+g;"
            "})()")


def js_valor(a):
    js = _js_valor(a.valor, a.periodo, a.salvar)
    print(js)
    return js


# ---------------- lotes prontos para o browser_batch (v2.2) ----------------
# Abrir uma tela e Voltar sao navegacoes: cada passo e um item separado do lote, e no maximo
# 4 telas por lote (lote de 60 itens estourou o tempo em 29/10/2026).

TELAS = {  # campo do JSON -> (titulo no iGile, tipo)
    "custo_real": ("Custo Real", "nivel"), "contingencia": ("Contingência de Custo", "lista"),
    "indice": ("Índice", "lista"), "estoque": ("Estoque", "nivel"),
    "adiantamento": ("Adiantamento ao Fornecedor", "nivel"), "previsao_despesas": ("Previsão de Despesas", "nivel"),
    "retencao": ("Retenção", "nivel"), "a_pagar": ("Á Pagar", "nivel"),
    "agregado_nao_pago": ("Agregado não Pago", "nivel"), "medicao_exclusiva_custo": ("Medição exclusiva de custo", "nivel8"),
    "realizado_recursos": ("Realizado de Recursos", "nivel"), "medicao": ("Medição", "nivel"),
    "realizado_banco": ("Realizado Banco", "nivel"),
}
CASAS_TELA = {"indice": 4, "medicao": 2, "realizado_banco": 2, "realizado_recursos": 0}

def _js_abrir(t):
    t = json.dumps(t, ensure_ascii=False)
    return ("(()=>{const a=[...document.querySelectorAll('a.btn-xs')].find(a=>(a.title||a.getAttribute('data-original-title')||'').trim()==="
            + t + ");if(!a)return 'PARADA: tela '+" + t + "+' não existe nesta obra';a.click();return 'aberta '+" + t + ";})()")


JS_VOLTAR = "(()=>{const a=[...document.querySelectorAll('a')].find(a=>a.innerText.trim()==='Voltar');if(!a)return 'sem Voltar';a.click();return 'ok';})()"
JS_STATUS = ("(()=>{const t=document.body.innerText;const p=(t.match(/TAREFAS DO PER[ÍI]ODO\\s*\\d+[^\\n]{0,26}/)||['período não achado'])[0];"
             "const r=[...document.querySelectorAll('a.btn-xs')].map(a=>{const ti=(a.title||a.getAttribute('data-original-title')||'').trim();"
             "const c=(String(a.className).match(/btn-(danger|white)/)||[''])[0];return c?ti+': '+(c==='btn-white'?'preenchido':'vazio'):null}).filter(Boolean);"
             "return (p+' || '+r.join(' | ')).replace(/[?&=]/g,' ');})()")


def _js_ler(titulo, per=None):
    filtro = "true" if per is None else f"new RegExp('^\\\\s*{int(per)}\\\\s').test(tr.innerText)"
    return ("(()=>{const R=[...document.querySelectorAll('tr')].filter(tr=>[...tr.querySelectorAll('input[type=text]')].some(i=>i.offsetParent)&&(" + filtro + "));"
            f"return {json.dumps(titulo, ensure_ascii=False)}+'='+R.slice(0,3).map(tr=>[...tr.querySelectorAll('input[type=text]')].filter(i=>i.offsetParent).map(i=>i.value).join(';')).join(' | ');}})()")


def _item_js(tab, js):
    return {"name": "javascript_tool", "input": {"tabId": tab, "action": "javascript_exec", "text": js}}


def _item_wait(tab, s):
    return {"name": "computer", "input": {"tabId": tab, "action": "wait", "duration": s}}


# ---------------- leitura rapida (v2.4): uma chamada le o periodo inteiro ----------------
# O iGile carrega cada tela por POST /EntradaDados/GetDados?o=<orcamento><sufixo>&m=<periodo>&t=<tipo>&n=<nivel>
# (so leitura, devolve JSON). Mapeado em 29/10/2026 contra valores gravados no Kronos Casas e Loteamento.
# Indice e Contingencia vem do HTML de Orcamento/TabMeses (data-token do icone da tela).
GETDADOS = {  # campo -> (sufixo do orcamento, tipo t, fator para exibir)
    "custo_real": ("-c", 8, 1), "estoque": ("-c", 7, 1), "retencao": ("-c", 9, 1),
    "medicao_exclusiva_custo": ("-c", 10, 1), "adiantamento": ("-c", 13, 1),
    "previsao_despesas": ("-c", 14, 1), "a_pagar": ("-c", 17, 1), "agregado_nao_pago": ("-c", 18, 1),
    "medicao": ("", 4, 100), "realizado_banco": ("", 16, 100), "realizado_recursos": ("-r", 3, 1),
}
LISTAS = {"indice": "Índice", "contingencia": "Contingência de Custo"}


def _js_rapido(esperado=None, itens8=None):
    """JS (uma chamada, sem navegar) que devolve periodo, icones, valor de cada tela e, com o JSON,
    as divergencias planilha x iGile e as do nivel 8. Saida sem '=' (a extensao bloqueia pares chave=valor)."""
    mapa = json.dumps(GETDADOS, ensure_ascii=False)
    listas = json.dumps(LISTAS, ensure_ascii=False)
    esp = json.dumps(esperado or {}, ensure_ascii=False)
    e8 = json.dumps(itens8 or [], ensure_ascii=False)
    return ("await (async()=>{const MAPA=" + mapa + ",LISTAS=" + listas + ",ESP=" + esp + ",E8=" + e8 + ";"
            "const txt=document.body.innerText;const pm=txt.match(/TAREFAS DO PER[ÍI]ODO\\s*(\\d+)/);"
            "if(!pm)return 'PARADA: abrir a página da obra (período não encontrado)';const M=+pm[1];"
            "const bt=t=>[...document.querySelectorAll('a.btn-xs')].find(a=>(a.title||a.getAttribute('data-original-title')||'').trim()===t);"
            "const cr=bt('Custo Real');let O=null;try{O=JSON.parse(atob(decodeURIComponent(cr.getAttribute('data-token')))).OrcamentoId.replace(/-c$/,'');}catch(e){}"
            "if(!O)return 'PARADA: orçamento não identificado pelo ícone Custo Real';"
            "const q=(o,t,n)=>fetch('/EntradaDados/GetDados?o='+o+'&m='+M+'&t='+t+'&n='+n,{method:'POST',credentials:'include',headers:{'X-Requested-With':'XMLHttpRequest'}}).then(r=>r.ok?r.json():null).catch(()=>null);"
            "const V={};await Promise.all(Object.entries(MAPA).map(async([k,[suf,t,f]])=>{const j=await q(O+suf,t,1);V[k]=(Array.isArray(j)&&j[0])?Math.round(j[0].Valor*f*1e6)/1e6:null;}));"
            "await Promise.all(Object.entries(LISTAS).map(async([k,ti])=>{const a=bt(ti);if(!a){V[k]=null;return;}const tk=a.getAttribute('data-token');"
            "const h=await fetch('/Orcamento/TabMeses?t='+encodeURIComponent(tk),{credentials:'include'}).then(r=>r.text()).catch(()=>'');"
            "const d=new DOMParser().parseFromString(h,'text/html');const tr=[...d.querySelectorAll('tr')].find(t=>new RegExp('^\\\\s*'+M+'\\\\s').test(t.textContent.trim()));"
            "const i=tr&&tr.querySelector('input[type=text]');V[k]=i?parseFloat(String(i.getAttribute('value')).replace(/\\./g,'').replace(',','.')):null;}));"
            "const ic=[...document.querySelectorAll('a.btn-xs')].map(a=>{const ti=(a.title||a.getAttribute('data-original-title')||'').trim();const c=(String(a.className).match(/btn-(danger|white)/)||[''])[0];return c?ti+': '+(c==='btn-white'?'ok':'vazio'):null}).filter(Boolean);"
            "const L=['PERÍODO '+M+' | orçamento '+O];L.push('VALORES | '+Object.entries(V).map(([k,v])=>k+': '+(v===null?'-':v)).join(' | '));"
            "const dv=Object.entries(ESP).filter(([k,e])=>e!==null&&V[k]!==undefined).map(([k,e])=>{const v=V[k];const tol=(k==='medicao'||k==='realizado_banco'||k==='medicao_exclusiva_custo')?0.01:(k==='indice'?0.0001:1);"
            "return (v===null||Math.abs(v-e)>tol)?k+': iGile '+v+' planilha '+e:null}).filter(Boolean);"
            "L.push('DIVERGÊNCIAS | '+(dv.join(' | ')||'nenhuma'));"
            "if(E8.length){const j8=await q(O+'-c',10,8);if(!Array.isArray(j8))L.push('NÍVEL 8 | sem leitura');else{"
            "const nn=s=>String(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase().replace(/\\s+/g,' ').trim();"
            "const d8=[];if(j8.length!==E8.length)d8.push('linhas iGile '+j8.length+' planilha '+E8.length);"
            "else j8.forEach((r,k)=>{const [nm,ev]=E8[k];if(ev!==null&&Math.abs((r.Valor||0)-ev)>0.005)d8.push((k+1)+' '+String(r.Descricao).slice(0,28)+': '+r.Valor+' planilha '+ev);});"
            "L.push('NÍVEL 8 | '+j8.length+' linhas | '+(d8.slice(0,12).join(' | ')||'sem divergência'));}}"
            "L.push('ÍCONES | '+ic.join(' | '));return L.join(' || ').replace(/[?&=]/g,' ');})()")


def lote(a):
    tab = int(a.tab)
    passos = []  # cada passo = lista de itens de uma tela
    if a.acao == "status":
        print(json.dumps([_item_js(tab, JS_STATUS)], ensure_ascii=False))
        return 0
    if a.acao == "rapido":
        esp, e8 = {}, []
        if a.json:
            C = json.load(open(a.json, encoding="utf-8"))["campos"]
            for k in list(GETDADOS) + list(LISTAS):
                v = (C.get(k) or {}).get("valor")
                if isinstance(v, (int, float)) and (C.get(k) or {}).get("status") != "pendente":
                    esp[k] = v
            mec = C.get("medicao_exclusiva_custo") or {}
            if mec.get("itens") and (mec.get("status") != "pendente"):
                e8 = [[norm(i.get("descricao")), i.get("valor_pct")] for i in mec["itens"]]
        print(json.dumps([_item_js(tab, _js_rapido(esp, e8))], ensure_ascii=False))
        return 0
    if a.acao == "leitura":
        for c in a.campos:
            if c not in TELAS:
                sys.exit(f"campo desconhecido: {c} (use: {', '.join(TELAS)})")
            t, tipo = TELAS[c]
            per = a.periodo if tipo == "lista" else None
            if tipo == "lista" and per is None:
                sys.exit(f"{c} é tela de lista: informe --periodo N")
            passos.append([_item_js(tab, _js_abrir(t)), _item_wait(tab, 3),
                           _item_js(tab, _js_ler(t, per)), _item_js(tab, JS_VOLTAR), _item_wait(tab, 3)])
    elif a.acao == "gravar":
        if not a.json:
            sys.exit("gravar precisa de --json da obra")
        d = json.load(open(a.json, encoding="utf-8"))
        C = d["campos"]
        for c in a.campos:
            if c not in TELAS:
                sys.exit(f"campo desconhecido: {c}")
            t, tipo = TELAS[c]
            if tipo != "nivel":
                sys.exit(f"{c}: tela '{tipo}' não grava por lote (nível 8 usa js-nivel8; telas de lista ainda sem sequência validada)")
            v = (C.get(c) or {}).get("valor")
            if v is None:
                sys.exit(f"{c}: sem valor no JSON ({(C.get(c) or {}).get('alerta', 'sem fonte')}). Não lançar.")
            vb = br(v, CASAS_TELA.get(c, 2)) if isinstance(v, (int, float)) else str(v)
            passos.append([_item_js(tab, _js_abrir(t)), _item_wait(tab, 3),
                           _item_js(tab, _js_valor(vb, None, True)), _item_js(tab, JS_VOLTAR), _item_wait(tab, 4)])
        passos.append([_item_js(tab, JS_STATUS)])
    lotes, atual = [], []
    for p in passos:
        if len(atual) + len(p) > 20:
            lotes.append(atual)
            atual = []
        atual += p
    if atual:
        lotes.append(atual)
    for i, l in enumerate(lotes, 1):
        print(f"### LOTE {i}/{len(lotes)} ({len(l)} itens)")
        print(json.dumps(l, ensure_ascii=False))
    return 0


# ---------------- log ----------------

def log_xlsx(a):
    """Gera o log do mes: uma aba por obra + aba de ajustes de equilibrio."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    resumo = wb.active
    resumo.title = "Resumo"
    tinta = Font(name="Calibri", size=10, color="2B2B2B")
    cab = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    fundo = PatternFill("solid", fgColor="3E5C50")
    alerta = PatternFill("solid", fgColor="F6E7C1")
    fina = Border(bottom=Side(style="hair", color="B8B8B8"))
    resumo.append(["Atualização iGile", "", "", ""])
    resumo["A1"].font = Font(name="Calibri", size=14, bold=True, color="3E5C50")
    resumo.append([])
    resumo.append(["Obra", "Campos lançados", "Pendências", "Ajuste de equilíbrio (R$)"])
    for c in resumo[3]:
        c.font, c.fill = cab, fundo
    for p in a.jsons:
        d = json.load(open(p, encoding="utf-8"))
        ws = wb.create_sheet(d["obra"][:28])
        ws.append(["Campo", "Valor", "Valor anterior no iGile", "Status", "Origem", "Alerta"])
        for c in ws[1]:
            c.font, c.fill = cab, fundo
        lanc = pend = 0
        for k, v in d["campos"].items():
            ws.append([k, v.get("valor"), v.get("anterior"), v.get("status", ""), v.get("origem"), v.get("alerta", "")])
            row = ws[ws.max_row]
            for c in row:
                c.font, c.border = tinta, fina
            if v.get("alerta"):
                for c in row:
                    c.fill = alerta
            if v.get("status") == "lançado":
                lanc += 1
            elif v.get("valor") is None or v.get("status") in ("pendente", "não lançado"):
                pend += 1
        aj = d.get("ajustes", {})
        tot = sum(abs(x) for x in (aj.get("delta_previsto", 0), aj.get("delta_previsao", 0)))
        resumo.append([d["obra"], lanc, pend, round(tot, 2)])
        for col, w in zip("ABCDEF", (30, 18, 22, 14, 70, 60)):
            ws.column_dimensions[col].width = w
        for row in ws.iter_rows(min_row=2):
            row[1].number_format = '#,##0.00'
            row[2].number_format = '#,##0.00'
            row[4].alignment = row[5].alignment = Alignment(wrap_text=True, vertical="top")
        if aj:
            ws.append([])
            ws.append(["Ajustes de equilíbrio (lançados para o OPR bater com a planilha; não são dados reais)"])
            ws[ws.max_row][0].font = Font(name="Calibri", size=11, bold=True, color="8A5A00")
            for k, v in aj.items():
                # v2.3: o agente pode gravar valores aninhados (dict/lista); o openpyxl so aceita escalares
                ws.append([k, v if isinstance(v, (int, float, str)) or v is None else json.dumps(v, ensure_ascii=False)])
                ws[ws.max_row][1].number_format = '#,##0.00'
    for col, w in zip("ABCD", (42, 18, 14, 26)):
        resumo.column_dimensions[col].width = w
    for row in resumo.iter_rows(min_row=4):
        row[3].number_format = '#,##0.00'
    wb.save(a.saida)
    print(a.saida)


def b64(src, outdir):
    """Converte o resultado salvo do download do Google Drive (JSON com base64) em arquivo."""
    import base64
    d = json.load(open(src, encoding="utf-8"))
    os.makedirs(outdir, exist_ok=True)
    p = os.path.join(outdir, d["title"])
    with open(p, "wb") as f:
        f.write(base64.b64decode(d["content"]))
    print(p)


# ---------------- linha de comando ----------------

def extrair(argv):
    ap = argparse.ArgumentParser(prog="igile.py")
    ap.add_argument("--obra", required=True)
    ap.add_argument("--mes", required=True, help="AAAA-MM do período de referência")
    ap.add_argument("--custo-mes", help="planilha de análise de custos do período (mês pelo nome, maior revisão)")
    ap.add_argument("--custo-prox", help=argparse.SUPPRESS)
    ap.add_argument("--curva")
    ap.add_argument("--pls")
    ap.add_argument("--efetivo")
    ap.add_argument("--efetivo-manual", type=int)
    ap.add_argument("--saida", required=True)
    a = ap.parse_args(argv)
    if a.custo_prox:
        sys.exit("--custo-prox saiu na v2.1: o custo real do período M vem da própria planilha M "
                 "(Incorrido N-1). Informe só --custo-mes com a planilha do mês que está no nome.")
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", a.mes):
        sys.exit(f"--mes deve ser AAAA-MM (recebi {a.mes})")
    if a.custo_mes:
        mn, _ = mes_do_nome(a.custo_mes)
        if mn and mn != a.mes:
            sys.exit(f"o nome da planilha de custos diz {mn}, mas --mes é {a.mes}. Vale o nome: use --mes {mn}.")
    r = {"obra": a.obra, "mes": a.mes, "versao_script": VERSAO,
         "gerado_em": dt.datetime.now().isoformat(timespec="seconds"), "campos": {}}
    C = r["campos"]
    informados = set()
    if a.custo_mes:
        informados.add("custo_mes")
        C.update(custo_mes(a.custo_mes, a.mes))
        C["custo_real"] = incorrido_do_mes(a.custo_mes, mes_anterior(a.mes))
    else:
        for k in ("contingencia", "diferenca", "retencao", "a_pagar", "custo_real"):
            C[k] = campo(None, "-", "planilha de custos do mês não informada")
    if a.curva:
        informados.add("curva")
        mc, _ = mes_do_nome(a.curva)
        C.update(curva(a.curva, a.mes, a.obra))
        if mc and mc != a.mes:
            for k in ("medicao", "avanco_fisico_acumulado", "medicao_exclusiva_custo"):
                C[k]["alerta"] = f"Curva Física é de {mc} pelo nome, não em {a.mes}"
    else:
        for k in ("medicao", "avanco_fisico_acumulado", "medicao_exclusiva_custo"):
            C[k] = campo(None, "-", "curva física não informada")
    if a.pls:
        informados.add("pls")
        C["realizado_banco"] = pls(a.pls)
    else:
        C["realizado_banco"] = campo(None, "-", "PLS não informada")
    if a.efetivo_manual is not None:
        C["realizado_recursos"] = campo(a.efetivo_manual, "informado manualmente")
    elif a.efetivo:
        informados.add("efetivo")
        C["realizado_recursos"] = efetivo(a.efetivo, a.mes)
    else:
        C["realizado_recursos"] = campo(None, "-", "efetivo não informado")
    paradas = validar(r, informados)
    r["paradas_layout"] = paradas
    with open(a.saida, "w", encoding="utf-8") as f:
        json.dump(r, f, ensure_ascii=False, indent=1, default=str)
    for k, v in C.items():
        extra = f" ({len(v['itens'])} linhas)" if v.get("itens") else ""
        print(f"{k:26s} {str(v['valor']):>16s}{extra}  {v.get('alerta', '')}")
    if paradas:
        print("\nPARADA (layout mudou ou arquivo de outro mês). Não lançar estes campos antes de ajustar:")
        for p in paradas:
            print("  -", p)
        sys.exit(2)
    print("\nOK: todos os campos dos arquivos informados foram lidos.")


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    cmd = argv[0] if argv and not argv[0].startswith("-") else "extrair"
    rest = argv[1:] if cmd != "extrair" or (argv and argv[0] == "extrair") else argv
    if cmd == "b64":
        return b64(rest[0], rest[1])
    if cmd == "ultima":
        ap = argparse.ArgumentParser(prog="igile.py ultima")
        ap.add_argument("nomes", nargs="+", help="nomes dos arquivos da pasta da obra (do Drive ou do disco)")
        ap.add_argument("--tipo", choices=sorted(_TIPOS), help="filtra pelo tipo de arquivo")
        a = ap.parse_args(rest)
        r = escolher(a.nomes, a.tipo)
        for c in r["candidatos"]:
            marca = ">>" if r["escolha"] is c else "  "
            extra = "  [cópia]" if c["copia"] else ""
            print(f"{marca} {c['mes'] or '???????'}  R{c['revisao']:02d}  {os.path.basename(c['arquivo'])}{extra}"
                  + (f"  ({c['alerta']})" if c["alerta"] else ""))
        if r["escolha"] is None:
            print(f"\nPARADA: {r.get('alerta')}")
            sys.exit(1)
        print(f"\nESCOLHA: {os.path.basename(r['escolha']['arquivo'])} | período {r['escolha']['mes']} | R{r['escolha']['revisao']:02d}")
        return 0
    if cmd == "ajuste":
        ap = argparse.ArgumentParser(prog="igile.py ajuste")
        ap.add_argument("--etapa", choices=["previsto", "previsao"], required=True)
        ap.add_argument("--json", help="JSON da obra: lê diferença, contingência e avanço exato")
        ap.add_argument("--diferenca", type=float)
        ap.add_argument("--contingencia", type=float)
        ap.add_argument("--acumulado", type=float, help="fração, ex. 0.4603")
        ap.add_argument("--gerado", type=float, required=True, help="Variação Término lida no OPR")
        ap.add_argument("--estoque-atual", type=float, default=0)
        ap.add_argument("--adiantamento-atual", type=float, default=0)
        ap.add_argument("--previsao-atual", type=float, default=0)
        ap.add_argument("--agregado-atual", type=float, default=0)
        ap.add_argument("--parcela-estoque", type=float, default=0.5, help="fração do delta para Estoque")
        ap.add_argument("--estoque-fixo", type=float, default=None, help="valor fixo para Estoque; resto vai para Adiantamento")
        return ajustes(ap.parse_args(rest))
    if cmd == "js-nivel8":
        ap = argparse.ArgumentParser(prog="igile.py js-nivel8")
        ap.add_argument("json")
        ap.add_argument("--modo", choices=["ler", "aplicar"], default="ler")
        ap.add_argument("--salvar", action="store_true", help="só no modo aplicar, com comando do Pedro")
        ap.add_argument("--max-divergencias", type=int, default=2)
        ap.add_argument("--saida")
        return js_nivel8(ap.parse_args(rest))
    if cmd == "js-valor":
        ap = argparse.ArgumentParser(prog="igile.py js-valor")
        ap.add_argument("valor", help="formato brasileiro, ex. 3043351,44")
        ap.add_argument("--periodo", type=int, help="telas de lista (Contingência, Índice)")
        ap.add_argument("--salvar", action="store_true")
        return js_valor(ap.parse_args(rest))
    if cmd == "lote":
        ap = argparse.ArgumentParser(prog="igile.py lote")
        ap.add_argument("acao", choices=["rapido", "status", "leitura", "gravar"])
        ap.add_argument("--tab", required=True, help="tabId do Chrome")
        ap.add_argument("--campos", nargs="*", default=[], help=f"campos: {', '.join(TELAS)}")
        ap.add_argument("--periodo", type=int, help="período N (telas de lista)")
        ap.add_argument("--json", help="JSON da obra (gravar)")
        return lote(ap.parse_args(rest))
    if cmd == "log":
        ap = argparse.ArgumentParser(prog="igile.py log")
        ap.add_argument("jsons", nargs="+")
        ap.add_argument("--saida", required=True)
        return log_xlsx(ap.parse_args(rest))
    if cmd == "extrair":
        return extrair(rest)
    sys.exit(f"comando desconhecido: {cmd}")


if __name__ == "__main__":
    main()
