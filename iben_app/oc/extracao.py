"""Extracao de OCs da pasta Processo de Compras (regras da skill extracao-oc-processo-compras).

Uma subpasta = um processo. O nome da pasta comeca pelo numero da OC
(400.149.2.00484 - descricao (Mov NNNN)). Do PDF saem cabecalho, itens e totais,
com a validacao obrigatoria (qtde x unit = total, soma x subtotal x total).
"""
import datetime as dt
import os
import re
import shutil
import subprocess
from collections import Counter

from iben_app.comum import norm, valor_br

JANELA_MESES = 3
NAO_OC = ("arquivo", "claude outputs", "processos antigos")
VAL = r"-?[\d\.]*\d,\d{2,4}"


def codigo_obra(pastas):
    """Codigo mais frequente 400.xxx nas pastas (ex. '137')."""
    c = Counter(m.group(1) for p in pastas for m in [re.match(r"\s*0?400\.(\d{3})\.", p)] if m)
    return c.most_common(1)[0][0] if c else None


def rx_oc(obra):
    """Regex da OC da obra. Nunca 400.xxx generico (regra da skill)."""
    return re.compile(rf"0?400\.{obra}\.2\.(\d{{5,6}})")


def numero_oc(texto, obra):
    m = rx_oc(obra).search(texto or "")
    return (f"400.{obra}.2.{m.group(1)}", int(m.group(1))) if m else (None, None)


# ---------------- leitura do PDF ----------------

def texto_pdf(path):
    """pdftotext -layout quando existir (melhor alinhamento); senao pdfplumber."""
    if shutil.which("pdftotext"):
        try:
            p = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, timeout=60)
            if p.returncode == 0:
                return p.stdout.decode("utf-8", "ignore")
        except (OSError, subprocess.SubprocessError):
            pass
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        return "\n".join((pg.extract_text(layout=True) or "") for pg in pdf.pages)


def _ultima_revisao(pdfs):
    def rev(n):
        r = re.findall(r"(?:_r|rev\.?\s*)(\d+)", norm(os.path.splitext(n)[0]))
        return int(r[-1]) if r else 0
    return max(pdfs, key=lambda p: (rev(os.path.basename(p)), os.path.getmtime(p)))


RX_ITEM = re.compile(
    rf"^\s*(?:(?P<item>\d{{1,3}})\s+)?(?:(?P<tarefa>\d{{1,3}}(?:\.\d{{1,3}}){{1,6}})\s+)?"
    rf"(?P<desc>.*?)\s+(?P<un>[A-Za-z][A-Za-z0-9²³/\.]{{0,7}})\s+(?P<qt>{VAL}|\d+)\s+"
    rf"(?:R\$\s*)?(?P<vu>{VAL})\s+(?:R\$\s*)?(?P<vt>{VAL})\s*$")
RX_OBS = re.compile(r"^\s*OBS\.?\s*:?\s*(.*)$", re.I)
RX_FIM_ITENS = re.compile(r"^\s*(SUB\s*-?\s*TOTAL|TOTAL|FRETE|DESCONTO|CONDI[CÇ][AÃ]O|PRAZO)", re.I)


def _cabecalho(t):
    def achar(rx):
        m = re.search(rx, t, re.I | re.M)
        return m.group(1).strip() if m else None
    return {
        "n_pedido": achar(r"N[º°o]\.?\s*(?:DO\s+)?PEDIDO\s*:?\s*(\d{5,})"),
        "id_totvs": achar(r"IDENTIFICADOR\s+TOTVS\s*:?\s*(\S+)"),
        "data": achar(r"\bDATA\s*(?:DE\s+EMISS[AÃ]O)?\s*:?\s*(\d{2}/\d{2}/\d{4})"),
        "fornecedor": achar(r"FORNECEDOR\s*:?\s*(.+?)(?:\s{2,}|\s+CNPJ|$)"),
        "cnpj": achar(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})"),
    }


def _totais(t):
    def ult(rx):
        m = re.findall(rx, t, re.I | re.M)
        return valor_br(m[-1]) if m else None
    return {
        # TOTAL so em linha que comeca com TOTAL: (nao pega "valor total" da observacao)
        "total": ult(rf"^\s*TOTAL\s*:\s*\|?\s*(?:R\$)?\s*({VAL})"),
        "subtotal": ult(rf"^\s*SUB\s*-?\s*TOTAL\s*:?\s*\|?\s*(?:R\$)?\s*({VAL})"),
        "frete": ult(rf"FRETE:?\s*\|?\s*\+?\s*R?\$?\s*({VAL})"),
        "desconto": ult(rf"DESCONTO:?\s*\|?\s*-?\s*R?\$?\s*({VAL})"),
    }


def _itens(t):
    """Uma linha por item. Linhas soltas entre itens: as que vem logo depois de um item
    (ate a OBS) completam a descricao dele; as que vem depois da OBS sao do proximo item."""
    itens, soltas, depois_obs, na_tabela = [], [], False, False
    for ln in t.splitlines():
        if not ln.strip():
            continue
        if not na_tabela and not itens:
            n = norm(ln)
            if "descri" in n and ("qtd" in n or "quant" in n or "unid" in n):
                na_tabela = True  # cabecalho da tabela de itens: daqui em diante e item
                continue
        if RX_FIM_ITENS.match(ln) and itens:
            break
        m = RX_ITEM.match(ln)
        if m and valor_br(m["vt"]) is not None:
            it = {"item": int(m["item"]) if m["item"] else None, "tarefa": m["tarefa"],
                  "descricao": m["desc"].strip(), "unid": m["un"],
                  "qtde": valor_br(m["qt"]), "vu": valor_br(m["vu"]), "vt": valor_br(m["vt"]), "obs": ""}
            if itens and soltas and not depois_obs and it["descricao"]:
                itens[-1]["descricao"] = (itens[-1]["descricao"] + " " + " ".join(soltas)).strip()
            elif soltas:
                it["descricao"] = (" ".join(soltas) + " " + it["descricao"]).strip()
            soltas, depois_obs = [], False
            itens.append(it)
            continue
        o = RX_OBS.match(ln)
        if o and itens:
            if soltas and not depois_obs:
                itens[-1]["descricao"] = (itens[-1]["descricao"] + " " + " ".join(soltas)).strip()
            soltas, depois_obs = [], True
            itens[-1]["obs"] = (itens[-1]["obs"] + " " + o.group(1)).strip()
            continue
        if itens or na_tabela:
            soltas.append(" ".join(ln.split()))
    if itens and soltas and not depois_obs:
        itens[-1]["descricao"] = (itens[-1]["descricao"] + " " + " ".join(soltas)).strip()
    return itens


def tipo_documento(t):
    n = norm(t[:1500])
    if "pmo.cot" in n or "cotacao" in n[:300]:
        return "cotação"
    if "orcamento" in n[:300]:
        return "orçamento"
    if "processo de locacao" in n or "formulario de locacao" in n:
        return "formulário de locação"
    return "OC"


def validar(oc):
    """Regras obrigatorias da skill. Preenche oc['pendencias'] e a Observacao dos itens."""
    pend = []
    for it in oc["itens"]:
        if it["qtde"] is None or it["vu"] is None or it["vt"] is None:
            it["conf"] = "valor não lido"
            pend.append(f"item {it['item']}: valor não lido")
            continue
        dif = it["qtde"] * it["vu"] - it["vt"]
        if abs(dif) > 0.01:
            if it["qtde"] and abs(it["vt"] / it["qtde"] - it["vu"]) <= 0.005:
                it["obs"] = (it["obs"] + " | preço unit. arredondado no PDF").strip(" |")
            elif abs(dif) > 0.005 * max(abs(it["vt"]), 0.01):
                pend.append(f"item {it['item']}: qtde x unit difere do total em {dif:,.2f}")
        if not it["unid"] or not it["descricao"]:
            pend.append(f"item {it['item']}: unidade ou descrição não lidas")
    soma = round(sum(i["vt"] or 0 for i in oc["itens"]), 2)
    tt = oc["totais"]
    if tt["subtotal"] is None:
        pend.append("pedido sem linha SUBTOTAL impressa: total não conferido")
    elif abs(soma - tt["subtotal"]) > 0.01:
        pend.append(f"soma dos itens {soma:,.2f} difere do SUBTOTAL {tt['subtotal']:,.2f}")
    if tt["total"] is not None and tt["subtotal"] is not None:
        esperado = tt["subtotal"] + (tt["frete"] or 0) - (tt["desconto"] or 0)
        if abs(esperado - tt["total"]) > 0.01:
            pend.append(f"TOTAL {tt['total']:,.2f} ≠ SUBTOTAL + FRETE − DESCONTO ({esperado:,.2f})")
    if tt["total"] is None:
        pend.append("linha TOTAL: não encontrada")
    oc["soma_itens"] = soma
    oc["pendencias"] = pend
    return oc


def ler_oc(pdf, obra):
    t = texto_pdf(pdf)
    oc = {"pdf": pdf, "texto_chars": len(t.strip())}
    if len(t.strip()) < 200:
        oc.update(status="Verificar", itens=[], totais={k: None for k in ("total", "subtotal", "frete", "desconto")},
                  pendencias=["PDF escaneado (sem texto): conferir no PDF; OCR não roda no app"], cab={})
        return oc
    oc["tipo"] = tipo_documento(t)
    oc["cab"] = _cabecalho(t)
    oc["n_oc_pdf"], _ = numero_oc(t, obra)
    oc["totais"] = _totais(t)
    oc["itens"] = _itens(t)
    validar(oc)
    if oc["tipo"] != "OC":
        oc["pendencias"].insert(0, f"documento é {oc['tipo']}, não OC: há OC emitida?")
    if not oc["itens"]:
        oc["pendencias"].append("nenhum item lido: layout novo, conferir no PDF")
    oc["status"] = "Verificar" if oc["pendencias"] else "OK"
    return oc


# ---------------- varredura da pasta ----------------

def processos(pasta, obra=None, janela_meses=JANELA_MESES, hoje=None):
    """Lista as subpastas (processos) com status. Retorna (obra, lista)."""
    hoje = hoje or dt.datetime.now()
    limite = hoje - dt.timedelta(days=round(30.44 * janela_meses))
    subs = sorted(d for d in os.listdir(pasta) if os.path.isdir(os.path.join(pasta, d)))
    obra = obra or codigo_obra(subs)
    if not obra:
        raise ValueError("não achei o código da obra (400.xxx) no nome das pastas: informe o código")
    contagem = Counter(numero_oc(s, obra)[0] for s in subs)
    out = []
    for s in subs:
        p = os.path.join(pasta, s)
        mod = dt.datetime.fromtimestamp(os.path.getmtime(p))
        n_oc, seq = numero_oc(s, obra)
        reg = {"pasta": s, "caminho": p, "modificada": mod, "n_oc": n_oc, "seq": seq, "obs": []}
        if norm(s) in NAO_OC or norm(s).startswith(NAO_OC):
            continue
        if mod < limite:
            continue  # fora da janela: nao entra (padrao 3 meses)
        outro = re.match(r"\s*0?400\.(\d{3})\.", s)
        if outro and outro.group(1) != obra:
            reg.update(status="Fora do escopo", obs=[f"pasta de outra obra (400.{outro.group(1)}): confirmar"])
            out.append(reg)
            continue
        if "cancelad" in norm(s):
            reg.update(status="Cancelado")
            out.append(reg)
            continue
        if n_oc and contagem[n_oc] > 1:
            reg["obs"].append("código repetido em outra pasta")
        pdfs = [os.path.join(p, f) for f in os.listdir(p) if f.lower().endswith(".pdf")]
        if not pdfs:
            sub_pdf = [os.path.join(r, f) for r, _, fs in os.walk(p) for f in fs if f.lower().endswith(".pdf")]
            reg.update(status="Sem OC", obs=reg["obs"] + (
                [f"sem OC na raiz; {len(sub_pdf)} PDF(s) em subpastas, conferir"] if sub_pdf else ["pasta sem PDF"]))
            out.append(reg)
            continue
        cand = [x for x in pdfs if numero_oc(os.path.basename(x), obra)[0] or re.search(r"\b(oc|pc|pedido)\b", norm(os.path.basename(x)))]
        reg["pdf"] = _ultima_revisao(cand or pdfs)
        out.append(reg)
    return obra, out


def extrair_pasta(pasta, obra=None, janela_meses=JANELA_MESES, hoje=None, progresso=None):
    obra, procs = processos(pasta, obra, janela_meses, hoje)
    for i, pr in enumerate(procs):
        if progresso:
            progresso(i, len(procs), pr["pasta"])
        if "pdf" not in pr:
            pr.setdefault("itens", [])
            continue
        try:
            oc = ler_oc(pr["pdf"], obra)
        except Exception as e:  # PDF corrompido ou protegido: segue com as outras
            oc = {"status": "Verificar", "itens": [], "pendencias": [f"erro ao ler o PDF: {e}"], "cab": {},
                  "totais": {k: None for k in ("total", "subtotal", "frete", "desconto")}}
        if not pr["n_oc"] and oc.get("n_oc_pdf"):
            pr["n_oc"] = oc["n_oc_pdf"]
            pr["seq"] = int(pr["n_oc"].split(".")[-1])
        elif pr["n_oc"] and oc.get("n_oc_pdf") and oc["n_oc_pdf"] != pr["n_oc"]:
            oc["pendencias"].append(f"OC no PDF ({oc['n_oc_pdf']}) difere da pasta")
            oc["status"] = "Verificar"
        pr.update(oc)
        pr["obs"] = pr["obs"] + pr.get("pendencias", [])
    return obra, procs
