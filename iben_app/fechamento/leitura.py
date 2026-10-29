"""Leitura da Analise de Custos para o e-mail de fechamento (regras da skill
email-fechamento-custo-iben). So fatos da planilha: nada de estimativa ou auditoria aqui."""
import os
import re

import openpyxl
from openpyxl.utils import get_column_letter as L

from iben_app.comum import norm, num

RX_EAP = re.compile(r"^\d+(\.\d+)*$")
LIMITE_EAP = 1000.0  # EAP com pelo menos R$ 1.000 fora do INCC e citada no texto


def _aba(wb, nome):
    for ws in wb.worksheets:
        if norm(ws.title) == norm(nome):
            return ws
    raise KeyError(f"aba '{nome}' não encontrada")


def _achar(ws, texto, max_row=80, max_col=40):
    alvo = norm(texto)
    for r in range(1, min(ws.max_row, max_row) + 1):
        for c in range(1, min(ws.max_column, max_col) + 1):
            if alvo in norm(ws.cell(r, c).value):
                return r, c
    return None, None


def mes_analise(ws):
    for r in range(1, 12):
        for c in range(1, 20):
            if norm(ws.cell(r, c).value) in ("mes analise", "mes de analise"):
                for rr, cc in ((r + 1, c), (r, c + 1)):
                    v = ws.cell(rr, cc).value
                    if hasattr(v, "strftime"):
                        return v.strftime("%Y-%m"), f"{L(cc)}{rr}"
    return None, None


def _subtotal_range(f):
    m = re.match(r"=\s*SUBTOTAL\(\s*\d+\s*,\s*\$?[A-Z]+\$?(\d+)\s*:\s*\$?[A-Z]+\$?(\d+)\s*\)\s*$", str(f or ""), re.I)
    return (int(m.group(1)), int(m.group(2))) if m else None


def analise_financeira(path):
    """Bloco ANALISE PREVISAO - 100% da aba Analise Financeira."""
    wb = openpyxl.load_workbook(path, data_only=True)
    wf = openpyxl.load_workbook(path, data_only=False)
    ws, wsf = _aba(wb, "Análise Financeira"), _aba(wf, "Análise Financeira")
    alertas = []
    mes, cel_mes = mes_analise(ws)
    r_bloco, _ = _achar(ws, "ANÁLISE PREVISÃO - 100%")
    if not r_bloco:
        raise ValueError("bloco 'ANÁLISE PREVISÃO - 100%' não encontrado na Análise Financeira")
    hr, cols = None, {}
    for r in range(r_bloco, r_bloco + 5):
        t = {norm(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
        if "diferenca" in t:
            hr, cols = r, t
            break
    if not hr:
        raise ValueError("cabeçalho com DIFERENÇA não encontrado abaixo do bloco ANÁLISE PREVISÃO - 100%")

    def col(*nomes, obrig=True):
        for n in nomes:
            if norm(n) in cols:
                return cols[norm(n)]
        if obrig:
            raise ValueError(f"coluna {nomes[0]} não encontrada no bloco (linha {hr})")
        return None
    C = {"dif": col("DIFERENÇA"), "ant": col("DIFERENÇA MÊS ANTERIOR"),
         "incc": col("VARIAÇÃO (INCC - MÊS)", "VARIAÇÃO (INCC-MÊS)", "VARIACAO INCC"), "var": col("VARIAÇÃO"),
         "just": col("JUSTIFICATIVA DO DESVIO", "JUSTIFICATIVA", obrig=False)}
    # EAP e descricao: no cabecalho do bloco ou ate 3 linhas acima
    c_eap = c_desc = None
    for r in range(hr, max(hr - 4, 0), -1):
        for c in range(1, 12):
            t = norm(ws.cell(r, c).value)
            if t in ("eap", "item") and not c_eap:
                c_eap = c
            elif t.startswith("descri") and not c_desc:
                c_desc = c
    c_eap, c_desc = c_eap or 2, c_desc or 3
    linhas, total = [], None
    for r in range(hr + 1, ws.max_row + 1):
        eap = ws.cell(r, c_eap).value
        desc = ws.cell(r, c_desc).value
        f_var = wsf.cell(r, C["var"]).value
        if total is None and not (eap and RX_EAP.match(str(eap).strip())) and _subtotal_range(f_var):
            total = {"linha": r, "rotulo": desc, "dif": ws.cell(r, C["dif"]).value, "ant": ws.cell(r, C["ant"]).value,
                     "var": ws.cell(r, C["var"]).value, "incc": ws.cell(r, C["incc"]).value,
                     "faixa": _subtotal_range(f_var), "formula_incc": wsf.cell(r, C["incc"]).value}
            continue
        if isinstance(ws.cell(r, C["dif"]).value, str) and ws.cell(r, C["dif"]).value.strip():
            break  # outro cabecalho: fim do bloco
        if eap is None or not RX_EAP.match(str(eap).strip()):
            continue
        val = lambda k: ws.cell(r, C[k]).value if C[k] and num(ws.cell(r, C[k]).value) else 0.0  # noqa: E731
        linhas.append({"linha": r, "eap": str(eap).strip(), "descricao": desc, "dif": val("dif"), "ant": val("ant"),
                       "incc": val("incc"), "var": val("var"),
                       "just": (ws.cell(r, C["just"]).value or "") if C["just"] else "",
                       "subtotal": str(f_var or "").upper().lstrip("=").startswith("SUBTOTAL")})
    if total is None:
        # sem linha SUBTOTAL: usa a ultima linha sem EAP com valor como total e avisa
        alertas.append("total do bloco não é SUBTOTAL: não dá para separar o que entra no resultado; usei as EAPs folha")
    folhas = {l["eap"] for l in linhas} - {l2["eap"].rsplit(".", 1)[0] for l2 in linhas if "." in l2["eap"]}
    for ln in linhas:
        if total:
            a, b = total["faixa"]
            ln["no_total"] = a <= ln["linha"] <= b and not ln["subtotal"]
        else:
            ln["no_total"] = ln["eap"] in folhas
    return {"arquivo": path, "mes": mes, "celula_mes": cel_mes, "linhas": linhas, "total": total,
            "colunas": {k: (L(v) if v else None) for k, v in C.items()}, "linha_cabecalho": hr, "alertas": alertas}


def desvios(path):
    """Bloco RESULTADO DA ANALISE DO PREVISAO - 100% da aba Desvios, com cores da planilha."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _aba(wb, "Desvios")
    r0, c0 = _achar(ws, "RESULTADO DA ANÁLISE DO PREVISÃO")
    if not r0:
        raise ValueError("bloco 'RESULTADO DA ANÁLISE DO PREVISÃO - 100%' não encontrado na aba Desvios")
    hr = r0 + 1
    cols = [(c, ws.cell(hr, c).value) for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value not in (None, "")]
    if not cols:
        raise ValueError("cabeçalho do bloco da aba Desvios vazio")

    def cor(cell):
        f = cell.fill
        if f is not None and f.fill_type == "solid" and isinstance(f.fgColor.rgb, str) and len(f.fgColor.rgb) == 8:
            return "#" + f.fgColor.rgb[2:]
        return None
    cab = [{"titulo": str(t), "fundo": cor(ws.cell(hr, c))} for c, t in cols]
    c_var = next((c for c, t in cols if norm(t) == "variacao"), None)
    c_dif = next((c for c, t in cols if norm(t) == "diferenca"), None)
    linhas, final = [], None
    for r in range(hr + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c, _ in cols]
        if all(v in (None, "") for v in vals):
            continue
        rot = " ".join(norm(v) for v in vals if isinstance(v, str))
        reg = {"linha": r, "valores": vals, "fundos": [cor(ws.cell(r, c)) for c, _ in cols]}
        if "resultado final" in rot:
            final = reg
            break
        if c_var and num(ws.cell(r, c_var).value) and abs(ws.cell(r, c_var).value) < 0.005:
            continue  # variacao zero sai do print
        linhas.append(reg)
    if not final:
        raise ValueError("linha RESULTADO FINAL não encontrada na aba Desvios")
    tot_dif = ws.cell(final["linha"], c_dif).value if c_dif else None
    return {"cabecalho": cab, "linhas": linhas, "final": final, "total_diferenca": tot_dif,
            "titulo": ws.cell(r0, c0).value}


def nome_anexo(path):
    n = os.path.basename(path)
    limpo = re.sub(r"^\s*c[oó]pia de\s+", "", n, flags=re.I)
    return limpo, (limpo != n)
