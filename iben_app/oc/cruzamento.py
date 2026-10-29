"""Cruzamento OC x Memoria de Calculo - Previsao, em quatro camadas (nesta ordem):
1. Nº pedido (coluna G), 2. citacao em texto (D e F), 3. semelhanca (so linhas sem OC),
4. sem correspondencia (so depois de procurar codigo e valor na aba inteira, inclusive formulas).
"""
import re

import openpyxl

from iben_app.comum import norm, num

ABA = "Memória de Cálculo - Previsão"
VAZIAS = set("de da do das dos e em para com por a o as os na no nas nos um uma ref referente conforme "
             "servico servicos fornecimento material materiais ltda me epp eireli sa s/a comercio "
             "industria industrial engenharia construcoes construtora e/ou".split())
FORN_VAZIAS = VAZIAS | {"comercial", "distribuidora", "servicos", "solucoes", "brasil"}


def _aba(wb):
    alvo = norm(ABA)
    for ws in wb.worksheets:
        if norm(ws.title) == alvo or (alvo in norm(ws.title)):
            return ws
    raise KeyError(f"aba '{ABA}' não encontrada (abas: {[w.title for w in wb.worksheets]})")


def _colunas(ws):
    """Cabecalho na linha 7 (padrao); procura o rotulo N° PEDIDO nas linhas 1-12 por seguranca."""
    padrao = {"eap": 2, "tipo": 3, "item": 4, "fornecedor": 5, "obs": 6, "pedido": 7, "valor": 8}
    for r in range(1, 13):
        cols = {norm(ws.cell(r, c).value): c for c in range(1, 20) if ws.cell(r, c).value}
        ped = next((c for t, c in cols.items() if "pedido" in t), None)
        if ped:
            def ach(*chaves, d=None):
                return next((c for t, c in cols.items() if any(k in t for k in chaves)), d)
            return r, {"eap": ach("eap", d=2), "tipo": ach("mat", d=3), "item": ach("item", d=4),
                       "fornecedor": ach("fornecedor", d=5), "obs": ach("observ", d=6), "pedido": ped,
                       "valor": ach("valor total", "valor", d=8)}
    return 7, padrao


def ler_memoria(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws = _aba(wb)
    hr, C = _colunas(ws)
    linhas = []
    for r in range(hr + 1, ws.max_row + 1):
        g = lambda k: ws.cell(r, C[k]).value  # noqa: E731
        if not any(g(k) not in (None, "") for k in ("eap", "item", "pedido", "valor")):
            continue
        linhas.append({"linha": r, "eap": g("eap"), "tipo": g("tipo"), "item": g("item"),
                       "fornecedor": g("fornecedor"), "obs": g("obs"), "pedido": g("pedido"),
                       "valor": g("valor") if num(g("valor")) else None})
    # texto completo (valores e formulas) para a camada 4
    todo = []
    for row in ws.iter_rows(values_only=True):
        todo.extend(str(v) for v in row if v not in (None, ""))
    wbf = openpyxl.load_workbook(path, data_only=False, read_only=True)
    try:
        wsf = next(w for w in wbf.worksheets if w.title == ws.title)
        for row in wsf.iter_rows(values_only=True):
            todo.extend(str(v) for v in row if isinstance(v, str) and v.startswith("="))
    finally:
        wbf.close()
    return {"arquivo": path, "aba": ws.title, "linha_cabecalho": hr, "linhas": linhas, "texto": "\n".join(todo)}


# ---------------- camada 1: numero do pedido ----------------

def numeros_na_celula(v, obra):
    """Sequencias de OC citadas numa celula de N° PEDIDO. Retorna lista de (seq, atipico)."""
    if v in (None, ""):
        return []
    s = str(v).replace(" ", "")
    out = []
    # dotted: 400.137.2.00351, 0400.137.2.00351, 400.137.00351 (curto), com sufixo .1 / A / _R01
    for m in re.finditer(rf"0?400\.{obra}\.(?:2\.)?(\d{{3,6}})", s):
        out.append((int(m.group(1)), False))
    # corrido: 400137 + 2 + sequencia de 3 a 6 digitos; 10 digitos 4001370xxx e outro pedido
    for m in re.finditer(rf"(?<!\d)400{obra}(\d+)", s):
        resto = m.group(1)
        if len(resto) == 4 and resto.startswith("0"):
            continue  # 4001370351: formato de 10 digitos, outro pedido (regra da skill)
        if resto.startswith("2") and 4 <= len(resto) <= 6:
            out.append((int(resto[1:]), False))
        elif resto.startswith("2") and len(resto) == 7:
            out.append((int(resto[1:6]), True))  # 13 digitos: formato atipico, 5 primeiros da sequencia
    vistos, unicos = set(), []
    for x in out:
        if x[0] not in vistos:
            vistos.add(x[0])
            unicos.append(x)
    return unicos


# ---------------- camada 3: semelhanca ----------------

def _termos(texto, vazias=VAZIAS):
    t = norm(texto).replace("m.o.", "mo").replace("m.o", "mo").replace("m/o", "mo")
    return {w for w in re.findall(r"[a-z0-9]+", t) if len(w) > 1 and w not in vazias}


def _contencao(a, b):
    return len(a & b) / len(a) if a else 0.0


def _score_valor(v_oc, v_lin):
    if not v_oc or not v_lin:
        return 0.0
    d = abs(v_oc - v_lin) / abs(v_oc)
    return 1.0 if d <= 0.05 else 0.6 if d <= 0.15 else 0.3 if d <= 0.40 else 0.0


def semelhanca(oc, linha):
    titulo = re.sub(r"^\s*0?400[\d\.]+\s*-?\s*", "", oc["pasta"])
    titulo = re.sub(r"\(\s*mov\.?\s*\d*\s*\)", "", titulo, flags=re.I)  # (Mov NNNN) nao e descricao
    t_item = _termos(titulo) | set().union(*[_termos(i["descricao"]) for i in oc.get("itens", [])] or [set()])
    t_tit = _termos(titulo)
    texto_lin = _termos(f"{linha['item'] or ''} {linha['obs'] or ''}")
    item = _contencao(t_tit, texto_lin) if t_tit else _contencao(t_item, texto_lin)
    if t_tit and t_item != t_tit:
        item = max(item, _contencao(t_item, texto_lin))
    nums = set(re.findall(r"\d+", norm(titulo)))
    if nums and not nums & set(re.findall(r"\d+", norm(f"{linha['item']} {linha['obs']}"))):
        item *= 0.6
    forn = _contencao(_termos((oc.get("cab") or {}).get("fornecedor") or "", FORN_VAZIAS),
                      _termos(f"{linha['fornecedor'] or ''}", FORN_VAZIAS))
    val = _score_valor(oc.get("totais", {}).get("total"), linha["valor"])
    score = 0.5 * item + 0.25 * forn + 0.25 * val
    if item >= 0.75 and (forn >= 0.5 or val >= 0.6):
        conf = "Alta"
    elif item >= 0.75:
        conf = "Média"
    elif item >= 0.5 and (forn >= 0.5 or val >= 0.6):
        conf = "Baixa"
    else:
        conf = None
    return {"score": round(score, 3), "item": round(item, 2), "fornecedor": round(forn, 2), "valor": val, "confianca": conf}


# ---------------- cruzamento ----------------

def _eap(v):
    return "" if v is None else str(v).strip()


def cruzar(procs, memoria, obra):
    linhas = memoria["linhas"]
    por_seq = {}
    for ln in linhas:
        ln["seqs"] = numeros_na_celula(ln["pedido"], obra)
        for seq, atip in ln["seqs"]:
            por_seq.setdefault(seq, []).append((ln, atip))
    sem_oc = [ln for ln in linhas if not ln["seqs"]]
    resultado = []
    for pr in procs:
        if pr.get("status") in ("Cancelado", "Fora do escopo") or not pr.get("seq"):
            continue
        seq, total = pr["seq"], (pr.get("totais") or {}).get("total")
        r = {"pr": pr, "metodo": None, "confianca": None, "linhas": [], "citada": [], "eaps": [],
             "valor_memoria": None, "diferenca": None, "situacao": ""}
        # camada 2 (registrada sempre): citacao em texto
        rx = re.compile(rf"\boc'?s?\.?\s*(?:n[º°o]\.?\s*)?(?:[\d/ e,]*\b)?0*{seq}\b", re.I)
        r["citada"] = [ln["linha"] for ln in linhas if rx.search(norm(f"{ln['item'] or ''} | {ln['obs'] or ''}"))]
        if seq in por_seq:
            achados = por_seq[seq]
            r["metodo"], r["confianca"] = "Nº pedido", "Alta"
            r["linhas"] = [ln for ln, _ in achados]
            r["valor_memoria"] = round(sum(ln["valor"] or 0 for ln in r["linhas"]), 2)
            if any(a for _, a in achados):
                r["situacao"] = "formato atípico do Nº pedido; "
            if total is None or all(ln["valor"] is None for ln in r["linhas"]):
                r["situacao"] += "valor não preenchido"
            else:
                r["diferenca"] = round(total - r["valor_memoria"], 2)
                r["situacao"] += "valor bate" if abs(r["diferenca"]) < 1 else "valor difere"
        else:
            cands = []
            for ln in sem_oc:
                s = semelhanca(pr, ln)
                if s["confianca"]:
                    cands.append((s, ln))
            cands.sort(key=lambda x: -x[0]["score"])
            if cands:
                melhor = cands[0][0]["score"]
                cands = [c for c in cands if c[0]["score"] >= melhor - 0.1][:4]
                r["metodo"], r["confianca"] = "Semelhança", cands[0][0]["confianca"]
                r["linhas"] = [ln for _, ln in cands]
                r["valor_memoria"] = cands[0][1]["valor"]
                if total is not None and r["valor_memoria"] is not None:
                    r["diferenca"] = round(total - r["valor_memoria"], 2)
                r["situacao"] = "candidata: validar com o engenheiro"
            elif r["citada"]:
                r["metodo"], r["situacao"] = "Só citada", "OC citada em texto, sem linha com o número"
            else:
                # camada 4: busca do codigo e do valor na aba inteira, inclusive formulas
                txt = memoria["texto"]
                achou_cod = re.search(rf"0*{seq}(?!\d)", txt) and re.search(rf"400\.?{obra}\.?2?\.?0*{seq}(?!\d)", txt.replace(" ", ""))
                achou_val = total is not None and (f"{total:.2f}" in txt or f"{total:.2f}".rstrip("0").rstrip(".") in txt)
                r["metodo"] = "Sem correspondência"
                r["situacao"] = ("código aparece em outra parte da aba: conferir" if achou_cod else
                                 "valor da OC aparece na aba: conferir" if achou_val else
                                 "lacuna a validar (pedido pequeno já pago pode não ter linha: conferir incorrido)")
        r["eaps"] = sorted({_eap(ln["eap"]) for ln in r["linhas"] if ln["eap"] is not None})
        tarefa_pdf = sorted({i["tarefa"] for i in pr.get("itens", []) if i.get("tarefa")})
        if tarefa_pdf and r["eaps"] and not set(tarefa_pdf) & set(r["eaps"]):
            r["situacao"] += " | a validar: EAP do PDF diverge da memória"
        resultado.append(r)
    return resultado
