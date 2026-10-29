#!/usr/bin/env python3
"""Teste de regressao da skill atualizacao-igile. Rodar antes de cada mes.

  python3 -I teste_regressao.py                 # formulas dos ajustes e JavaScript (sem arquivos)
  python3 -I teste_regressao.py --pasta DIR     # + extracao das planilhas de set/26 achadas em DIR

Sai com codigo 1 se algum caso falhar. Nesse caso, nao lancar nada no iGile:
corrigir o igile.py (ou o caso de referencia, se o Pedro confirmar que mudou) e rodar de novo.
"""
import argparse, json, os, re, shutil, subprocess, sys, tempfile, unicodedata

AQUI = os.path.dirname(os.path.abspath(__file__))
IGILE = os.path.join(AQUI, "..", "..", "iben_app", "igile", "igile.py")
REF = os.path.join(AQUI, "referencia_set26.json")
resultados = []


def ok(caso, passou, detalhe=""):
    resultados.append((caso, passou, detalhe))


def rodar(args):
    p = subprocess.run([sys.executable, "-I", IGILE] + args, capture_output=True, text=True)
    return p.returncode, p.stdout, p.stderr


def compacto(nome):
    s = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def achar(pasta, tokens):
    """Arquivo .xlsx/.xlsm cujo nome contem todos os tokens; maior revisao Rxx."""
    cands = []
    for raiz, _, arqs in os.walk(pasta):
        for a in arqs:
            if not a.lower().endswith((".xlsx", ".xlsm")) or "orcamento" in compacto(a):
                continue
            c = compacto(a)
            if all(t in c for t in tokens):
                m = re.findall(r"r(\d{2})(?:xlsx|xlsm)$", c)
                cands.append((int(m[-1]) if m else -1, os.path.join(raiz, a)))
    return max(cands)[1] if cands else None


def perto(a, b, tol):
    return a is not None and b is not None and abs(a - b) <= tol


def teste_ajustes(ref):
    for c in ref["ajustes"]:
        cod, out, err = rodar(["ajuste"] + c["args"])
        if cod != 0:
            ok(c["caso"], False, err.strip()[:200])
            continue
        r = json.loads(out)
        falhas = [f"{k}: {r.get(k)} x {v}" for k, v in c["esperado"].items() if not perto(r.get(k), v, 0.01)]
        ok(c["caso"], not falhas, "; ".join(falhas))


def teste_extracao(ref, pasta):
    tmp = tempfile.mkdtemp(prefix="igile_teste_")
    try:
        for obra, c in ref["extracao"].items():
            custo = achar(pasta, c["busca"]["custo"])
            curva = achar(pasta, c["busca"]["curva"])
            if not custo or not curva:
                ok(f"extração {obra}", False, f"arquivo não achado em {pasta} (custo={bool(custo)}, curva={bool(curva)})")
                continue
            saida = os.path.join(tmp, "obra.json")
            cod, out, err = rodar(["--obra", obra, "--mes", "2026-09", "--custo-mes", custo, "--curva", curva, "--saida", saida])
            if cod != 0:
                ok(f"extração {obra}", False, (out + err).strip().splitlines()[-1][:200] if (out + err).strip() else f"código {cod}")
                continue
            C = json.load(open(saida, encoding="utf-8"))["campos"]
            falhas = []
            for k, v in c["esperado"].items():
                if k == "itens_nivel8":
                    n = len(C["medicao_exclusiva_custo"].get("itens") or [])
                    if n != v:
                        falhas.append(f"nível 8: {n} linhas x {v}")
                elif not perto(C[k]["valor"], v, 0.0001 if k == "avanco_fisico_acumulado" else 0.01):
                    falhas.append(f"{k}: {C[k]['valor']} x {v}")
            ok(f"extração {obra} ({os.path.basename(custo)[:45]}...)", not falhas, "; ".join(falhas))
            if "alvo_previsto" in c:
                av = C["avanco_fisico_acumulado"]
                alvo = (C["diferenca"]["valor"] + C["contingencia"]["valor"]) * av.get("exato", av["valor"])
                tol = c.get("tolerancia_alvo", 1)
                ok(f"alvo previsto {obra}", perto(alvo, c["alvo_previsto"], tol),
                   f"{alvo:,.2f} x {c['alvo_previsto']:,.2f} (tolerância R$ {tol})")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def teste_escolha(ref):
    """Regra do Pedro (29/10/2026): último mês pelo NOME da planilha e maior revisão."""
    for c in ref.get("escolha_arquivo", []):
        cod, out, err = rodar(["ultima", "--tipo", c["tipo"]] + c["nomes"])
        esp = c["esperado"]
        if esp.get("parada"):
            ok(c["caso"], cod == 1 and "PARADA" in out, f"código {cod}")
            continue
        linha = next((l for l in out.splitlines() if l.startswith("ESCOLHA:")), "")
        passou = cod == 0 and esp["arquivo"] in linha and f"período {esp['mes']}" in linha
        ok(c["caso"], passou, "" if passou else (linha or (out + err).strip()[-200:]))


def teste_js():
    node = shutil.which("node")
    tmp = tempfile.mkdtemp(prefix="igile_js_")
    try:
        exemplo = {"obra": "teste", "campos": {"medicao_exclusiva_custo": {"valor": 5.77, "itens": [
            {"descricao": "VIVA PARQUE UNIQUE", "valor_pct": None},
            {"descricao": "CANTEIRO DE OBRAS", "valor_pct": 5.77},
            {"descricao": "CAÇAMBA + BOTA FORA", "valor_pct": 0}]}}}
        j = os.path.join(tmp, "ex.json")
        json.dump(exemplo, open(j, "w", encoding="utf-8"), ensure_ascii=False)
        gerados = {}
        for nome, args in (("js-nivel8 ler", ["js-nivel8", j, "--modo", "ler"]),
                           ("js-nivel8 aplicar", ["js-nivel8", j, "--modo", "aplicar", "--salvar"]),
                           ("js-valor", ["js-valor", "3043351,44"]),
                           ("js-valor período", ["js-valor", "1296,8890", "--periodo", "16", "--salvar"])):
            cod, out, err = rodar(args)
            if cod != 0:
                ok(nome, False, err.strip()[:200])
                continue
            gerados[nome] = out.strip()
        # 29/10/2026: o primeiro js-valor clicou num "Salvar" oculto das configurações da conta.
        # O JS de gravação só pode usar #botaoCalcular e #salvar pelo id, nunca procurar o texto "Salvar".
        for nome, js in gerados.items():
            if "--salvar" in nome or "aplicar" in nome or "período" in nome:
                seguro = ("getElementById('salvar')" in js and "getElementById('botaoCalcular')" in js
                          and "==='Salvar'" not in js and "buttonChangeApiKey" not in js)
                ok(f"{nome}: grava só por Calcular + #salvar", seguro)
        if not node:
            ok("JavaScript gerado (sintaxe)", True, "node não instalado: sintaxe não conferida")
            return
        for nome, js in gerados.items():
            f = os.path.join(tmp, "t.mjs")
            open(f, "w", encoding="utf-8").write(js)
            p = subprocess.run([node, "--check", f], capture_output=True, text=True)
            ok(f"{nome} (sintaxe)", p.returncode == 0, p.stderr.strip()[:200])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def teste_lote():
    """v2.2: lotes prontos para o browser_batch. Até 20 itens por lote, JS válido, gravação só por Calcular + #salvar,
    e recusa de campo sem valor, de nível 8 e de tela de lista."""
    node = shutil.which("node")
    tmp = tempfile.mkdtemp(prefix="igile_lote_")
    try:
        j = os.path.join(tmp, "obra.json")
        json.dump({"obra": "teste", "campos": {"custo_real": {"valor": 874668.09}, "retencao": {"valor": 333789.33},
                   "a_pagar": {"valor": 641624.36}, "estoque": {"valor": 0}, "medicao": {"valor": None, "alerta": "sem curva"},
                   "contingencia": {"valor": 1000000.0}}}, open(j, "w", encoding="utf-8"))
        cod, out, err = rodar(["lote", "rapido", "--tab", "1", "--json", j])
        js = json.loads(out)[0]["input"]["text"] if cod == 0 else ""
        seguro = ("GetDados" in js and "SalvarEntradaDados" not in js and ".click()" not in js
                  and "replace(/[?&=]/g" in js)
        sint = True
        if node and js:
            f = os.path.join(tmp, "r.mjs")
            open(f, "w", encoding="utf-8").write(js)
            sint = subprocess.run([node, "--check", f], capture_output=True, text=True).returncode == 0
        ok("lote rapido: só leitura, sem clique, saída sem '='", cod == 0 and seguro and sint)
        casos = (("lote leitura", ["lote", "leitura", "--tab", "1", "--periodo", "19", "--campos", "custo_real", "contingencia",
                                   "indice", "retencao", "a_pagar"], 0),
                 ("lote gravar", ["lote", "gravar", "--tab", "1", "--json", j, "--campos", "custo_real", "retencao", "a_pagar", "estoque"], 0),
                 ("lote gravar recusa campo sem valor", ["lote", "gravar", "--tab", "1", "--json", j, "--campos", "medicao"], 1),
                 ("lote gravar recusa tela de lista", ["lote", "gravar", "--tab", "1", "--json", j, "--campos", "contingencia"], 1),
                 ("lote gravar recusa nível 8", ["lote", "gravar", "--tab", "1", "--json", j, "--campos", "medicao_exclusiva_custo"], 1))
        for nome, args, esperado in casos:
            cod, out, err = rodar(args)
            if esperado:
                ok(nome, cod != 0, (out + err).strip()[-120:])
                continue
            falhas = []
            for linha in out.splitlines():
                if linha.startswith("###"):
                    continue
                itens = json.loads(linha)
                if len(itens) > 20:
                    falhas.append(f"lote com {len(itens)} itens")
                for it in itens:
                    if it["name"] != "javascript_tool":
                        continue
                    js = it["input"]["text"]
                    if "==='Salvar'" in js or "buttonChangeApiKey" in js:
                        falhas.append("JS procura Salvar pelo texto")
                    if node:
                        f = os.path.join(tmp, "t.mjs")
                        open(f, "w", encoding="utf-8").write(js)
                        r = subprocess.run([node, "--check", f], capture_output=True, text=True)
                        if r.returncode:
                            falhas.append("sintaxe: " + r.stderr.strip()[:80])
            if nome == "lote gravar" and ('const V=\\"874668,09\\"' not in out or "getElementById('salvar')" not in out):
                falhas.append("valor ou Salvar do painel ausente")
            ok(nome, cod == 0 and not falhas, "; ".join(falhas[:3]))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pasta", help="pasta com as planilhas de set/26 (custo e curva física)")
    a = ap.parse_args()
    ref = json.load(open(REF, encoding="utf-8"))
    teste_ajustes(ref)
    teste_escolha(ref)
    teste_js()
    teste_lote()
    if a.pasta:
        teste_extracao(ref, a.pasta)
    larg = max(len(c) for c, _, _ in resultados)
    for caso, passou, det in resultados:
        print(f"{'OK    ' if passou else 'FALHOU'} {caso:{larg}s}  {det}")
    falhas = sum(1 for _, p, _ in resultados if not p)
    print(f"\n{len(resultados) - falhas} de {len(resultados)} casos OK" + ("" if a.pasta else " (extração não testada: use --pasta)"))
    sys.exit(1 if falhas else 0)


if __name__ == "__main__":
    main()
