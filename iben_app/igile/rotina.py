"""Preparacao do lancamento mensal do iGile: escolhe os arquivos do mes, extrai os valores com o
igile.py (testado contra set/26) e gera o Excel 'valores para lancar'.

O app nao abre o iGile e nao grava nada: o engenheiro lanca e confere. Os ajustes de equilibrio
do OPR dependem da Variacao Termino lida no relatorio e ficam na calculadora da tela.
"""
import json
import os
import subprocess
import sys
import tempfile

from iben_app.comum import listar_arquivos, norm
from iben_app.igile import igile

SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "igile.py")

OBRAS = {
    "Viva Parque Unique": {"id": "1.44", "custo": "02 - Viva Parque Unique", "prazo": "02 - Viva Parque Unique"},
    "Viva Parque Altus": {"id": "1.36", "custo": "03 - Viva Parque Altus", "prazo": "Altus"},
    "Villa Flora II - Condomínio das Azaleias": {"id": "1.38", "custo": "01 - Villa Flora II - Condomínio das Azaleias",
                                                 "prazo": "01 - Vila Flora"},
    "Viva e Realize Campinas - Dom Pedro II": {"id": "1.45", "custo": "06 - Viva e Realize Campinas - Dom Pedro II",
                                               "prazo": "Dom Pedro"},
    "Veccon Kronos - Loteamento": {"id": "1.41", "custo": "04 - Veccon Kronos - Loteamento",
                                   "prazo": "05 - Veccon Kronos - Casas"},
    "Veccon Kronos - Casas": {"id": "1.42", "custo": "05 - Veccon Kronos - Casas", "prazo": "05 - Veccon Kronos - Casas"},
}
IGNORAR_CUSTO = ("comparativo", "orcamento")


def escolher_arquivos(pasta_custo, pasta_prazo=None):
    """Planilha de custos (ultimo mes, maior revisao) e, do mesmo mes, curva, PLS e efetivo."""
    out = {"alertas": [], "candidatos": {}}
    nomes = [a["arquivo"] for a in listar_arquivos(pasta_custo, (".xlsm", ".xlsx"))
             if not any(x in norm(a["nome"]) for x in IGNORAR_CUSTO)]
    r = igile.escolher(nomes, "custo")
    out["candidatos"]["custo"] = r["candidatos"]
    if not r["escolha"]:
        out["alertas"].append(f"planilha de custos: {r.get('alerta')}")
        out["custo"] = None
        return out
    out["custo"] = r["escolha"]["arquivo"]
    out["mes"] = r["escolha"]["mes"]
    if r["escolha"]["copia"]:
        out["alertas"].append("planilha de custos com 'Cópia de' no nome: confirmar que é a versão certa")
    if pasta_prazo:
        todos = [a["arquivo"] for a in listar_arquivos(pasta_prazo, (".xlsm", ".xlsx"))]
        for tipo in ("curva", "pls", "efetivo"):
            rr = igile.escolher(todos, tipo)
            out["candidatos"][tipo] = rr["candidatos"]
            esc = rr["escolha"]
            if esc and esc["mes"] == out["mes"]:
                out[tipo] = esc["arquivo"]
            else:
                out[tipo] = None
                motivo = rr.get("alerta") or (f"a mais recente é de {esc['mes']}" if esc else "nenhum arquivo")
                out["alertas"].append(f"{tipo}: sem arquivo de {out['mes']} ({motivo}); campos ficam pendentes")
    return out


def extrair(obra, arquivos, pasta_saida, efetivo_manual=None):
    """Roda o igile.py (mesma versao testada) e devolve o JSON da obra + saida do script."""
    os.makedirs(pasta_saida, exist_ok=True)
    js = os.path.join(pasta_saida, f"{norm(obra).replace(' ', '_')}_{arquivos['mes']}.json")
    cmd = [sys.executable, "-I", SCRIPT, "--obra", obra, "--mes", arquivos["mes"], "--custo-mes", arquivos["custo"],
           "--saida", js]
    for k, flag in (("curva", "--curva"), ("pls", "--pls")):
        if arquivos.get(k):
            cmd += [flag, arquivos[k]]
    if efetivo_manual is not None:
        cmd += ["--efetivo-manual", str(int(efetivo_manual))]
    elif arquivos.get("efetivo"):
        cmd += ["--efetivo", arquivos["efetivo"]]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if not os.path.exists(js):
        raise RuntimeError(p.stderr.strip() or p.stdout.strip() or "igile.py não gerou o JSON")
    d = json.load(open(js, encoding="utf-8"))
    d["_saida_script"] = p.stdout
    d["_parada"] = p.returncode == 2
    return js, d


def gerar_log(jsons, saida):
    p = subprocess.run([sys.executable, "-I", SCRIPT, "log", *jsons, "--saida", saida],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip())
    return saida


def ajuste(etapa, json_obra, gerado, **atuais):
    """Calculadora dos ajustes de equilibrio do OPR (mesma formula do igile.py)."""
    args = [sys.executable, "-I", SCRIPT, "ajuste", "--etapa", etapa, "--json", json_obra, "--gerado", str(gerado)]
    for k, v in atuais.items():
        if v is not None:
            args += [f"--{k.replace('_', '-')}", str(v)]
    p = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return p.returncode, (p.stdout + p.stderr).strip()


def tabela_lancamento(d):
    """Linhas campo / valor no formato do iGile (sem milhar, virgula) / origem / alerta."""
    nomes = {"custo_real": "Custo Real", "contingencia": "Contingência de Custo", "retencao": "Retenção",
             "a_pagar": "Á Pagar", "diferenca": "Diferença (para os ajustes)", "medicao": "Medição (%)",
             "avanco_fisico_acumulado": "Avanço físico acumulado (%)",
             "medicao_exclusiva_custo": "Medição exclusiva de custo (nível 8, total)",
             "realizado_banco": "Realizado Banco (%)", "realizado_recursos": "Realizado de Recursos",
             "mes_analise": "MÊS ANÁLISE (conferência)"}
    out = []
    for k, v in d["campos"].items():
        val = v.get("valor")
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            casas = 0 if k == "realizado_recursos" else 2
            digitar = f"{val:.{casas}f}".replace(".", ",")
        else:
            digitar = "" if val is None else str(val)
        out.append({"Tela no iGile": nomes.get(k, k), "Valor": val, "Digitar": digitar,
                    "Origem": v.get("origem"), "Alerta": v.get("alerta", ""),
                    "Situação": "pendente" if val is None else ("conferir" if v.get("alerta") else "ok")})
    return out
