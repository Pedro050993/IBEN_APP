import datetime as dt
import email
import json
import os

import pytest

from iben_app.fechamento import email as fe
from iben_app.fechamento import leitura
from tests.fixtures import analise_custos

CFG = {"empreendimento": "VILLA FLORA II", "obras": {"VILLA FLORA II": {}},
       "para": ["obra@iben.eng.br", "cliente@veccon.com.br"], "cc": ["obra@iben.eng.br", "dir@iben.eng.br"]}


@pytest.fixture(scope="module")
def planilha(tmp_path_factory):
    return analise_custos(str(tmp_path_factory.mktemp("ac")))


def test_leitura_bloco(planilha):
    af = leitura.analise_financeira(planilha)
    assert af["mes"] == "2026-10"
    assert af["total"]["var"] == -4800 and af["total"]["faixa"] == (10, 16)
    no_total = {l["eap"] for l in af["linhas"] if l["no_total"]}
    assert no_total == {"1.1", "1.2", "2.1", "2.2", "2.3"}, "pais (SUBTOTAL) e item fora da faixa nao entram"
    assert not next(l for l in af["linhas"] if l["eap"] == "5.13.3")["no_total"]


def test_desvios_tira_variacao_zero(planilha):
    dv = leitura.desvios(planilha)
    assert [ln["valores"][0] for ln in dv["linhas"]] == ["2.1", "1.2", "1.1", "2.2"]
    assert dv["total_diferenca"] == 32000
    assert dv["cabecalho"][0]["fundo"] == "#1F3864"


def test_resumo_e_alertas(planilha):
    m = fe.montar(CFG, {"VILLA FLORA II": planilha}, "2026-10", saudacao="boa tarde")
    assert m["assunto"] == "VILLA FLORA II | ANÁLISE CUSTO (FECHAMENTO) - OUTUBRO/26"
    assert m["cc"] == ["dir@iben.eng.br"], "repetido em Para e Cc fica so no Para"
    n = m["conferencia"]["VILLA FLORA II"]["numeros"]
    assert n["incc"] == -4000 and n["mudancas"] == -800
    texto = m["html"]
    assert "variação negativa de -R$ 4.800,00" in texto
    assert "EAP 1.1 Canteiro: -R$ 3.800,00 (Aumento do prazo do canteiro e locação de container)" in texto
    assert "EAP 2.2 Alvenaria: R$ 3.000,00 (Novo pedido OC 400.137.2.00576 abaixo da previsão)" in texto
    assert "Item fora do resultado, EAP 5.13.3" in texto
    assert any("fora da Iben" in a for a in m["alertas"])
    assert not any("não fecha" in a for a in m["alertas"]), m["alertas"]


def test_caixa_normal():
    assert fe.caixa_normal("AUMENTO DO PRAZO DO CANTEIRO") == "Aumento do prazo do canteiro"
    assert fe.caixa_normal("SM DO MURO EXTERNO EAP 5.13.3") == "SM do muro externo EAP 5.13.3"
    assert fe.caixa_normal("Já em caixa normal") == "Já em caixa normal"


def test_para_obra_com_mes_errado(planilha):
    with pytest.raises(fe.ObraParada, match="MÊS ANÁLISE"):
        fe.montar(CFG, {"VILLA FLORA II": planilha}, "2026-11")


def test_recusa_pendente(planilha):
    with pytest.raises(fe.ObraParada, match="pendentes"):
        fe.montar(dict(CFG, pendente=True), {"VILLA FLORA II": planilha}, "2026-10")


def test_eml(planilha, tmp_path):
    m = fe.montar(CFG, {"VILLA FLORA II": planilha}, "2026-10")
    out = fe.salvar(m, str(tmp_path))
    msg = email.message_from_bytes(open(out["eml"], "rb").read())
    assert msg["X-Unsent"] == "1"
    assert msg["Subject"].startswith("VILLA FLORA II")
    nomes = [p.get_filename() for p in msg.walk() if p.get_filename()]
    assert nomes == [os.path.basename(planilha)]
    assert json.load(open(out["conferencia"]))["mes"] == "2026-10"
    assert "Assunto:" in open(out["previa"], encoding="utf-8").read()
