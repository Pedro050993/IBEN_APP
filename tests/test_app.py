"""Teste da interface com o AppTest do Streamlit (sem navegador)."""
import os

import pytest

st_testing = pytest.importorskip("streamlit.testing.v1")
from tests.fixtures import analise_custos, memoria_xlsx, pasta_compras  # noqa: E402

APP = os.path.join(os.path.dirname(__file__), "..", "iben_app", "app.py")


def _app(tmp_path):
    at = st_testing.AppTest.from_file(APP, default_timeout=60)
    at.run()
    at.sidebar.text_input(key="saida").set_value(str(tmp_path / "saidas")).run()
    return at


def _pagina(at, nome):
    at.sidebar.radio[0].set_value(nome).run()
    assert not at.exception, at.exception
    return at


def test_todas_as_paginas_abrem(tmp_path):
    at = _app(tmp_path)
    for p in ["Início", "Extração de OCs", "iGile do mês", "E-mail de fechamento", "Auditoria da análise"]:
        _pagina(at, p)


def test_link_web_explica_o_que_fazer(tmp_path):
    at = _pagina(_app(tmp_path), "Extração de OCs")
    at.text_input(key="oc_pasta").set_value("https://iben.sharepoint.com/sites/x").run()
    assert any("Sincronize" in e.value for e in at.error)


def test_fluxo_oc(tmp_path):
    pasta_compras(str(tmp_path / "compras"))
    os.makedirs(tmp_path / "custo")
    memoria_xlsx(str(tmp_path / "custo" / "Análise de Custos - VF - Out_26_R01.xlsx"))
    at = _pagina(_app(tmp_path), "Extração de OCs")
    at.text_input(key="oc_pasta").set_value(str(tmp_path / "compras"))
    at.text_input(key="oc_mem").set_value(str(tmp_path / "custo")).run()
    at.button[0].click().run()
    assert not at.exception, at.exception
    assert any("Obra 400.137" in s.value for s in at.success)
    saidas = [f for r, _, fs in os.walk(tmp_path / "saidas") for f in fs]
    assert any(f.startswith("OCs 400.137") for f in saidas)


def test_fluxo_igile_e_auditoria(tmp_path):
    analise_custos(str(tmp_path / "custo"), nome="Análise de Custos - VF - Out_26_R01.xlsx")
    at = _pagina(_app(tmp_path), "iGile do mês")
    at.text_input(key="ig_custo").set_value(str(tmp_path / "custo")).run()
    at.button[0].click().run()
    assert not at.exception, at.exception
    assert any("2026-10" in m.value for m in at.markdown)
    at = _pagina(at, "Auditoria da análise")
    at.text_input(key="au_arq").set_value(str(tmp_path / "custo")).run()
    at.button[0].click().run()
    assert not at.exception, at.exception
