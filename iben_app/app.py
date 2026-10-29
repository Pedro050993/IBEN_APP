"""Interface do Iben App (rodar com: streamlit run iben_app/app.py, ou iniciar.bat).

Roda no computador do engenheiro e le as pastas que ele informar (Drive ou SharePoint
sincronizado). Nada e gravado no iGile, no SharePoint, no TOTVS ou no Outlook: o app gera
arquivos numa pasta de saida e o engenheiro confere e usa.
"""
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402
import streamlit as st  # noqa: E402

from iben_app.comum import CaminhoInvalido, br, resolver_caminho, ultima_planilha  # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG = os.path.join(RAIZ, "config", "destinatarios.json")

st.set_page_config(page_title="Iben App", page_icon="📐", layout="wide")


def pasta_saida(sub):
    base = st.session_state.get("saida") or os.path.join(os.path.expanduser("~"), "Documents", "Iben App")
    p = os.path.join(base, f"{dt.date.today():%Y-%m-%d} {sub}")
    os.makedirs(p, exist_ok=True)
    return p


def caminho(label, chave, ajuda=None):
    txt = st.text_input(label, key=chave, help=ajuda,
                        placeholder=r"Cole o caminho do Explorador de Arquivos, ex.: G:\Meu Drive\01 - OBRAS\...")
    if not txt:
        return None
    try:
        return resolver_caminho(txt)
    except CaminhoInvalido as e:
        st.error(str(e))
        return None


def baixar(path, rotulo):
    with open(path, "rb") as f:
        st.download_button(rotulo, f.read(), file_name=os.path.basename(path), key=f"dl-{path}")


def alertas(lista):
    for a in lista:
        st.warning(a)


# ---------------- paginas ----------------

def inicio():
    st.title("Iben App")
    st.write("Rotinas mensais de custo e suprimentos, sem IA. Cada tela pede o caminho da pasta "
             "e entrega um arquivo para você conferir.")
    st.markdown(
        "| Tela | Entrada | Saída |\n|---|---|---|\n"
        "| Extração de OCs | pasta *1. PROCESSO DE COMPRAS* da obra + Análise de Custos | Excel com itens, status e cruzamento com a Memória |\n"
        "| iGile do mês | pasta de custo da obra (+ pasta de prazo) | valores prontos para digitar no iGile + calculadora dos ajustes do OPR |\n"
        "| E-mail de fechamento | pasta de custo de cada obra do e-mail | rascunho .eml para abrir no Outlook, prévia e conferência |\n"
        "| Auditoria da análise | Análise de Custos do mês | Excel com os erros de estrutura da planilha |")
    st.info("Link do SharePoint não é lido direto. Sincronize a pasta no OneDrive e cole o caminho do "
            "Explorador de Arquivos.")


def tela_oc():
    from iben_app.oc import cruzamento, excel, extracao
    st.header("Extração de OCs do Processo de Compras")
    pasta = caminho("Pasta 1. PROCESSO DE COMPRAS da obra", "oc_pasta")
    analise = caminho("Análise de Custos (arquivo ou pasta de custo da obra; vale a última revisão)", "oc_mem",
                      "Se informar a pasta, o app escolhe o último mês e a maior revisão pelo nome do arquivo.")
    c1, c2 = st.columns(2)
    obra = c1.text_input("Código da obra (ex.: 137). Em branco = detectar pelo nome das pastas", key="oc_obra")
    janela = c2.number_input("Janela (meses, pela data de modificação da pasta)", 1, 60, 3)
    if not st.button("Extrair", type="primary", disabled=not pasta):
        return
    barra = st.progress(0.0, "lendo pastas")
    try:
        cod, procs = extracao.extrair_pasta(pasta, obra.strip() or None, janela,
                                            progresso=lambda i, n, nome: barra.progress((i + 1) / max(n, 1), nome[:60]))
    except Exception as e:
        st.error(str(e))
        return
    mem, cz = None, []
    if analise:
        esc = ultima_planilha(analise, "custo", ignorar=("comparativo", "orcamento"))
        if not esc["escolha"]:
            st.error(f"Análise de Custos: {esc.get('alerta')}")
        else:
            arq = esc["escolha"]["arquivo"]
            mes = esc["escolha"]["mes"] or "não identificado no nome"
            st.caption(f"Memória: {os.path.basename(arq)} (mês {mes}, R{esc['escolha']['revisao']:02d})")
            mem = cruzamento.ler_memoria(arq)
            cz = cruzamento.cruzar(procs, mem, cod)
    else:
        st.warning("Sem Análise de Custos: o cruzamento com a Memória não roda e a Tarefa fica vazia.")
    saida = os.path.join(pasta_saida("OCs"), f"OCs 400.{cod} {dt.date.today():%Y-%m-%d}.xlsx")
    excel.gerar(cod, procs, cz, saida, mem, pasta)
    barra.empty()
    st.success(f"Obra 400.{cod}: {len(procs)} pastas na janela, "
               f"{sum(len(p.get('itens', [])) for p in procs)} itens. Arquivo: {saida}")
    cont = pd.Series([p.get("status") for p in procs]).value_counts()
    met = pd.Series([c["metodo"] for c in cz]).value_counts() if cz else pd.Series(dtype=int)
    a, b = st.columns(2)
    a.subheader("Status das pastas")
    a.dataframe(cont.rename("pastas"))
    b.subheader("Cruzamento com a Memória")
    b.dataframe(met.rename("OCs"))
    if cz:
        st.subheader("Maiores diferenças de valor (levar ao engenheiro da obra)")
        df = pd.DataFrame([{"Nº OC": c["pr"].get("n_oc"), "Pasta": c["pr"]["pasta"], "Método": c["metodo"],
                            "Confiança": c["confianca"], "EAP": "; ".join(c["eaps"]),
                            "Diferença (R$)": c["diferenca"], "Situação": c["situacao"]} for c in cz])
        df = df.reindex(df["Diferença (R$)"].abs().sort_values(ascending=False, na_position="last").index)
        st.dataframe(df.head(15), width="stretch", hide_index=True)
    verificar = [p for p in procs if p.get("status") == "Verificar"]
    if verificar:
        with st.expander(f"{len(verificar)} pasta(s) para verificar"):
            for p in verificar:
                st.write(f"**{p['pasta']}**: " + " | ".join(p.get("obs", [])))
    baixar(saida, "Baixar Excel")


def tela_igile():
    from iben_app.igile import rotina
    st.header("iGile do mês: valores para lançar")
    st.caption("O app não abre o iGile. Ele lê as planilhas e entrega o valor no formato para digitar, "
               "com a origem de cada número. Quem lança e confere o ícone branco é você.")
    obra = st.selectbox("Obra (nome no iGile)", list(rotina.OBRAS))
    info = rotina.OBRAS[obra]
    st.caption(f"ID no iGile {info['id']} | pasta de custo '{info['custo']}' | pasta de prazo '{info['prazo']}'")
    custo = caminho("Pasta de custo da obra (dentro de 02 - Custo)", "ig_custo")
    prazo = caminho("Pasta de prazo da obra (Curva Física, PLS e efetivo). Opcional", "ig_prazo")
    efm = st.number_input("Efetivo médio do mês (se não houver planilha de efetivo; 0 = não informar)", 0, 5000, 0)
    if st.button("Ler planilhas", type="primary", disabled=not custo):
        arq = rotina.escolher_arquivos(custo, prazo)
        if not arq.get("custo"):
            alertas(arq["alertas"])
            return
        st.write(f"Período **{arq['mes']}** pela planilha **{os.path.basename(arq['custo'])}**")
        alertas(arq["alertas"])
        try:
            js, d = rotina.extrair(obra, arq, pasta_saida("iGile"), efm or None)
        except RuntimeError as e:
            st.error(str(e))
            return
        if d["_parada"]:
            st.error("PARADA: o layout da planilha mudou ou o arquivo é de outro mês. Não lance os campos abaixo "
                     "antes de ajustar: " + "; ".join(d.get("paradas_layout", [])))
        st.session_state["ig_json"] = js
        st.session_state["ig_dados"] = d
    d = st.session_state.get("ig_dados")
    if not d:
        return
    df = pd.DataFrame(rotina.tabela_lancamento(d))
    st.dataframe(df, width="stretch", hide_index=True,
                 column_config={"Valor": st.column_config.NumberColumn(format="%.2f")})
    st.info("Regras: período fechado no iGile não se edita; período com valor diferente da planilha só se "
            "sobrescreve depois de confirmar qual revisão vale; índice do período M é o de M−1, com 4 casas.")
    log = os.path.join(os.path.dirname(st.session_state["ig_json"]), f"Log iGile {d['mes']} {d['obra'][:30]}.xlsx")
    rotina.gerar_log([st.session_state["ig_json"]], log)
    baixar(log, "Baixar log em Excel")

    st.subheader("Ajustes de equilíbrio do OPR")
    st.caption("Depois de lançar e atualizar o relatório, leia a Variação Término no One Page Report. "
               "Os valores calculados aqui são de equilíbrio, não são dados reais.")
    av = (d["campos"].get("avanco_fisico_acumulado") or {}).get("exato") or (d["campos"].get("avanco_fisico_acumulado") or {}).get("valor")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Previsto** (ÍNDICE IDC PREVISTO)")
        g = st.number_input("Variação Término gerada (R$)", value=0.0, format="%.2f", key="g1")
        acum = st.number_input("Avanço físico acumulado (fração, ex. 0,4603)", value=float(av or 0) / (100 if (av or 0) > 1 else 1),
                               format="%.6f", help="Sem a curva do mês, use a última curva: o ajuste fica provisório.")
        est = st.number_input("Estoque atual", value=0.0, format="%.2f")
        adi = st.number_input("Adiantamento atual", value=0.0, format="%.2f")
        agr = st.number_input("Agregado não Pago atual", value=0.0, format="%.2f")
        if st.button("Calcular previsto"):
            rc, txt = rotina.ajuste("previsto", st.session_state["ig_json"], g, acumulado=acum or None,
                                    estoque_atual=est, adiantamento_atual=adi, agregado_atual=agr)
            (st.code if rc == 0 else st.error)(txt)
    with c2:
        st.markdown("**Previsão** (ÍNDICE IDC PREVISÃO)")
        g2 = st.number_input("Variação Término gerada (R$)", value=0.0, format="%.2f", key="g2")
        prev = st.number_input("Previsão de Despesas atual", value=0.0, format="%.2f")
        if st.button("Calcular previsão"):
            rc, txt = rotina.ajuste("previsao", st.session_state["ig_json"], g2, previsao_atual=prev)
            (st.code if rc == 0 else st.error)(txt)


def tela_email():
    from iben_app.fechamento import email as fe
    st.header("E-mail de fechamento da análise de custo")
    try:
        cfg = json.load(open(CONFIG, encoding="utf-8"))
    except (OSError, ValueError) as e:
        st.error(f"config/destinatarios.json: {e}")
        return
    nome = st.selectbox("E-mail", list(cfg["emails"]))
    bloco = cfg["emails"][nome]
    if bloco.get("pendente"):
        st.error("Destinatários pendentes: preencha Para e Cc em config/destinatarios.json e tire o \"pendente\".")
    hoje = dt.date.today()
    mes = st.text_input("Mês de competência (AAAA-MM)", value=f"{hoje:%Y-%m}")
    planilhas = {}
    for obra, oc in bloco["obras"].items():
        p = caminho(f"{obra}: pasta de custo ou arquivo da análise", f"em_{nome}_{obra}",
                    f"Padrão: {oc.get('pasta', '')}")
        if p:
            esc = ultima_planilha(p, "custo", extensoes=(".xlsm",), ignorar=("comparativo", "orcamento"))
            if not esc["escolha"]:
                st.error(f"{obra}: {esc.get('alerta')}")
            else:
                planilhas[obra] = esc["escolha"]["arquivo"]
                st.caption(f"{obra}: {os.path.basename(esc['escolha']['arquivo'])}")
    c1, c2 = st.columns(2)
    sauda = c1.selectbox("Saudação", ["automática", "bom dia", "boa tarde"])
    abertura = c2.text_area("Parágrafo específico deste e-mail (opcional)")
    pronto = len(planilhas) == len(bloco["obras"]) and not bloco.get("pendente")
    if not st.button("Montar e-mail", type="primary", disabled=not pronto):
        return
    try:
        m = fe.montar(bloco, planilhas, mes.strip(), None if sauda == "automática" else sauda,
                      abertura.strip() or None, cfg.get("assinatura"))
    except fe.ObraParada as e:
        st.error(str(e))
        return
    except Exception as e:
        st.error(f"Não consegui ler a planilha: {e}")
        return
    out = fe.salvar(m, pasta_saida(f"E-mail {nome}"))
    st.success("Rascunho pronto. Abra o .eml no Outlook (abre como mensagem nova), confira e envie você.")
    alertas(m["alertas"])
    st.write(f"**Assunto:** {m['assunto']}  \n**Para:** {'; '.join(m['para'])}  \n**Cc:** {'; '.join(m['cc'])}")
    for obra, c in m["conferencia"].items():
        n = c["numeros"]
        st.write(f"{obra}: variação {br(n['variacao'])} | INCC {br(n['incc'])} | previsão e pedidos {br(n['mudancas'])}")
    st.components.v1.html(open(out["previa"], encoding="utf-8").read(), height=700, scrolling=True)
    baixar(out["eml"], "Baixar rascunho (.eml)")
    baixar(out["conferencia"], "Baixar conferência (.json)")


def tela_auditoria():
    from iben_app.custo import auditoria
    st.header("Auditoria da Análise de Custos")
    st.caption("Estrutura da planilha: valor digitado no lugar de fórmula, subtotal desalinhado, +L parcial, "
               "linhas sem EAP, links externos e checagens FALSO. A estimativa a 100% continua com você.")
    p = caminho("Análise de Custos (arquivo ou pasta de custo da obra)", "au_arq")
    if not st.button("Auditar", type="primary", disabled=not p):
        return
    esc = ultima_planilha(p, "custo", ignorar=("comparativo", "orcamento"))
    if not esc["escolha"]:
        st.error(esc.get("alerta"))
        return
    arq = esc["escolha"]["arquivo"]
    st.caption(os.path.basename(arq))
    ach = auditoria.auditar(arq)
    saida = os.path.join(pasta_saida("Auditoria"), f"Auditoria {os.path.splitext(os.path.basename(arq))[0]}.xlsx")
    auditoria.gerar_excel(arq, ach, saida)
    if ach:
        st.warning(f"{len(ach)} achado(s)")
        st.dataframe(pd.DataFrame(ach), width="stretch", hide_index=True)
    else:
        st.success("Nenhum achado de estrutura.")
    baixar(saida, "Baixar Excel")


PAGINAS = {"Início": inicio, "Extração de OCs": tela_oc, "iGile do mês": tela_igile,
           "E-mail de fechamento": tela_email, "Auditoria da análise": tela_auditoria}
escolha = st.sidebar.radio("Rotina", list(PAGINAS))
st.sidebar.text_input("Pasta onde o app grava as saídas", key="saida",
                      value=os.path.join(os.path.expanduser("~"), "Documents", "Iben App"))
st.sidebar.caption("Nada é gravado no iGile, SharePoint, TOTVS ou Outlook. O app só gera arquivos.")
PAGINAS[escolha]()
