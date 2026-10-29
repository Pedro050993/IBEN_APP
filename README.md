# Iben App

App local, sem IA, para as rotinas mensais de custo e suprimentos dos engenheiros de planejamento da Iben. Ele roda no computador de quem usa, lê as pastas que a pessoa informar (Google Drive ou SharePoint sincronizado) e gera arquivos para conferência.

Ele não grava nada no iGile, no SharePoint, no TOTVS nem no Outlook. Quem lança, envia e confere é o engenheiro.

## O que tem hoje

| Tela | Você informa | O app entrega | Origem das regras |
|---|---|---|---|
| Extração de OCs | pasta `1. PROCESSO DE COMPRAS` da obra e a Análise de Custos | Excel com Resumo, itens por OC, Cruzamento memória e Pastas e status | skill `extracao-oc-processo-compras` |
| iGile do mês | pasta de custo da obra e, se houver, a pasta de prazo | valores no formato de digitação do iGile, com a origem de cada número, log em Excel e a calculadora dos ajustes do OPR | skill `atualizacao-igile` (o `igile.py` v2.4 está dentro do app, sem mudança de regra) |
| E-mail de fechamento | pasta de custo de cada obra do e-mail | rascunho `.eml` que abre no Outlook como mensagem nova, `previa.html` e `conferencia.json` | skill `email-fechamento-custo-iben` |
| Auditoria da análise | Análise de Custos do mês | Excel com os erros de estrutura da planilha | parte mecânica da skill `analise-custo-obra-iben` |

## Instalar (Windows)

1. Instale o Python 3.11 ou mais novo em python.org e marque **Add python.exe to PATH**.
2. Baixe esta pasta (GitHub, botão Code, Download ZIP) e descompacte.
3. Dê dois cliques em `iniciar.bat`. Na primeira vez ele instala as dependências; depois abre o app no navegador.

Para rodar pelo terminal:

```bash
pip install -r requirements.txt
```

```bash
streamlit run iben_app/app.py
```

## Como informar o caminho

Cole o caminho do Explorador de Arquivos, por exemplo `G:\Meu Drive\01 - OBRAS\02 - Custo\04 - Veccon Kronos - Loteamento`. Aspas e `file:///` são aceitos.

Link do SharePoint (`https://...sharepoint.com/...`) não é lido direto. Sincronize a pasta no OneDrive (Sincronizar, ou Adicionar atalho em Meus arquivos) e cole o caminho local. O app avisa quando recebe um link.

Quando você informa uma pasta de custo, o app escolhe a planilha pela regra do Pedro: último mês pelo nome do arquivo e maior revisão (R01, R02...). Empate no mesmo mês e revisão para a rotina e pede a sua escolha.

## Regras que o app segue

- Arquivo original nunca é regravado. Toda saída é um arquivo novo na pasta de saídas (por padrão `Documentos\Iben App\<data> <rotina>`).
- Extração de OCs: janela de 3 meses pela data de modificação da pasta; última revisão do PDF; cancelado, sem OC e fora do escopo com status próprio; validação qtde × unit × total, soma × SUBTOTAL × TOTAL; cruzamento em quatro camadas (Nº pedido, citação em texto, semelhança, sem correspondência). O formato de 10 dígitos `4001370xxx` nunca é juntado com a OC de mesma sequência.
- Semelhança é hipótese: a Tarefa sai em laranja com "validar" e a confirmação é do engenheiro.
- iGile: período = mês do nome da planilha; custo real = Incorrido N-1; trava de layout (código 2) bloqueia o campo em vez de seguir. Os ajustes do OPR são valores de equilíbrio, não dados reais.
- E-mail: só fatos da planilha. Mês da planilha diferente do fechamento para a obra. Destinatários pendentes no `config/destinatarios.json` bloqueiam o e-mail. O app nunca envia.

## Diferenças em relação às skills

- O app não opera navegador nem Outlook. No iGile ele entrega os valores e a calculadora; o lançamento é manual. No e-mail ele gera o `.eml`; você abre no Outlook, confere e envia.
- O print da aba Desvios vai no corpo do e-mail como tabela HTML com as cores da planilha, não como imagem.
- PDF escaneado sai como "Verificar" (o app não faz OCR).
- A estimativa a 100% (a parte de julgamento da `analise-custo-obra-iben`) não está no app. A auditoria cobre a estrutura da planilha.

## A validar antes de usar em produção

1. Rodar as quatro telas com os arquivos reais de uma obra (sugestão: Vila Flora II, que tem caso de regressão conhecido: OC 00351 → linha 1266, 00576 → linha 986, 00349 por Nº pedido) e conferir contra o resultado das skills.
2. Layout do bloco ANÁLISE PREVISÃO - 100% e da aba Desvios em cada obra: o app acha as colunas pelo título, mas os testes usam planilha sintética.
3. Layout dos PDFs de OC por obra (PMO.01.01, sem nº de item, formulário de locação): conferir as primeiras 3 a 5 OCs.
4. Preencher Para, Cc e a assinatura em `config/destinatarios.json` e copiar o logo e os selos para `config/assets/`.
5. Teste de regressão do iGile com as planilhas de set/26:

```bash
python -I tests/igile/teste_regressao.py --pasta "G:\Meu Drive\01 - OBRAS\02 - Custo"
```

## Estrutura

```
iben_app/
  app.py            interface (Streamlit)
  comum.py          caminho, escolha do arquivo, estilo dos Excel
  oc/               extração dos PDFs, cruzamento com a Memória, Excel
  igile/            igile.py (v2.4 da skill) e a rotina de preparação
  fechamento/       leitura da Análise Financeira e da aba Desvios, montagem do e-mail
  custo/            auditoria da Análise de Custos
config/destinatarios.json
tests/              testes com planilhas e PDFs sintéticos
```

Testes:

```bash
pip install -r requirements-dev.txt
```

```bash
python -m pytest -q tests
```
