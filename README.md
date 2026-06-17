# Dashboard de Conciliação de Estoque — Grupo Minas Port

Painel que concilia, por **número de pedido** e por **navio**, as três etapas do
processo de estoque: **Retorno → Industrialização → Venda**. Identifica
automaticamente os pedidos que **não fecham** (quantidades divergentes entre
etapas), os **pendentes** (entraram mas não foram vendidos) e as **vendas
diretas** (sem retorno/industrialização).

Atualiza sozinho a partir da planilha online, via GitHub Actions.

## Como a conciliação é feita

A operação de cada NF é identificada pela coluna **"Série da Ordem de Compra"** na aba ENTRADA:

| Série | Etapa | CFOPs típicos |
|-------|-------|---------------|
| `RET` | Retorno | 902, 906, 925 |
| `IND` | Industrialização | 124, 125 |
| (aba SAÍDA) | Venda | — |

Tipos de pedido detectados:

- **Completo** — tem Retorno + Industrialização + Venda → as três quantidades devem bater.
- **Sem industrialização** — tem Retorno + Venda (típico da Filial 6 / Sul Norte / Imbituba) → Retorno deve bater com Venda.
- **Venda direta** — só Venda.

Um pedido fica **Divergente** quando a diferença entre etapas passa de
`0,01 t` (tolerância para arredondamento de NF). Ex.: pedido 8344 — Retorno e
Industrialização batem em 6.227,23 t, mas a Venda soma 6.292,07 t (diferença de
64,84 t), sinalizando NF de venda com quantidade incorreta.

## Arquivos

| Arquivo | Função |
|---------|--------|
| `index.html` | O dashboard (abre sozinho no navegador, lê o `dados.json`). |
| `build_data.py` | Lê a planilha `.xlsx` e gera o `dados.json`. |
| `dados.json` | Dados já conciliados que o dashboard consome. |
| `requirements.txt` | Dependência Python (openpyxl). |
| `.github/workflows/atualizar.yml` | Automação que baixa a planilha e regenera os dados. |

## Publicação (passo a passo)

1. **Deixe a planilha pública para leitura**: no Google Sheets → *Compartilhar* →
   "Qualquer pessoa com o link" → *Leitor*. (É o único requisito para o download
   automático funcionar.)
2. Crie um repositório no GitHub e suba todos estes arquivos.
3. Em **Settings → Pages**, defina *Source = Deploy from a branch*, branch `main`, pasta `/ (root)`.
4. Em **Settings → Actions → General → Workflow permissions**, marque
   *Read and write permissions* (para o robô poder commitar o `dados.json`).
5. O dashboard ficará em `https://<seu-usuario>.github.io/<repo>/`.

## Atualização automática

O workflow roda a cada 30 minutos (ajustável no `cron`), baixa a versão mais
recente da planilha (`export?format=xlsx`), regera o `dados.json` e faz commit
só quando algo muda. Você também pode rodar na hora pelo botão **Run workflow**
na aba *Actions*. O botão **↻ atualizar** no dashboard recarrega o `dados.json`
publicado.

> O ID da planilha está fixado no workflow (`SHEET_ID`). Se trocar de planilha,
> edite esse valor em `.github/workflows/atualizar.yml`.
