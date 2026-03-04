# media-report-automation

Projeto base robusto para automação de relatórios de mídia: ingestão de planilha, cálculo de KPIs de investimento, geração de visualizações e exportação de apresentação executiva em PowerPoint.

## Estrutura do projeto

```text
media-report-automation/
├── data/
├── scripts/
│   ├── analyze_campaign.py
│   └── generate_mock_campaign.py
├── output/
└── README.md
```

## Objetivo

Apoiar rotinas de análise para times de mídia/performance, criando um fluxo reproduzível para:

1. Consolidar investimento por canal.
2. Medir concentração e distribuição de verba.
3. Gerar artefatos executivos (gráficos + PPT).
4. Disponibilizar outputs tabulares para reuso em BI ou outras automações.

## Requisitos

- Python 3.9+
- Bibliotecas:
  - `matplotlib`
  - `python-pptx`
  - `openpyxl` (necessário para entrada/saída `.xlsx`)

Instalação sugerida:

```bash
pip install matplotlib python-pptx openpyxl
```

## Formato da base de entrada

Colunas obrigatórias:

- `Canal`
- `Investimento`

Formato suportado:

- `.csv` (recomendado para simulação rápida)
- `.xlsx` (quando disponível)

## Execução com dados fictícios (exploração antes dos dados reais)

### 1) Gerar base mock automaticamente

```bash
python scripts/generate_mock_campaign.py
```

Arquivos gerados:

- `data/campanha_ficticia.csv`
- `data/campanha_ficticia.xlsx` (se `openpyxl` estiver disponível)

### 2) Rodar análise automatizada com a base mock

```bash
python scripts/analyze_campaign.py \
  --input data/campanha_ficticia.csv \
  --output-dir output \
  --title "Share de Investimento - Campanha Fictícia" \
  --campaign "Campanha Potência (simulação)" \
  --period "Q1/2026"
```

## Execução com dados reais

```bash
python scripts/analyze_campaign.py --input data/campanha_real.xlsx --sheet "Resumo"
```

## Parâmetros disponíveis

### `scripts/analyze_campaign.py`

- `--input`: caminho do arquivo de entrada (`.csv` ou `.xlsx`) (obrigatório)
- `--sheet`: nome da aba quando a entrada for `.xlsx` (padrão = `Resumo`)
- `--output-dir`: diretório de saída (padrão = `output`)
- `--title`: título principal dos gráficos e relatório
- `--campaign`: nome da campanha para o slide executivo
- `--period`: período de análise
- `--log-level`: nível de log (`DEBUG`, `INFO`, `WARNING`, `ERROR`)

### `scripts/generate_mock_campaign.py`

- `--output-csv`: caminho do CSV fictício (padrão: `data/campanha_ficticia.csv`)
- `--output-xlsx`: caminho do XLSX fictício (padrão: `data/campanha_ficticia.xlsx`)
- `--sheet`: nome da aba no XLSX (padrão: `Resumo`)
- `--skip-xlsx`: gera apenas CSV

## KPIs calculados

- Investimento total
- Número de canais ativos
- Investimento médio por canal
- Canal líder (valor absoluto e share)
- Índice HHI de concentração
- Share e share acumulado por canal

## Saídas geradas em `output/`

- `base_processada.csv` → dados limpos para auditoria/reprocessamento
- `ranking_canais.csv` → consolidação por canal com shares e acumulados
- `kpis_resumo.json` → KPIs em formato estruturado
- `kpis_resumo.csv` → KPIs em formato tabular
- `investment_share_pie.png` → gráfico de pizza
- `investment_share_bar.png` → gráfico de barras de investimento
- `investment_share_report.pptx` → apresentação executiva com KPIs e gráficos
