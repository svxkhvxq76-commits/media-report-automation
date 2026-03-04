# media-report-automation

Projeto base robusto para automação de relatórios de mídia: ingestão de planilha Excel, cálculo de KPIs de investimento, geração de visualizações e exportação de apresentação executiva em PowerPoint.

## Estrutura do projeto

```text
media-report-automation/
├── data/
├── scripts/
│   └── analyze_campaign.py
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
  - `pandas`
  - `matplotlib`
  - `python-pptx`
  - `openpyxl`

Instalação sugerida:

```bash
pip install pandas matplotlib python-pptx openpyxl
```

## Formato da planilha de entrada

Arquivo Excel (`.xlsx`) com colunas obrigatórias:

- `Canal`
- `Investimento`

Exemplo:

| Canal    | Investimento |
|----------|--------------|
| TV       | 150000       |
| Search   | 80000        |
| Social   | 120000       |

## Execução

Execução mínima:

```bash
python scripts/analyze_campaign.py --input data/campanha.xlsx
```

Execução recomendada (contexto executivo):

```bash
python scripts/analyze_campaign.py \
  --input data/campanha.xlsx \
  --sheet "Resumo" \
  --output-dir output \
  --title "Share de Investimento por Canal" \
  --campaign "Lançamento Q1" \
  --period "Jan-Mar/2026"
```

## Parâmetros disponíveis

- `--input`: caminho do arquivo Excel (obrigatório)
- `--sheet`: nome da aba (opcional; padrão = primeira)
- `--output-dir`: diretório de saída (padrão = `output`)
- `--title`: título principal dos gráficos e relatório
- `--campaign`: nome da campanha para o slide executivo
- `--period`: período de análise
- `--log-level`: nível de log (`DEBUG`, `INFO`, `WARNING`, `ERROR`)

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

## Próximos passos sugeridos

- Adicionar integração com APIs de mídia (Meta, Google Ads, DV360 etc.) para coleta automática.
- Incluir comparação de períodos (MoM, QoQ).
- Incorporar métricas de performance (CPM, CPC, CPA, ROAS) quando disponíveis.
- Agendar execução recorrente via CI/CD ou orquestrador (Airflow, cron, GitHub Actions).
