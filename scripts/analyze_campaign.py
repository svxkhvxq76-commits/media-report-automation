#!/usr/bin/env python3
"""Pipeline robusto de análise de investimento por canal com exportação para PowerPoint."""

from __future__ import annotations

import argparse
import json
import logging
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt

REQUIRED_COLUMNS = ["Canal", "Investimento"]


@dataclass
class KPIBundle:
    total_investimento: float
    canais_ativos: int
    investimento_medio_por_canal: float
    top_canal: str
    top_canal_investimento: float
    top_canal_share: float
    hhi_concentracao: float


@dataclass
class ReportArtifacts:
    base_processada_csv: Path
    ranking_canais_csv: Path
    kpis_json: Path
    kpis_csv: Path
    grafico_pizza_png: Path
    grafico_barras_png: Path
    apresentacao_pptx: Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Lê uma planilha Excel com Canal e Investimento, calcula KPIs de distribuição de verba, "
            "gera visualizações e exporta um relatório em PowerPoint."
        )
    )
    parser.add_argument("--input", required=True, help="Caminho do arquivo Excel de entrada (.xlsx).")
    parser.add_argument("--sheet", default=None, help="Nome da aba a ser lida (padrão: primeira aba).")
    parser.add_argument("--output-dir", default="output", help="Diretório para arquivos gerados.")
    parser.add_argument("--title", default="Share de Investimento por Canal", help="Título principal do relatório.")
    parser.add_argument("--campaign", default="Campanha não informada", help="Nome da campanha para contextualização.")
    parser.add_argument("--period", default="Período não informado", help="Período da campanha (ex.: Jan/2026).")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return parser.parse_args()


def setup_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level),
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def load_and_validate_data(excel_path: Path, sheet_name: str | None) -> pd.DataFrame:
    logging.info("Lendo arquivo: %s", excel_path)
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes: {missing}")

    working = df[REQUIRED_COLUMNS].copy()
    working["Canal"] = working["Canal"].astype(str).str.strip()
    working["Investimento"] = pd.to_numeric(working["Investimento"], errors="coerce")
    working = working.dropna(subset=["Canal", "Investimento"])
    working = working[working["Canal"] != ""]

    if working.empty:
        raise ValueError("Não há dados válidos após limpeza da planilha.")

    if (working["Investimento"] < 0).any():
        raise ValueError("Foram encontrados valores negativos de investimento.")

    logging.info("Linhas válidas após limpeza: %s", len(working))
    return working


def build_channel_ranking(df: pd.DataFrame) -> pd.DataFrame:
    ranking = (
        df.groupby("Canal", as_index=False)["Investimento"]
        .sum()
        .sort_values("Investimento", ascending=False)
        .reset_index(drop=True)
    )

    total = ranking["Investimento"].sum()
    if total <= 0:
        raise ValueError("A soma total de investimento deve ser maior que zero.")

    ranking["Share"] = ranking["Investimento"] / total
    ranking["Share_Percentual"] = ranking["Share"] * 100
    ranking["Investimento_Acumulado"] = ranking["Investimento"].cumsum()
    ranking["Share_Acumulado"] = ranking["Share"].cumsum()
    return ranking


def compute_kpis(ranking: pd.DataFrame) -> KPIBundle:
    top_row = ranking.iloc[0]
    hhi = float((ranking["Share"] ** 2).sum())

    return KPIBundle(
        total_investimento=float(ranking["Investimento"].sum()),
        canais_ativos=int(ranking["Canal"].nunique()),
        investimento_medio_por_canal=float(ranking["Investimento"].mean()),
        top_canal=str(top_row["Canal"]),
        top_canal_investimento=float(top_row["Investimento"]),
        top_canal_share=float(top_row["Share"]),
        hhi_concentracao=hhi,
    )


def save_table_outputs(clean_df: pd.DataFrame, ranking_df: pd.DataFrame, kpis: KPIBundle, output_dir: Path) -> ReportArtifacts:
    base_processada_csv = output_dir / "base_processada.csv"
    ranking_canais_csv = output_dir / "ranking_canais.csv"
    kpis_json = output_dir / "kpis_resumo.json"
    kpis_csv = output_dir / "kpis_resumo.csv"

    clean_df.to_csv(base_processada_csv, index=False)
    ranking_df.to_csv(ranking_canais_csv, index=False)

    with kpis_json.open("w", encoding="utf-8") as fp:
        json.dump(asdict(kpis), fp, ensure_ascii=False, indent=2)

    pd.DataFrame([asdict(kpis)]).to_csv(kpis_csv, index=False)

    return ReportArtifacts(
        base_processada_csv=base_processada_csv,
        ranking_canais_csv=ranking_canais_csv,
        kpis_json=kpis_json,
        kpis_csv=kpis_csv,
        grafico_pizza_png=output_dir / "investment_share_pie.png",
        grafico_barras_png=output_dir / "investment_share_bar.png",
        apresentacao_pptx=output_dir / "investment_share_report.pptx",
    )


def save_pie_chart(ranking: pd.DataFrame, title: str, output_path: Path) -> None:
    fig, ax = plt.subplots(figsize=(10, 6))
    labels = [f"{canal} ({share:.1%})" for canal, share in zip(ranking["Canal"], ranking["Share"]) ]
    ax.pie(
        ranking["Investimento"],
        labels=labels,
        autopct="%1.1f%%",
        startangle=90,
        counterclock=False,
    )
    ax.set_title(title)
    ax.axis("equal")
    plt.tight_layout()
    fig.savefig(output_path, dpi=220)
    plt.close(fig)


def save_bar_chart(ranking: pd.DataFrame, title: str, output_path: Path) -> None:
    fig, ax = plt.subplots(figsize=(11, 6))
    ax.bar(ranking["Canal"], ranking["Investimento"], color="#4472C4")
    ax.set_title(f"{title} - Investimento Absoluto")
    ax.set_xlabel("Canal")
    ax.set_ylabel("Investimento")
    ax.tick_params(axis="x", rotation=30)

    for idx, value in enumerate(ranking["Investimento"]):
        ax.text(idx, value, f"{value:,.0f}", ha="center", va="bottom", fontsize=9)

    plt.tight_layout()
    fig.savefig(output_path, dpi=220)
    plt.close(fig)


def add_kpi_slide(presentation: Presentation, kpis: KPIBundle, campaign: str, period: str) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "Resumo Executivo de KPIs"

    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(12.0), Inches(4.8))
    tf = box.text_frame
    tf.word_wrap = True

    lines = [
        f"Campanha: {campaign}",
        f"Período: {period}",
        f"Total investido: R$ {kpis.total_investimento:,.2f}",
        f"Canais ativos: {kpis.canais_ativos}",
        f"Investimento médio por canal: R$ {kpis.investimento_medio_por_canal:,.2f}",
        f"Canal líder: {kpis.top_canal} (R$ {kpis.top_canal_investimento:,.2f} | {kpis.top_canal_share:.1%})",
        f"HHI de concentração: {kpis.hhi_concentracao:.3f}",
    ]

    for i, line in enumerate(lines):
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.text = line
        paragraph.font.size = Pt(20 if i == 0 else 16)


def add_chart_slide(presentation: Presentation, title: str, image_path: Path) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = title
    slide.shapes.add_picture(str(image_path), Inches(0.8), Inches(1.4), width=Inches(11.7))


def save_ppt(
    chart_pie_path: Path,
    chart_bar_path: Path,
    title: str,
    campaign: str,
    period: str,
    kpis: KPIBundle,
    output_path: Path,
) -> None:
    presentation = Presentation()
    add_kpi_slide(presentation, kpis, campaign, period)
    add_chart_slide(presentation, title, chart_pie_path)
    add_chart_slide(presentation, f"{title} - Visão Absoluta", chart_bar_path)
    presentation.save(output_path)


def main() -> None:
    args = parse_args()
    setup_logging(args.log_level)

    excel_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")

    clean_df = load_and_validate_data(excel_path, args.sheet)
    ranking_df = build_channel_ranking(clean_df)
    kpis = compute_kpis(ranking_df)
    artifacts = save_table_outputs(clean_df, ranking_df, kpis, output_dir)

    save_pie_chart(ranking_df, args.title, artifacts.grafico_pizza_png)
    save_bar_chart(ranking_df, args.title, artifacts.grafico_barras_png)
    save_ppt(
        artifacts.grafico_pizza_png,
        artifacts.grafico_barras_png,
        args.title,
        args.campaign,
        args.period,
        kpis,
        artifacts.apresentacao_pptx,
    )

    print("Análise concluída com sucesso.")
    print(f"Data de execução: {datetime.now().isoformat(timespec='seconds')}")
    print("Arquivos gerados:")
    for field_name, file_path in asdict(artifacts).items():
        print(f"- {field_name}: {file_path}")


if __name__ == "__main__":
    main()
