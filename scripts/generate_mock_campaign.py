#!/usr/bin/env python3
"""Gera base fictícia de campanha para exploração do pipeline antes dos dados reais."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

DEFAULT_DATA = [
    {"Canal": "Search", "Investimento": 185000},
    {"Canal": "Social", "Investimento": 142500},
    {"Canal": "Programática", "Investimento": 96500},
    {"Canal": "Vídeo", "Investimento": 118000},
    {"Canal": "Display", "Investimento": 61000},
    {"Canal": "Influenciadores", "Investimento": 47000},
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera arquivos fictícios (.csv e opcionalmente .xlsx).")
    parser.add_argument("--output-csv", default="data/campanha_ficticia.csv", help="Caminho do CSV de saída.")
    parser.add_argument("--output-xlsx", default="data/campanha_ficticia.xlsx", help="Caminho do XLSX de saída.")
    parser.add_argument("--sheet", default="Resumo", help="Nome da aba para saída .xlsx.")
    parser.add_argument("--skip-xlsx", action="store_true", help="Não gera arquivo .xlsx.")
    return parser.parse_args()


def write_csv(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fp:
        writer = csv.DictWriter(fp, fieldnames=["Canal", "Investimento"])
        writer.writeheader()
        writer.writerows(DEFAULT_DATA)


def write_xlsx(path: Path, sheet_name: str) -> None:
    from openpyxl import Workbook

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name

    sheet.append(["Canal", "Investimento"])
    for row in DEFAULT_DATA:
        sheet.append([row["Canal"], row["Investimento"]])

    workbook.save(path)


def main() -> None:
    args = parse_args()
    output_csv = Path(args.output_csv)
    output_xlsx = Path(args.output_xlsx)

    write_csv(output_csv)
    print(f"CSV fictício gerado em: {output_csv}")

    if args.skip_xlsx:
        print("Geração de XLSX ignorada por parâmetro --skip-xlsx.")
    else:
        try:
            write_xlsx(output_xlsx, args.sheet)
            print(f"XLSX fictício gerado em: {output_xlsx} (aba: {args.sheet})")
        except ModuleNotFoundError:
            print("openpyxl não disponível: XLSX não foi gerado. O CSV está pronto para uso.")

    print("Dados de campanha fictícia:")
    for row in DEFAULT_DATA:
        print(f"- {row['Canal']}: R$ {row['Investimento']:,.2f}")


if __name__ == "__main__":
    main()
