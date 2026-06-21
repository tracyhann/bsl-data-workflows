#!/usr/bin/env python3
"""Apply date recognition/standardization to one table column."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Iterable


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[2]))

from scripts.standardize_date.date_standardizer import find_dates


OUTPUT_COLUMNS = [
    "date_detected",
    "date_standardized",
    "date_match",
    "date_parse_status",
    "date_score",
    "date_start",
    "date_end",
]


def _dedupe_headers(headers: Iterable[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for header in headers:
        name = str(header)
        if name not in seen:
            seen[name] = 0
            result.append(name)
            continue
        seen[name] += 1
        result.append(f"{name}.{seen[name]}")
    return result


def read_delimited_table(path: Path, delimiter: str) -> tuple[list[str], list[list[str]]]:
    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.reader(file, delimiter=delimiter)
        headers = next(reader)
        rows = list(reader)
    return _dedupe_headers(headers), rows


def read_excel_table(path: Path, sheet_name: str | int | None = None) -> tuple[list[str], list[list[str]]]:
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("Reading Excel files requires pandas/openpyxl in this environment.") from exc

    selected_sheet = 0 if sheet_name is None else sheet_name
    dataframe = pd.read_excel(path, sheet_name=selected_sheet, dtype=str, keep_default_na=False)
    headers = _dedupe_headers([str(column) for column in dataframe.columns])
    rows = dataframe.astype(str).values.tolist()
    return headers, rows


def read_table(path: Path, sheet_name: str | int | None = None) -> tuple[list[str], list[list[str]]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return read_delimited_table(path, ",")
    if suffix == ".tsv":
        return read_delimited_table(path, "\t")
    if suffix in {".xlsx", ".xls"}:
        return read_excel_table(path, sheet_name)
    raise ValueError(f"Unsupported input extension {path.suffix!r}; expected CSV, TSV, XLSX, or XLS.")


def resolve_column_index(headers: list[str], column_name: str | None = None, column_index: int = 0) -> int:
    if column_name:
        if column_name in headers:
            return headers.index(column_name)
        lowered = column_name.lower()
        for index, header in enumerate(headers):
            if header.lower() == lowered:
                return index
        raise ValueError(f"Column name {column_name!r} was not found.")

    if column_index < 0 or column_index >= len(headers):
        raise ValueError(f"Column index {column_index} is out of range for {len(headers)} columns.")
    return column_index


def annotate_row(row: list[str], column_index: int) -> list[str]:
    value = row[column_index] if column_index < len(row) else ""
    matches = find_dates(value)
    if not matches:
        return ["False", "NA", "", "", "", "", ""]

    match = matches[0]
    return [
        "True",
        match.canonicalized,
        match.text,
        match.parse_status,
        f"{match.score:.3f}",
        str(match.start),
        str(match.end),
    ]


def process_file(
    input_path: str | Path,
    output_path: str | Path,
    column_name: str | None = None,
    column_index: int = 0,
    sheet_name: str | int | None = None,
) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path)
    headers, rows = read_table(input_path, sheet_name)
    target_index = resolve_column_index(headers, column_name, column_index)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(headers + OUTPUT_COLUMNS)
        for row in rows:
            padded = row + [""] * (len(headers) - len(row))
            writer.writerow(padded[: len(headers)] + annotate_row(padded, target_index))

    return output_path


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_date_standardized.csv")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Identify and standardize dates in a selected CSV/Excel column.")
    parser.add_argument("input", type=Path, help="Input CSV, TSV, XLSX, or XLS file")
    parser.add_argument("--out", type=Path, help="Output CSV path")
    parser.add_argument("--column-name", help="Column name to process, e.g. visit_date")
    parser.add_argument("--column-index", type=int, default=0, help="Zero-based column index to process")
    parser.add_argument("--sheet", help="Excel sheet name; defaults to the first sheet")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_path = args.out or default_output_path(args.input)
    result = process_file(
        args.input,
        output_path,
        column_name=args.column_name,
        column_index=args.column_index,
        sheet_name=args.sheet,
    )
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
