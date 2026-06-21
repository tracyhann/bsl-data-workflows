#!/usr/bin/env python3
"""Apply REDCap subject ID recognition/standardization to one table column."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Iterable


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[2]))

from scripts.discover_record_id import (
    CandidateColumn,
    find_event_column,
    format_summary_rows,
    profile_candidate_column,
    write_csv,
)
from scripts.standardize_record_id.presidio.redcap_subid_presidio import find_redcap_subids


OUTPUT_COLUMNS = [
    "redcap_subid_detected",
    "redcap_subid_canonicalized",
    "redcap_subid_match",
    "redcap_subid_entity_type",
    "redcap_subid_score",
    "redcap_subid_start",
    "redcap_subid_end",
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


def read_csv_table(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open(newline="", encoding="utf-8-sig") as file:
        sample = file.read(8192)
        file.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel
        reader = csv.reader(file, dialect)
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
    if suffix in {".csv", ".tsv"}:
        return read_csv_table(path)
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
    matches = find_redcap_subids(value)
    if not matches:
        return ["False", "NA", "", "", "", "", ""]

    match = matches[0]
    return [
        "True",
        match.canonicalized,
        match.text,
        match.entity_type,
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
    summary_output_path: str | Path | None = None,
    summary_examples: int = 5,
    random_seed: int = 42,
) -> dict[str, Path]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    summary_path = Path(summary_output_path) if summary_output_path else default_summary_output_path(output_path)
    headers, rows = read_table(input_path, sheet_name)
    target_index = resolve_column_index(headers, column_name, column_index)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(headers + OUTPUT_COLUMNS)
        for row in rows:
            padded = row + [""] * (len(headers) - len(row))
            writer.writerow(padded[: len(headers)] + annotate_row(padded, target_index))

    candidate = CandidateColumn(
        index=target_index,
        header=headers[target_index],
        score=100,
        reason="selected standardization column",
    )
    _, value_rows = profile_candidate_column(
        input_path,
        candidate,
        rows,
        find_event_column(headers),
        max_examples=summary_examples,
    )
    summary_rows = format_summary_rows(
        input_path,
        value_rows,
        scope=f"selected_column_index_{target_index}",
        random_seed=random_seed,
        examples_per_format=summary_examples,
    )
    write_csv(summary_path, summary_rows)

    return {"standardized": output_path, "format_summary": summary_path}


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_redcap_subid_standardized.csv")


def default_summary_output_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_record_id_format_summary.csv")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Identify and standardize REDCap subject IDs in a selected CSV/Excel column."
    )
    parser.add_argument("input", type=Path, help="Input CSV, TSV, XLSX, or XLS file")
    parser.add_argument("--out", type=Path, help="Output CSV path")
    parser.add_argument("--summary-out", type=Path, help="Output CSV path for coded-format summary")
    parser.add_argument("--column-name", help="Column name to process, e.g. record_id or Record ID")
    parser.add_argument("--column-index", type=int, default=0, help="Zero-based column index to process")
    parser.add_argument("--sheet", help="Excel sheet name; defaults to the first sheet")
    parser.add_argument(
        "--summary-examples",
        type=int,
        default=5,
        help="Number of random examples to include for each coded-format summary row",
    )
    parser.add_argument("--random-seed", type=int, default=42, help="Seed for deterministic summary examples")
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
        summary_output_path=args.summary_out,
        summary_examples=args.summary_examples,
        random_seed=args.random_seed,
    )
    for path in result.values():
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
