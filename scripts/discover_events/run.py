#!/usr/bin/env python3
"""Discover REDCap event arms and ordered events from a CSV/Excel table."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Iterable


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[2]))

from scripts.discover_events.event_discovery import (
    discover_event_groups,
    parse_event_value,
    resolve_event_column_index,
)


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


def output_paths(input_path: Path, out_dir: Path) -> dict[str, Path]:
    stem = input_path.stem
    return {
        "event_values": out_dir / f"{stem}_event_values.csv",
        "events_by_arm": out_dir / f"{stem}_events_by_arm.csv",
    }


def _normalized_header(header: str) -> str:
    return str(header).strip().lower().replace("_", " ")


def _find_header(headers: list[str], names: set[str]) -> int:
    normalized_names = {_normalized_header(name) for name in names}
    for index, header in enumerate(headers):
        if _normalized_header(header) in normalized_names:
            return index
    raise ValueError(f"Could not find any of {sorted(names)} in headers {headers!r}.")


def _cell(row: list[str], index: int) -> str:
    return row[index].strip() if index < len(row) else ""


def expected_events_from_codebook(codebook_path: str | Path, codebook_sheet: str | int) -> list[tuple[str, str, str]]:
    """Return collapsed (arm, event_name, first_unique_event_name) entries from a REDCap event sheet."""
    headers, rows = read_table(Path(codebook_path), codebook_sheet)
    unique_event_index = _find_header(headers, {"Unique event name", "unique_event_name"})
    expected: list[tuple[str, str, str]] = []
    seen: set[tuple[str, str]] = set()

    for row in rows:
        unique_event_name = _cell(row, unique_event_index)
        if not unique_event_name:
            continue
        parsed = parse_event_value(unique_event_name)
        key = (parsed.arm, parsed.event_name)
        if key not in seen:
            expected.append((parsed.arm, parsed.event_name, unique_event_name))
            seen.add(key)

    return expected


def codebook_verification_rows(groups, expected_events: list[tuple[str, str, str]]) -> list[dict[str, str]]:
    discovered_keys = [(group.arm, group.event_name) for group in groups]
    discovered_order = {key: str(index + 1) for index, key in enumerate(discovered_keys)}
    expected_keys = [(arm, event_name) for arm, event_name, _ in expected_events]
    expected_order = {key: str(index + 1) for index, key in enumerate(expected_keys)}
    expected_raw = {(arm, event_name): raw for arm, event_name, raw in expected_events}
    first_raw = {(group.arm, group.event_name): group.first_raw_event for group in groups}
    rows: list[dict[str, str]] = []

    for key in expected_keys:
        if key not in discovered_order:
            rows.append(
                {
                    "status": "missing_from_discovery",
                    "expected_order": expected_order[key],
                    "discovered_order": "",
                    "expected_arm": key[0],
                    "expected_event_name": key[1],
                    "discovered_arm": "",
                    "discovered_event_name": "",
                    "expected_unique_event": expected_raw[key],
                    "first_raw_event": "",
                    "message": "Codebook unique event name was not found in discovered event groups.",
                }
            )

    for key in discovered_keys:
        if key not in expected_order:
            rows.append(
                {
                    "status": "extra_in_discovery",
                    "expected_order": "",
                    "discovered_order": discovered_order[key],
                    "expected_arm": "",
                    "expected_event_name": "",
                    "discovered_arm": key[0],
                    "discovered_event_name": key[1],
                    "expected_unique_event": "",
                    "first_raw_event": first_raw[key],
                    "message": "Discovered event group was not found in codebook unique event names.",
                }
            )

    return rows


def write_codebook_verification(path: Path, rows: list[dict[str, str]]) -> None:
    fieldnames = [
        "status",
        "expected_order",
        "discovered_order",
        "expected_arm",
        "expected_event_name",
        "discovered_arm",
        "discovered_event_name",
        "expected_unique_event",
        "first_raw_event",
        "message",
    ]
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def process_file(
    input_path: str | Path,
    out_dir: str | Path,
    column_name: str | None = None,
    column_index: int = 1,
    sheet_name: str | int | None = None,
    codebook_path: str | Path | None = None,
    codebook_sheet: str | int | None = None,
) -> dict[str, Path]:
    input_path = Path(input_path)
    out_dir = Path(out_dir)
    headers, rows = read_table(input_path, sheet_name)
    event_index = resolve_event_column_index(headers, column_name=column_name, column_index=column_index)

    parsed_rows = []
    raw_events = []
    for row_number, row in enumerate(rows, start=1):
        value = row[event_index] if event_index < len(row) else ""
        parsed = parse_event_value(value)
        raw_events.append(value)
        parsed_rows.append(
            {
                "row_number": row_number,
                "raw_event": parsed.raw_event,
                "arm": parsed.arm,
                "event_name": parsed.event_name,
            }
        )

    groups = discover_event_groups(raw_events)
    paths = output_paths(input_path, out_dir)
    if codebook_path:
        if codebook_sheet is None:
            raise ValueError("--codebook-sheet is required when --codebook is provided.")
        paths["event_codebook_verification"] = out_dir / f"{input_path.stem}_event_codebook_verification.csv"
    out_dir.mkdir(parents=True, exist_ok=True)

    with paths["event_values"].open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["row_number", "raw_event", "arm", "event_name"])
        writer.writeheader()
        writer.writerows(parsed_rows)

    with paths["events_by_arm"].open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=["arm", "event_order", "event_name", "first_raw_event", "first_row", "count"],
        )
        writer.writeheader()
        for group in groups:
            writer.writerow(
                {
                    "arm": group.arm,
                    "event_order": group.event_order,
                    "event_name": group.event_name,
                    "first_raw_event": group.first_raw_event,
                    "first_row": group.first_row,
                    "count": group.count,
                }
            )

    if codebook_path:
        expected_events = expected_events_from_codebook(codebook_path, codebook_sheet)
        verification_rows = codebook_verification_rows(groups, expected_events)
        write_codebook_verification(paths["event_codebook_verification"], verification_rows)

    return paths


def default_out_dir(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_event_discovery")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Discover REDCap arms and ordered per-arm events from an event-name column."
    )
    parser.add_argument("input", type=Path, help="Input CSV, TSV, XLSX, or XLS file")
    parser.add_argument("--out-dir", type=Path, help="Output directory")
    parser.add_argument("--column-name", help="Column name to process, e.g. redcap_event_name or Event Name")
    parser.add_argument("--column-index", type=int, default=1, help="Zero-based column index to process")
    parser.add_argument("--sheet", help="Excel sheet name; defaults to the first sheet")
    parser.add_argument("--codebook", type=Path, help="Optional REDCap codebook workbook for verification")
    parser.add_argument("--codebook-sheet", help="Codebook event sheet name for verification")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    out_dir = args.out_dir or default_out_dir(args.input)
    paths = process_file(
        args.input,
        out_dir,
        column_name=args.column_name,
        column_index=args.column_index,
        sheet_name=args.sheet,
        codebook_path=args.codebook,
        codebook_sheet=args.codebook_sheet,
    )
    for path in paths.values():
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
