#!/usr/bin/env python3
"""Create a REDCap index dictionary workbook from workflow intermediates."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[4]))

from scripts.discover_events.event_discovery import parse_event_value, resolve_event_column_index
from scripts.standardize_record_id.presidio.redcap_subid_presidio import find_redcap_subids


STANDARDIZED_RE = re.compile(r"^(?P<irb>\d+)_+(?P<subid>s\d+)$", re.IGNORECASE)


@dataclass(frozen=True)
class DictionaryResult:
    output_path: Path
    subject_count: int
    event_count: int
    instrument_count: int
    verification_failures: list[dict[str, str]]


def read_csv_dicts(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as file:
        return list(csv.DictReader(file))


def read_csv_table(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.reader(file)
        headers = next(reader)
        rows = list(reader)
    return headers, rows


def _nonblank(value: object) -> bool:
    return value is not None and str(value).strip() != ""


def _sort_standardized(value: str) -> tuple[int, int, str]:
    match = STANDARDIZED_RE.match(value)
    if not match:
        return (10**12, 10**12, value)
    return (int(match.group("irb")), int(match.group("subid")[1:]), value)


def build_subject_rows(
    input_path: Path,
    standardized_output_path: Path,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    original_headers, original_rows = read_csv_table(input_path)
    standardized_rows = read_csv_dicts(standardized_output_path)

    raw_entries_by_standardized: dict[str, set[str]] = {}
    failures: list[dict[str, str]] = []

    for row_number, (original_row, standardized_row) in enumerate(
        zip(original_rows, standardized_rows),
        start=1,
    ):
        if standardized_row.get("redcap_subid_detected") != "True":
            continue

        primary_standardized = standardized_row.get("redcap_subid_canonicalized", "")
        matches = []
        for cell in original_row:
            matches.extend(find_redcap_subids(cell))

        canonical_values = sorted({match.canonicalized for match in matches})
        raw_values = sorted({match.text for match in matches if match.canonicalized == primary_standardized})
        if primary_standardized:
            raw_entries_by_standardized.setdefault(primary_standardized, set()).update(raw_values)

        if len(canonical_values) != 1 or primary_standardized not in canonical_values:
            failures.append(
                {
                    "row_number": str(row_number),
                    "primary_standardized": primary_standardized,
                    "canonical_values_found": ", ".join(canonical_values),
                    "raw_record_id": standardized_row.get(original_headers[0], ""),
                }
            )

    subject_rows = []
    for standardized in sorted(raw_entries_by_standardized, key=_sort_standardized):
        match = STANDARDIZED_RE.match(standardized)
        irb = match.group("irb") if match else ""
        subid = match.group("subid") if match else ""
        subject_rows.append(
            {
                "IRB": irb,
                "subid": subid,
                "standardized": standardized,
                "raw_entries": ", ".join(sorted(raw_entries_by_standardized[standardized])),
            }
        )

    return subject_rows, failures


def build_event_rows(events_by_arm_path: Path) -> list[dict[str, str]]:
    rows = []
    for row in read_csv_dicts(events_by_arm_path):
        rows.append(
            {
                "arm": row.get("arm", ""),
                "order": row.get("event_order", ""),
                "event_name": row.get("event_name", ""),
                "event_label": "",
                "abbreviation": "",
            }
        )
    return rows


def _event_sort_key(event_row: dict[str, str]) -> tuple[int, int, str]:
    arm = event_row.get("arm", "")
    order = event_row.get("order", "")
    return (
        int(arm) if arm.isdigit() else 10**9,
        int(order) if order.isdigit() else 10**9,
        event_row.get("event_name", ""),
    )


def build_instrument_rows(
    input_path: Path,
    instrument_summary_path: Path,
    events_by_arm_path: Path,
    event_column_name: str | None = None,
    event_column_index: int = 1,
) -> list[dict[str, str]]:
    headers, rows = read_csv_table(input_path)
    instrument_summary = read_csv_dicts(instrument_summary_path)
    event_rows = build_event_rows(events_by_arm_path)
    event_order = {
        (row["arm"], row["event_name"]): index
        for index, row in enumerate(sorted(event_rows, key=_event_sort_key))
    }

    event_index = resolve_event_column_index(headers, column_name=event_column_name, column_index=event_column_index)
    parsed_events = []
    for row in rows:
        raw_event = row[event_index] if event_index < len(row) else ""
        parsed_events.append(parse_event_value(raw_event))

    output_rows = []
    for instrument in instrument_summary:
        start = int(instrument["start_index"])
        end = int(instrument["end_index"])
        events_present: set[tuple[str, str]] = set()

        for row, parsed_event in zip(rows, parsed_events):
            values = row[start : end + 1]
            if any(_nonblank(value) for value in values):
                events_present.add((parsed_event.arm, parsed_event.event_name))

        ordered_events = sorted(
            events_present,
            key=lambda key: event_order.get(key, 10**9),
        )
        event_text = "; ".join(f"Arm {arm}: {event_name}" for arm, event_name in ordered_events)
        output_rows.append(
            {
                "instrument": instrument["instrument_key"],
                "instrument_label": "",
                "abbreviation": "",
                "number_of_events": str(len(ordered_events)),
                "events": event_text,
                "number_of_columns": instrument["column_count"],
            }
        )

    return output_rows


def write_sheet(workbook: Workbook, title: str, rows: list[dict[str, str]], headers: list[str]) -> None:
    worksheet = workbook.active if workbook.active.title == "Sheet" else workbook.create_sheet()
    worksheet.title = title
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row.get(header, "") for header in headers])
    for column_cells in worksheet.columns:
        width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 80)
        worksheet.column_dimensions[column_cells[0].column_letter].width = width


def create_dictionary(
    input_path: str | Path,
    standardized_output_path: str | Path,
    events_by_arm_path: str | Path,
    instrument_summary_path: str | Path,
    output_path: str | Path,
    event_column_name: str | None = None,
    event_column_index: int = 1,
) -> DictionaryResult:
    input_path = Path(input_path)
    standardized_output_path = Path(standardized_output_path)
    events_by_arm_path = Path(events_by_arm_path)
    instrument_summary_path = Path(instrument_summary_path)
    output_path = Path(output_path)

    subject_rows, verification_failures = build_subject_rows(input_path, standardized_output_path)
    event_rows = build_event_rows(events_by_arm_path)
    instrument_rows = build_instrument_rows(
        input_path,
        instrument_summary_path,
        events_by_arm_path,
        event_column_name=event_column_name,
        event_column_index=event_column_index,
    )

    workbook = Workbook()
    write_sheet(workbook, "subject_id", subject_rows, ["IRB", "subid", "standardized", "raw_entries"])
    write_sheet(workbook, "event", event_rows, ["arm", "order", "event_name", "event_label", "abbreviation"])
    write_sheet(
        workbook,
        "instrument",
        instrument_rows,
        ["instrument", "instrument_label", "abbreviation", "number_of_events", "events", "number_of_columns"],
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)

    return DictionaryResult(
        output_path=output_path,
        subject_count=len(subject_rows),
        event_count=len(event_rows),
        instrument_count=len(instrument_rows),
        verification_failures=verification_failures,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create a REDCap index dictionary workbook.")
    parser.add_argument("input", type=Path, help="Original REDCap CSV")
    parser.add_argument("--standardized", type=Path, required=True, help="Standardized record ID CSV")
    parser.add_argument("--events-by-arm", type=Path, required=True, help="Events-by-arm CSV")
    parser.add_argument("--instrument-summary", type=Path, required=True, help="Instrument summary CSV")
    parser.add_argument("--out", type=Path, required=True, help="Output XLSX path")
    parser.add_argument("--event-column-name", help="Event column name override")
    parser.add_argument("--event-column-index", type=int, default=1, help="Event zero-based column index")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = create_dictionary(
        input_path=args.input,
        standardized_output_path=args.standardized,
        events_by_arm_path=args.events_by_arm,
        instrument_summary_path=args.instrument_summary,
        output_path=args.out,
        event_column_name=args.event_column_name,
        event_column_index=args.event_column_index,
    )
    print(result.output_path)
    print(f"subjects={result.subject_count}")
    print(f"events={result.event_count}")
    print(f"instruments={result.instrument_count}")
    print(f"verification_failures={len(result.verification_failures)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
