#!/usr/bin/env python3
"""Run REDCap ID standardization, event/instrument discovery, and dictionary creation."""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[4]))

from scripts.discover_events.run import process_file as discover_events
from scripts.discover_instruments.run import process_file as discover_instruments
from scripts.standardize_record_id.run import process_file as standardize_record_id
from scripts.workflows.clean_redcap_instruments.steps.create_redcap_index_dictionary import DictionaryResult, create_dictionary


DEFAULT_INPUT = Path("examples/redcap/example_DATA.csv")
DEFAULT_OUT_DIR = Path("intermediates")
DEFAULT_LOG = Path("log.md")


def timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d-%H%M%S")


def read_csv_dicts(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as file:
        return list(csv.DictReader(file))


def markdown_table(headers: list[str], rows: list[list[object]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(str(cell) for cell in row) + " |")
    return "\n".join(lines)


def markdown_arg_value(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", " ").replace("|", "\\|")


def workflow_arguments_section(input_path: Path, workflow_args: dict[str, object] | None = None) -> str:
    args = workflow_args or {"input": input_path}
    rows = [[key, markdown_arg_value(value)] for key, value in args.items()]
    return "## Workflow Arguments\n\n" + markdown_table(["arg", "value"], rows)


def grid_table(items: list[str], columns: int = 10) -> str:
    headers = [str(index) for index in range(1, columns + 1)]
    rows = []
    for start in range(0, len(items), columns):
        row = items[start : start + columns]
        rows.append(row + [""] * (columns - len(row)))
    return markdown_table(headers, rows)


def path_list(paths: list[Path]) -> str:
    return "\n".join(f"- `{path}`" for path in paths)


def summarize_record_id(standardized_path: Path, format_summary_path: Path) -> tuple[str, list[dict[str, str]]]:
    standardized_rows = read_csv_dicts(standardized_path)
    format_rows = read_csv_dicts(format_summary_path)
    detected = [row for row in standardized_rows if row.get("redcap_subid_detected") == "True"]
    unique_raw = sorted({row.get("redcap_subid_match", "") for row in detected if row.get("redcap_subid_match", "")})
    unique_canonical = sorted(
        {
            row.get("redcap_subid_canonicalized", "")
            for row in detected
            if row.get("redcap_subid_canonicalized", "") != "NA"
        }
    )

    table = markdown_table(
        ["coded_format", "unique_id_values", "unique_canonical_ids", "row_occurrences", "examples"],
        [
            [
                row.get("coded_format", ""),
                row.get("unique_id_values", ""),
                row.get("unique_canonical_ids", ""),
                row.get("row_occurrences", ""),
                row.get("examples", ""),
            ]
            for row in format_rows
        ],
    )
    summary = (
        f"Rows identified and standardized: **{len(detected)}** / {len(standardized_rows)}\n\n"
        f"Unique raw detected entries: **{len(unique_raw)}**\n\n"
        f"Unique canonical standardized IDs: **{len(unique_canonical)}**\n\n"
        f"{table}"
    )
    return summary, format_rows


def summarize_instruments(instrument_summary_path: Path) -> tuple[str, list[dict[str, str]]]:
    rows = read_csv_dicts(instrument_summary_path)
    names = [row.get("instrument_key", "") for row in rows]
    summary = f"Instruments discovered: **{len(rows)}**\n\n{grid_table(names, columns=5)}"
    return summary, rows


def summarize_events(events_by_arm_path: Path) -> tuple[str, list[dict[str, str]]]:
    rows = read_csv_dicts(events_by_arm_path)
    table_rows = [
        [
            row.get("arm", ""),
            row.get("event_order", ""),
            row.get("event_name", ""),
            row.get("first_raw_event", ""),
            row.get("count", ""),
        ]
        for row in rows
    ]
    arms = list(dict.fromkeys(row.get("arm", "") for row in rows))
    summary = (
        f"Arms discovered: **{', '.join(arms)}**\n\n"
        f"Arm-event rows: **{len(rows)}**\n\n"
        f"{markdown_table(['arm', 'order', 'event_name', 'first_raw_event', 'count'], table_rows)}"
    )
    return summary, rows


def summarize_codebook_verification(path: Path | None, table_headers: list[str]) -> str:
    if path is None or not path.exists():
        return ""

    rows = read_csv_dicts(path)
    advisory = "Raw discovery remains the source of truth; codebook verification is advisory."
    if not rows:
        return f"Codebook verification mismatches: **0**\n\n{advisory}"

    table_rows = [[row.get(header, "") for header in table_headers] for row in rows]
    return (
        f"Codebook verification mismatches: **{len(rows)}**\n\n"
        f"{advisory}\n\n"
        f"{markdown_table(table_headers, table_rows)}"
    )


def dictionary_log(dictionary_result: DictionaryResult, dictionary_time: str, script_path: Path) -> str:
    verification_rule = (
        "Verification rule: for every row whose primary record ID standardizes to a REDCap subject ID, "
        "the workflow scans all cells in that row for REDCap-like IDs and expects exactly one canonical "
        "`IRB_s0*` value."
    )
    if dictionary_result.verification_failures:
        failure_table = markdown_table(
            ["row_number", "primary_standardized", "canonical_values_found", "raw_record_id"],
            [
                [
                    row.get("row_number", ""),
                    row.get("primary_standardized", ""),
                    row.get("canonical_values_found", ""),
                    row.get("raw_record_id", ""),
                ]
                for row in dictionary_result.verification_failures
            ],
        )
        verification = (
            f"{verification_rule}\n\n"
            f"Verification failures: **{len(dictionary_result.verification_failures)}**\n\n"
            "A failure means the row-level scan found zero or multiple REDCap-like canonical IDs.\n\n"
            f"{failure_table}"
        )
    else:
        verification = (
            f"{verification_rule}\n\n"
            "Verification failures: **0**. Every row with an identified REDCap subid standardized to "
            "exactly one canonical `IRB_s0*` value."
        )

    return (
        "## Subject ID\n\n"
        f"script: `{script_path}`\n\n"
        f"time: {dictionary_time}\n\n"
        "outputs:\n"
        f"- `{dictionary_result.output_path}`\n\n"
        f"Subject IDs indexed: **{dictionary_result.subject_count}**\n\n"
        f"{verification}\n\n"
        "## Event\n\n"
        f"script: `{script_path}`\n\n"
        f"time: {dictionary_time}\n\n"
        "outputs:\n"
        f"- `{dictionary_result.output_path}` sheet `event`\n\n"
        f"Events indexed: **{dictionary_result.event_count}**\n\n"
        "## Instrument\n\n"
        f"script: `{script_path}`\n\n"
        f"time: {dictionary_time}\n\n"
        "outputs:\n"
        f"- `{dictionary_result.output_path}` sheet `instrument`\n\n"
        f"Instruments indexed: **{dictionary_result.instrument_count}**"
    )


def build_log(
    input_path: Path,
    record_time: str,
    instrument_time: str,
    event_time: str,
    dictionary_time: str,
    record_paths: dict[str, Path],
    instrument_paths: dict[str, Path],
    event_paths: dict[str, Path],
    dictionary_result: DictionaryResult,
    workflow_args: dict[str, object] | None = None,
) -> str:
    record_summary, _ = summarize_record_id(record_paths["standardized"], record_paths["format_summary"])
    instrument_summary, _ = summarize_instruments(instrument_paths["instrument_summary"])
    event_summary, _ = summarize_events(event_paths["events_by_arm"])
    instrument_codebook_summary = summarize_codebook_verification(
        instrument_paths.get("instrument_codebook_verification"),
        ["status", "expected_form_name", "discovered_instrument_key", "message"],
    )
    event_codebook_summary = summarize_codebook_verification(
        event_paths.get("event_codebook_verification"),
        ["status", "expected_arm", "expected_event_name", "discovered_arm", "discovered_event_name", "message"],
    )

    standardize_script = Path("scripts/standardize_record_id/run.py")
    instrument_script = Path("scripts/discover_instruments/run.py")
    event_script = Path("scripts/discover_events/run.py")
    dictionary_script = Path("scripts/workflows/clean_redcap_instruments/steps/create_redcap_index_dictionary.py")

    return (
        "# REDCap Workflow Log\n\n"
        f"{workflow_arguments_section(input_path, workflow_args)}\n\n"
        "<details>\n"
        "<summary><h1>Discovery</h1></summary>\n\n"
        "## Record ID Discovery\n\n"
        f"script: `{standardize_script}`\n\n"
        f"time: {record_time}\n\n"
        "outputs:\n"
        f"{path_list(list(record_paths.values()))}\n\n"
        f"{record_summary}\n\n"
        "## Instrument Discovery\n\n"
        f"script: `{instrument_script}`\n\n"
        f"time: {instrument_time}\n\n"
        "outputs:\n"
        f"{path_list(list(instrument_paths.values()))}\n\n"
        f"{instrument_summary}\n\n"
        f"{instrument_codebook_summary}\n\n"
        "## Event Discovery\n\n"
        f"script: `{event_script}`\n\n"
        f"time: {event_time}\n\n"
        "outputs:\n"
        f"{path_list(list(event_paths.values()))}\n\n"
        f"{event_summary}\n\n"
        f"{event_codebook_summary}\n\n"
        "</details>\n\n"
        "<details>\n"
        "<summary><h1>Dictionary</h1></summary>\n\n"
        f"{dictionary_log(dictionary_result, dictionary_time, dictionary_script)}\n\n"
        "</details>\n"
    )


def run_workflow(
    input_path: str | Path = DEFAULT_INPUT,
    out_dir: str | Path = DEFAULT_OUT_DIR,
    log_path: str | Path = DEFAULT_LOG,
    codebook_path: str | Path | None = None,
    instrument_codebook_sheet: str | int | None = None,
    event_codebook_sheet: str | int | None = None,
) -> dict[str, object]:
    input_path = Path(input_path)
    out_dir = Path(out_dir)
    log_path = Path(log_path)
    codebook = Path(codebook_path) if codebook_path else None
    out_dir.mkdir(parents=True, exist_ok=True)

    stem = input_path.stem

    record_time = timestamp()
    record_paths = standardize_record_id(
        input_path=input_path,
        output_path=out_dir / f"{stem}_redcap_subid_standardized.csv",
        column_index=0,
        summary_output_path=out_dir / f"{stem}_record_id_format_summary.csv",
    )

    instrument_time = timestamp()
    instrument_paths = discover_instruments(
        input_path=input_path,
        out_dir=out_dir,
        codebook_path=codebook if instrument_codebook_sheet else None,
        codebook_sheet=instrument_codebook_sheet,
    )

    event_time = timestamp()
    event_paths = discover_events(
        input_path=input_path,
        out_dir=out_dir,
        codebook_path=codebook if event_codebook_sheet else None,
        codebook_sheet=event_codebook_sheet,
    )

    dictionary_time = timestamp()
    dictionary_result = create_dictionary(
        input_path=input_path,
        standardized_output_path=record_paths["standardized"],
        events_by_arm_path=event_paths["events_by_arm"],
        instrument_summary_path=instrument_paths["instrument_summary"],
        output_path=out_dir / "dictionary.xlsx",
    )

    log_text = build_log(
        input_path=input_path,
        record_time=record_time,
        instrument_time=instrument_time,
        event_time=event_time,
        dictionary_time=dictionary_time,
        record_paths=record_paths,
        instrument_paths=instrument_paths,
        event_paths=event_paths,
        dictionary_result=dictionary_result,
    )
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(log_text, encoding="utf-8")

    return {
        "record_id": record_paths,
        "instruments": instrument_paths,
        "events": event_paths,
        "dictionary": dictionary_result.output_path,
        "log": log_path,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the REDCap discovery and dictionary workflow.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Input REDCap CSV")
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR, help="Intermediate output directory")
    parser.add_argument("--log", type=Path, default=DEFAULT_LOG, help="Markdown log output path")
    parser.add_argument("--codebook", type=Path, help="Optional REDCap codebook workbook for verification")
    parser.add_argument("--instrument-codebook-sheet", help="Instrument sheet name in the codebook workbook")
    parser.add_argument("--event-codebook-sheet", help="Event sheet name in the codebook workbook")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = run_workflow(
        input_path=args.input,
        out_dir=args.out_dir,
        log_path=args.log,
        codebook_path=args.codebook,
        instrument_codebook_sheet=args.instrument_codebook_sheet,
        event_codebook_sheet=args.event_codebook_sheet,
    )
    print(result["log"])
    print(result["dictionary"])
    for group in ("record_id", "instruments", "events"):
        for path in result[group].values():
            print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
