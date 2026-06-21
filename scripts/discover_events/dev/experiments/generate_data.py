#!/usr/bin/env python3
"""Generate event-discovery experiment data from real and synthetic events."""

from __future__ import annotations

import argparse
import csv
import random
import sys
from pathlib import Path


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[4]))

from scripts.discover_events.event_discovery import parse_event_value, resolve_event_column_index
from scripts.discover_events.run import read_table


DEFAULT_SOURCE_PATHS = [
    Path("58807-54909-BRAINS-wiki/data/FREEZE/raw_exports/redcap/all/58807SchatzbergRapid_DATA_2026-06-15_0333.csv"),
    Path("58807-54909-BRAINS-wiki/data/FREEZE/raw_exports/redcap/all/54909OutpatientBrain_DATA_2026-06-15_0334.csv"),
    Path("63771-LEAP-wiki/63771-REDCap-codebook_events.csv"),
]

SYNTHETIC_BASE_EVENTS = [
    "synthetic_event",
    "screening_visit_1",
    "baseline_visit_2",
    "study_day_1",
    "follow_up",
    "patient_contact_log",
    "regulatory",
    "exit_visit",
]

ARM_VARIANTS = [
    "{event}_arm{arm}",
    "{event} arm {arm}",
    "{event} arm {arm} a",
    "{event} arm {arm}a",
    "{event}_arm_{arm}",
    "{event}_arm_{arm}fxyz",
    "{event} (Arm {arm}: Synthetic Arm)",
    "arm{arm} {event}",
    "arm second {event}",
]


def _real_rows_from_source(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    headers, rows = read_table(path)
    try:
        event_index = resolve_event_column_index(headers)
    except ValueError:
        return []

    seen: set[str] = set()
    results: list[dict[str, str]] = []
    for row in rows:
        raw = row[event_index] if event_index < len(row) else ""
        if not raw or raw in seen:
            continue
        seen.add(raw)
        parsed = parse_event_value(raw)
        results.append(
            {
                "raw": parsed.raw_event,
                "arm": parsed.arm,
                "event_name": parsed.event_name,
                "source": str(path),
            }
        )
    return results


def _synthetic_rows() -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for event in SYNTHETIC_BASE_EVENTS:
        for arm in [1, 2, 3, 4]:
            for template in ARM_VARIANTS:
                if "second" in template and arm != 2:
                    continue
                raw = template.format(event=event, arm=arm)
                parsed = parse_event_value(raw)
                rows.append(
                    {
                        "raw": parsed.raw_event,
                        "arm": parsed.arm,
                        "event_name": parsed.event_name,
                        "source": "synthetic",
                    }
                )
    return rows


def generate_rows(
    total_rows: int = 5000,
    random_seed: int = 42,
    source_paths: list[Path] | None = None,
) -> list[dict[str, str]]:
    rng = random.Random(random_seed)
    source_paths = DEFAULT_SOURCE_PATHS if source_paths is None else source_paths

    seed_rows: list[dict[str, str]] = []
    for source_path in source_paths:
        seed_rows.extend(_real_rows_from_source(source_path))
    seed_rows.extend(_synthetic_rows())

    if not seed_rows:
        seed_rows = _synthetic_rows()

    rows = list(seed_rows)
    while len(rows) < total_rows:
        rows.append(dict(rng.choice(seed_rows)))

    rng.shuffle(rows)
    return rows[:total_rows]


def write_rows(rows: list[dict[str, str]], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["raw", "arm", "event_name", "source"])
        writer.writeheader()
        writer.writerows(rows)
    return output_path


def default_output_path() -> Path:
    return Path(__file__).with_name("data.csv")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate event-discovery experiment data.")
    parser.add_argument("--rows", type=int, default=5000, help="Number of rows to generate")
    parser.add_argument("--random-seed", type=int, default=42, help="Random seed")
    parser.add_argument("--source", type=Path, action="append", help="CSV/Excel source file; repeatable")
    parser.add_argument("--out", type=Path, default=default_output_path(), help="Output CSV path")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = generate_rows(total_rows=args.rows, random_seed=args.random_seed, source_paths=args.source)
    output_path = write_rows(rows, args.out)
    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
