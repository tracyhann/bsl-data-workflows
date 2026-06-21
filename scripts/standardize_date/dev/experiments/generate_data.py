#!/usr/bin/env python3
"""Generate date-standardization experiment data."""

from __future__ import annotations

import argparse
import csv
import random
import sys
from datetime import date, timedelta
from pathlib import Path


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[4]))


MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_FULL = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

DISTRACTORS = [
    "",
    "NA",
    "N/A",
    ".",
    "MDD_10598",
    "58807_s025",
    "General Note to File 1",
    "March 2026",
    "6/7",
    "2026",
    "02/30/2026",
    "13/40/2026",
    "555-123-4567",
    "visit 1",
    "baseline",
    "screening_arm_1",
    "2026-99-99",
    "QIDS total 18",
]


def _ordinal(day: int) -> str:
    if 10 <= day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def _date_variants(value: date) -> list[str]:
    month_abbr = MONTH_ABBR[value.month - 1]
    month_full = MONTH_FULL[value.month - 1]
    two_digit_year = f"{value.year % 100:02d}"
    return [
        value.isoformat(),
        value.strftime("%Y/%m/%d"),
        value.strftime("%Y.%m.%d"),
        value.strftime("%Y%m%d"),
        value.strftime("%m/%d/%Y"),
        value.strftime("%-m/%-d/%Y"),
        value.strftime("%m/%d/") + two_digit_year,
        value.strftime("%-m/%-d/") + two_digit_year,
        f"{month_abbr} {value.day}, {value.year}",
        f"{month_full} {value.day} {value.year}",
        f"{month_full} {_ordinal(value.day)}, {value.year}",
        f"{value.day}-{month_abbr}-{value.year}",
        f"{value.day} {month_full} {value.year}",
        value.strftime("%Y-%m-%d 14:30"),
        f"Assessment date: {value.strftime('%m/%d/%Y')}",
        f"completed on {month_abbr} {value.day}, {value.year}",
    ]


def _random_date(rng: random.Random) -> date:
    start = date(1990, 1, 1)
    end = date(2035, 12, 31)
    offset = rng.randint(0, (end - start).days)
    return start + timedelta(days=offset)


def generate_rows(total_rows: int = 5000, random_seed: int = 42) -> list[dict[str, str]]:
    rng = random.Random(random_seed)
    rows: list[dict[str, str]] = []

    fixed_date = date(2026, 6, 15)
    for raw in _date_variants(fixed_date):
        rows.append({"raw": raw, "canonicalized": fixed_date.isoformat()})
    for raw in DISTRACTORS:
        rows.append({"raw": raw, "canonicalized": "NA"})

    while len(rows) < total_rows:
        if rng.random() < 0.7:
            value = _random_date(rng)
            raw = rng.choice(_date_variants(value))
            if rng.random() < 0.15:
                raw = f"  {raw}  "
            rows.append({"raw": raw, "canonicalized": value.isoformat()})
        else:
            rows.append({"raw": rng.choice(DISTRACTORS), "canonicalized": "NA"})

    rng.shuffle(rows)
    return rows[:total_rows]


def write_rows(rows: list[dict[str, str]], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["raw", "canonicalized"])
        writer.writeheader()
        writer.writerows(rows)
    return output_path


def default_output_path() -> Path:
    return Path(__file__).with_name("data.csv")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate synthetic date standardization experiment data.")
    parser.add_argument("--rows", type=int, default=5000, help="Number of rows to generate")
    parser.add_argument("--random-seed", type=int, default=42, help="Random seed")
    parser.add_argument("--out", type=Path, default=default_output_path(), help="Output CSV path")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = generate_rows(total_rows=args.rows, random_seed=args.random_seed)
    output_path = write_rows(rows, args.out)
    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
