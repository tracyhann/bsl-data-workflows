#!/usr/bin/env python3
"""Generate REDCap subject ID standardization experiment data."""

from __future__ import annotations

import argparse
import csv
import random
import sys
from pathlib import Path


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[4]))

from scripts.standardize_record_id.presidio.redcap_subid_presidio import standardize_redcap_subid


DEFAULT_SOURCE_FILES = [
    Path(
        "58807-54909-BRAINS-wiki/data/FREEZE/raw_exports/redcap/all/"
        "58807SchatzbergRapid_DATA_2026-06-15_0333.csv"
    ),
    Path(
        "58807-54909-BRAINS-wiki/data/FREEZE/raw_exports/redcap/all/"
        "54909OutpatientBrain_DATA_2026-06-15_0334.csv"
    ),
]

SEED_ROWS = [
    {"raw": "58807s25", "canonicalized": "58807_s025"},
    {"raw": "58807_s025", "canonicalized": "58807_s025"},
    {"raw": "58807-s025", "canonicalized": "58807_s025"},
    {"raw": "IRB 58807 subject 25", "canonicalized": "58807_s025"},
    {"raw": "54909_sub002a", "canonicalized": "54909_s002"},
    {"raw": "General Note to File 1", "canonicalized": "NA"},
    {"raw": "MDD_10598", "canonicalized": "NA"},
    {"raw": "24", "canonicalized": "NA"},
]

STATIC_DISTRACTORS = [
    "General Note to File 1",
    "General Note to File 10",
    "MDD_10598",
    "OCD_2273",
    "screening_visit_1_arm_1",
    "baseline_visit_2_arm_2",
    "2024-06-15",
    "03/15/2024",
    "TEST",
    "Complete",
    "Unchecked",
    "Patient Contact Log",
    "",
]


def read_record_ids(path: Path) -> list[str]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.reader(file)
        headers = next(reader)
        try:
            record_index = [header.lower() for header in headers].index("record_id")
        except ValueError:
            record_index = 0
        values = []
        for row in reader:
            if record_index < len(row):
                value = row[record_index].strip()
                if value:
                    values.append(value)
    return sorted(set(values))


def split_canonical(canonical: str) -> tuple[str, int]:
    irb, subject = canonical.split("_s", 1)
    return irb, int(subject)


def synthetic_positive_variant(canonical: str, rng: random.Random) -> str:
    irb, subject_number = split_canonical(canonical)
    subject_3 = f"{subject_number:03d}"
    subject_loose = str(subject_number)
    templates = [
        f"{irb}_s{subject_3}",
        f"{irb}s{subject_3}",
        f"{irb}s{subject_loose}",
        f"{irb}-s{subject_3}",
        f"{irb}.s{subject_3}",
        f"{irb} s{subject_3}",
        f"{irb} S {subject_loose}",
        f"IRB{irb}_s{subject_3}",
        f"IRB {irb} subject {subject_loose}",
        f"{irb}_sub{subject_3}",
        f"{irb}_subject{subject_loose}",
        f"{irb}_S{subject_3}",
        f"{irb}_s{subject_3}a",
        f"{irb}s{subject_3}XOVER",
        f"{irb}_s{subject_3}_v2",
    ]
    return rng.choice(templates)


def row_for_raw(raw: str) -> dict[str, str]:
    canonicalized = standardize_redcap_subid(raw)
    return {"raw": raw, "canonicalized": canonicalized or "NA"}


def generate_rows(total_rows: int = 5000, random_seed: int = 42) -> list[dict[str, str]]:
    rng = random.Random(random_seed)
    real_values: list[str] = []
    for source_file in DEFAULT_SOURCE_FILES:
        real_values.extend(read_record_ids(source_file))
    real_values = sorted(set(real_values))

    real_positive = [value for value in real_values if standardize_redcap_subid(value)]
    real_negative = [value for value in real_values if not standardize_redcap_subid(value)]
    canonical_pool = sorted({standardize_redcap_subid(value) for value in real_positive if standardize_redcap_subid(value)})

    rows = list(SEED_ROWS[: min(total_rows, len(SEED_ROWS))])
    while len(rows) < total_rows:
        roll = rng.random()
        if roll < 0.40 and real_values:
            rows.append(row_for_raw(rng.choice(real_values)))
        elif roll < 0.86 and canonical_pool:
            canonical = rng.choice(canonical_pool)
            rows.append({"raw": synthetic_positive_variant(canonical, rng), "canonicalized": canonical})
        else:
            distractor_pool = real_negative + STATIC_DISTRACTORS
            rows.append({"raw": rng.choice(distractor_pool), "canonicalized": "NA"})

    rng.shuffle(rows)
    return rows[:total_rows]


def write_rows(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["raw", "canonicalized"])
        writer.writeheader()
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate REDCap subid standardization experiment data.")
    parser.add_argument("--out", type=Path, default=Path(__file__).with_name("data.csv"))
    parser.add_argument("--rows", type=int, default=5000)
    parser.add_argument("--random-seed", type=int, default=42)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = generate_rows(total_rows=args.rows, random_seed=args.random_seed)
    write_rows(args.out, rows)
    print(args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
