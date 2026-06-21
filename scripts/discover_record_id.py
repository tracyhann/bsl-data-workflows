#!/usr/bin/env python3
"""Discover and profile record/subject ID naming formats in CSV exports.

This script is intentionally REDCap-friendly but not REDCap-specific. It reads
CSV files by physical column position so duplicate labels such as "Record ID"
do not collapse into one dictionary key.
"""

from __future__ import annotations

import argparse
import csv
import difflib
import json
import random
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


SEPARATOR_RE = re.compile(r"[\s._/\\-]+")
ALNUM_RE = re.compile(r"^[a-z0-9_]+$")


@dataclass(frozen=True)
class CandidateColumn:
    index: int
    header: str
    score: int
    reason: str


@dataclass(frozen=True)
class CanonicalRecordId:
    original: str
    cleaned: str
    canonical: str
    format_family: str
    coded_format: str
    shape_signature: str
    prefix_numeric: str
    prefix_alpha: str
    subject_token: str
    subject_token_prefix: str
    subject_number: str
    suffix_tokens: tuple[str, ...]
    needs_review: bool
    review_reason: str


@dataclass
class ColumnProfile:
    source_file: str
    column_index: int
    header: str
    score: int
    reason: str
    total_rows: int
    nonblank_rows: int
    unique_values: int
    duplicated_values: int
    rows_with_duplicated_values: int
    class_counts: Counter
    coded_format_counts: Counter
    shape_counts: Counter
    suffix_counts: Counter
    review_count: int
    examples: list[dict[str, str | int | bool]]


@dataclass
class CsvProfile:
    source_file: str
    total_rows: int
    total_columns: int
    event_column_index: int | None
    event_column_header: str
    columns: list[ColumnProfile]


def normalize_header(header: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(header).strip().lower())


def tokenize_header(header: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", str(header).strip().lower())


def candidate_score(header: str) -> tuple[int, str]:
    compact = normalize_header(header)
    tokens = tokenize_header(header)
    token_set = set(tokens)
    if not compact:
        return 0, ""

    # Survey identifiers are usually REDCap metadata, not subject identifiers.
    if "survey" in token_set and ({"identifier", "identifiers", "id", "ids"} & token_set):
        return 0, ""

    exact = {
        "recordid": 100,
        "recordsubjectid": 98,
        "recordno": 96,
        "recordnumber": 96,
        "subjectid": 95,
        "subjectno": 94,
        "subjectnumber": 94,
        "subectid": 94,
        "subjid": 94,
        "participantid": 95,
        "participantno": 94,
        "participantnumber": 94,
        "participantstudyid": 94,
        "patientid": 92,
        "patientno": 91,
        "patientnumber": 91,
        "studyid": 88,
        "studyno": 87,
        "studynumber": 87,
        "screeningid": 86,
        "screeningno": 85,
        "screeningnumber": 85,
        "screenid": 86,
        "screenno": 85,
        "screennumber": 85,
        "subid": 84,
        "subno": 83,
        "subnumber": 83,
    }
    if compact in exact:
        return exact[compact], "exact ID header"

    score = 0
    reasons: list[str] = []
    identifier_tokens = {"id", "ids", "identifier", "identifiers", "number"}
    has_identifier_token = bool(identifier_tokens & token_set)
    if has_identifier_token:
        score += 20
        reasons.append("contains ID/identifier token")

    weighted_terms = {
        "record": 45,
        "subject": 45,
        "subect": 44,
        "subj": 44,
        "participant": 42,
        "patient": 40,
        "study": 34,
        "screening": 34,
        "screen": 32,
        "sub": 20,
    }
    matched_stem = False
    for term, weight in weighted_terms.items():
        if term in token_set:
            score += weight
            reasons.append(f"contains {term}")
            matched_stem = True

    # Catch light typos in short identifier headers, e.g. "Subect ID".
    if has_identifier_token and not matched_stem:
        for token in tokens:
            close = difflib.get_close_matches(token, weighted_terms.keys(), n=1, cutoff=0.84)
            if close:
                term = close[0]
                score += weighted_terms[term] - 4
                reasons.append(f"contains likely {term}")
                matched_stem = True
                break

    # Do not treat survey questions as identifier columns just because they
    # contain words like "subject", "study", "recorded", "suicidal", or "valid".
    if not has_identifier_token:
        return 0, ""

    if not matched_stem:
        return 0, ""

    # Long labels can still be real identifier prompts, but keep them below
    # concise columns in the report unless they are exact matches.
    if len(tokens) > 8:
        score -= min(25, len(tokens) - 8)
        reasons.append("long header penalty")

    # Avoid generic labels where "id" appears inside unrelated words.
    if score < 45:
        return 0, ""

    return score, "; ".join(reasons)


def discover_candidate_columns(headers: Iterable[str]) -> list[CandidateColumn]:
    candidates = []
    for index, header in enumerate(headers):
        score, reason = candidate_score(header)
        if score:
            candidates.append(CandidateColumn(index=index, header=header, score=score, reason=reason))
    return candidates


def primary_record_id_column(headers: list[str]) -> CandidateColumn:
    if not headers:
        raise ValueError("CSV has no header row; cannot find primary Record ID column at index 0.")

    header = headers[0]
    if normalize_header(header) != "recordid":
        raise ValueError(
            f"Expected column index 0 to be 'Record ID' or 'record_id' for targeted discovery, found {header!r}. "
            "Use --all-candidate-columns to scan every ID-like column."
        )

    return CandidateColumn(
        index=0,
        header=header,
        score=100,
        reason="primary column index 0 Record ID",
    )


def selected_id_columns(headers: list[str], all_candidate_columns: bool = False) -> list[CandidateColumn]:
    if all_candidate_columns:
        return discover_candidate_columns(headers)
    return [primary_record_id_column(headers)]


def clean_record_id(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\ufeff", "").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def split_suffix(value: str) -> tuple[str, ...]:
    if not value:
        return ()
    return tuple(token for token in re.split(r"_+", value.strip("_")) if token)


def token_shape(token: str) -> str:
    if not token:
        return "EMPTY"
    if token.isdigit():
        return "DIGITS"
    if token.isalpha():
        return "WORD"
    if re.match(r"^[a-z]+\d+$", token):
        return "LETTERDIGITS"
    if re.match(r"^\d+[a-z]+$", token):
        return "DIGITSWORD"
    return "ALNUM"


def shape_signature(value: str) -> str:
    cleaned = clean_record_id(value).lower()
    if not cleaned:
        return "BLANK"
    separator_normalized = SEPARATOR_RE.sub("_", cleaned).strip("_")
    if not separator_normalized:
        return "BLANK"
    return "_".join(token_shape(token) for token in separator_normalized.split("_") if token)


def coded_format_for_family(format_family: str, suffix_tokens: tuple[str, ...] = ()) -> str:
    has_suffix = bool(suffix_tokens)
    if format_family == "numeric_prefix_subject_token":
        base = "DIGITS+[SPACER]+SUBJECT_TOKEN"
    elif format_family == "alpha_prefix_numeric_token":
        base = "WORD+[SPACER]+DIGITS"
    elif format_family == "subject_token_only":
        base = "SUBJECT_TOKEN"
    elif format_family == "numeric_only":
        base = "DIGITS"
    elif format_family == "literal_word":
        base = "WORD"
    elif format_family == "mixed_alphanumeric":
        return "ALPHANUMERIC_OTHER"
    elif format_family == "unknown":
        return "UNCLASSIFIED"
    elif format_family == "blank":
        return "BLANK"
    else:
        return format_family.upper()

    if has_suffix:
        return f"{base}+[SPACER]+SUFFIX"
    return base


def _canonical_from_parts(
    original: str,
    cleaned: str,
    format_family: str,
    prefix_numeric: str = "",
    prefix_alpha: str = "",
    subject_prefix: str = "",
    subject_number: str = "",
    numeric_token: str = "",
    suffix: str = "",
    review_reason: str = "",
) -> CanonicalRecordId:
    suffix_tokens = split_suffix(suffix)
    coded_format = coded_format_for_family(format_family, suffix_tokens)
    subject_token = f"{subject_prefix}{subject_number}" if subject_prefix or subject_number else ""

    parts: list[str] = []
    if prefix_numeric:
        parts.append(prefix_numeric)
    if prefix_alpha:
        parts.append(prefix_alpha)
    if subject_token:
        parts.append(subject_token)
    elif numeric_token:
        parts.append(numeric_token)
    parts.extend(suffix_tokens)
    canonical = "_".join(parts) if parts else cleaned.lower()

    normalized_for_comparison = SEPARATOR_RE.sub("_", cleaned.lower()).strip("_")
    needs_review = bool(review_reason)
    if canonical != normalized_for_comparison:
        needs_review = True
        review_reason = review_reason or "canonicalized separators or missing token boundary"
    if suffix_tokens:
        needs_review = True
        review_reason = review_reason or "suffix/modifier token present"

    return CanonicalRecordId(
        original=original,
        cleaned=cleaned,
        canonical=canonical,
        format_family=format_family,
        coded_format=coded_format,
        shape_signature=shape_signature(cleaned),
        prefix_numeric=prefix_numeric,
        prefix_alpha=prefix_alpha,
        subject_token=subject_token,
        subject_token_prefix=subject_prefix,
        subject_number=subject_number,
        suffix_tokens=suffix_tokens,
        needs_review=needs_review,
        review_reason=review_reason,
    )


def canonicalize_record_id(value: object) -> CanonicalRecordId:
    original = "" if value is None else str(value)
    cleaned = clean_record_id(value)
    lowered = cleaned.lower()
    normalized = SEPARATOR_RE.sub("_", lowered).strip("_")

    if not normalized:
        return _canonical_from_parts(
            original,
            cleaned,
            "blank",
            review_reason="blank ID value",
        )

    if normalized.isdigit():
        return _canonical_from_parts(
            original,
            cleaned,
            "numeric_only",
            numeric_token=normalized,
        )

    # 58807_s025, 58807-s025, 58807 s025 xover, 58807.sub025.ol
    match = re.match(r"^(\d+)_+([a-z]+)(\d+)(?:_+(.+))?$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "numeric_prefix_subject_token",
            prefix_numeric=match.group(1),
            subject_prefix=match.group(2),
            subject_number=match.group(3),
            suffix=match.group(4) or "",
        )

    # 58807_s025xover or 58807_s002a.
    match = re.match(r"^(\d+)_+([a-z]+?)(\d+)([a-z][a-z0-9_]*)$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "numeric_prefix_subject_token",
            prefix_numeric=match.group(1),
            subject_prefix=match.group(2),
            subject_number=match.group(3),
            suffix=match.group(4),
        )

    # 58807s025, 58807b025, 58807s025xover. This deliberately treats any
    # letters+digits token as a subject token, not only s###.
    match = re.match(r"^(\d+)([a-z]+?)(\d+)([a-z][a-z0-9_]*)?$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "numeric_prefix_subject_token",
            prefix_numeric=match.group(1),
            subject_prefix=match.group(2),
            subject_number=match.group(3),
            suffix=match.group(4) or "",
        )

    # MDD_10598, OCD_2780, LEAP_001, MDD_10598_duplicate.
    match = re.match(r"^([a-z]+)_+(\d+)(?:_+(.+))?$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "alpha_prefix_numeric_token",
            prefix_alpha=match.group(1),
            numeric_token=match.group(2),
            suffix=match.group(3) or "",
        )

    # MDD_10598new.
    match = re.match(r"^([a-z]+)_+(\d+)([a-z][a-z0-9_]*)$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "alpha_prefix_numeric_token",
            prefix_alpha=match.group(1),
            numeric_token=match.group(2),
            suffix=match.group(3),
        )

    # MDD10598 or MDD10598duplicate.
    match = re.match(r"^([a-z]+?)(\d+)([a-z][a-z0-9_]*)?$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "alpha_prefix_numeric_token",
            prefix_alpha=match.group(1),
            numeric_token=match.group(2),
            suffix=match.group(3) or "",
        )

    # s025, b025, sub025, sub025_ol.
    match = re.match(r"^([a-z]+)(\d+)(?:_+(.+))?$", normalized)
    if match:
        return _canonical_from_parts(
            original,
            cleaned,
            "subject_token_only",
            subject_prefix=match.group(1),
            subject_number=match.group(2),
            suffix=match.group(3) or "",
        )

    if re.match(r"^[a-z]+$", normalized):
        return _canonical_from_parts(
            original,
            cleaned,
            "literal_word",
            prefix_alpha=normalized,
            review_reason="literal word ID; likely admin/test/placeholder unless expected",
        )

    if ALNUM_RE.match(normalized):
        return _canonical_from_parts(
            original,
            cleaned,
            "mixed_alphanumeric",
            prefix_alpha=normalized,
            review_reason="mixed alphanumeric pattern not parsed into standard components",
        )

    return _canonical_from_parts(
        original,
        cleaned,
        "unknown",
        review_reason="unrecognized ID pattern",
    )


def classify_record_id(value: object) -> str:
    return canonicalize_record_id(value).format_family


def code_record_id(value: object) -> str:
    return canonicalize_record_id(value).coded_format


def find_event_column(headers: list[str]) -> int | None:
    for index, header in enumerate(headers):
        compact = normalize_header(header)
        if compact in {"eventname", "redcapeventname"}:
            return index
    return None


def read_csv_rows(path: Path) -> tuple[list[str], list[list[str]]]:
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
    return headers, rows


def profile_candidate_column(
    path: Path,
    candidate: CandidateColumn,
    rows: list[list[str]],
    event_column_index: int | None,
    max_examples: int = 8,
) -> tuple[ColumnProfile, list[dict[str, object]]]:
    values = []
    events_by_value: dict[str, Counter] = defaultdict(Counter)
    for row in rows:
        value = row[candidate.index].strip() if candidate.index < len(row) else ""
        values.append(value)
        if event_column_index is not None and event_column_index < len(row):
            event = row[event_column_index].strip()
            if event:
                events_by_value[value][event] += 1

    value_counts = Counter(values)
    canonical_by_value = {value: canonicalize_record_id(value) for value in value_counts}
    canonical_to_originals: dict[str, set[str]] = defaultdict(set)
    for value, canonical in canonical_by_value.items():
        canonical_to_originals[canonical.canonical].add(value)

    class_counts = Counter(
        canonical.format_family
        for canonical in canonical_by_value.values()
        for _ in range(value_counts[canonical.original])
    )
    coded_format_counts = Counter(
        canonical.coded_format
        for canonical in canonical_by_value.values()
        for _ in range(value_counts[canonical.original])
    )
    shape_counts = Counter(
        canonical.shape_signature
        for canonical in canonical_by_value.values()
        for _ in range(value_counts[canonical.original])
    )
    suffix_counts = Counter(
        suffix
        for canonical in canonical_by_value.values()
        for suffix in canonical.suffix_tokens
        for _ in range(value_counts[canonical.original])
    )
    review_count = sum(value_counts[value] for value, canonical in canonical_by_value.items() if canonical.needs_review)

    examples = []
    for value, count in value_counts.most_common(max_examples):
        canonical = canonical_by_value[value]
        examples.append(
            {
                "value": value,
                "count": count,
                "canonical": canonical.canonical,
                "format_family": canonical.format_family,
                "coded_format": canonical.coded_format,
                "shape_signature": canonical.shape_signature,
                "needs_review": canonical.needs_review,
            }
        )

    value_rows: list[dict[str, object]] = []
    for value in sorted(value_counts):
        canonical = canonical_by_value[value]
        event_counts = events_by_value.get(value, Counter())
        aliases = sorted(canonical_to_originals[canonical.canonical])
        value_rows.append(
            {
                "source_file": str(path),
                "column_index": candidate.index,
                "column_number": candidate.index + 1,
                "header": candidate.header,
                "original_value": value,
                "canonical_record_id": canonical.canonical,
                "format_family": canonical.format_family,
                "coded_format": canonical.coded_format,
                "shape_signature": canonical.shape_signature,
                "prefix_numeric": canonical.prefix_numeric,
                "prefix_alpha": canonical.prefix_alpha,
                "subject_token": canonical.subject_token,
                "subject_token_prefix": canonical.subject_token_prefix,
                "subject_number": canonical.subject_number,
                "suffix_tokens": ";".join(canonical.suffix_tokens),
                "row_count": value_counts[value],
                "event_count": len(event_counts),
                "event_examples": "; ".join(event for event, _ in event_counts.most_common(5)),
                "canonical_alias_count": len(aliases),
                "canonical_alias_examples": "; ".join(aliases[:8]),
                "needs_review": canonical.needs_review,
                "review_reason": canonical.review_reason,
            }
        )

    profile = ColumnProfile(
        source_file=str(path),
        column_index=candidate.index,
        header=candidate.header,
        score=candidate.score,
        reason=candidate.reason,
        total_rows=len(rows),
        nonblank_rows=sum(1 for value in values if value),
        unique_values=len(value_counts),
        duplicated_values=sum(1 for count in value_counts.values() if count > 1),
        rows_with_duplicated_values=sum(count for count in value_counts.values() if count > 1),
        class_counts=class_counts,
        coded_format_counts=coded_format_counts,
        shape_counts=shape_counts,
        suffix_counts=suffix_counts,
        review_count=review_count,
        examples=examples,
    )
    return profile, value_rows


def profile_csv(path: str | Path, max_examples: int = 8, all_candidate_columns: bool = False) -> CsvProfile:
    path = Path(path)
    headers, rows = read_csv_rows(path)
    candidates = selected_id_columns(headers, all_candidate_columns)
    event_column_index = find_event_column(headers)

    column_profiles: list[ColumnProfile] = []
    for candidate in candidates:
        profile, _ = profile_candidate_column(path, candidate, rows, event_column_index, max_examples)
        column_profiles.append(profile)

    return CsvProfile(
        source_file=str(path),
        total_rows=len(rows),
        total_columns=len(headers),
        event_column_index=event_column_index,
        event_column_header=headers[event_column_index] if event_column_index is not None else "",
        columns=column_profiles,
    )


def profile_csv_with_value_rows(
    path: Path,
    max_examples: int = 8,
    all_candidate_columns: bool = False,
) -> tuple[CsvProfile, list[dict[str, object]]]:
    headers, rows = read_csv_rows(path)
    candidates = selected_id_columns(headers, all_candidate_columns)
    event_column_index = find_event_column(headers)

    column_profiles: list[ColumnProfile] = []
    value_rows: list[dict[str, object]] = []
    for candidate in candidates:
        profile, candidate_value_rows = profile_candidate_column(
            path, candidate, rows, event_column_index, max_examples
        )
        column_profiles.append(profile)
        value_rows.extend(candidate_value_rows)

    return (
        CsvProfile(
            source_file=str(path),
            total_rows=len(rows),
            total_columns=len(headers),
            event_column_index=event_column_index,
            event_column_header=headers[event_column_index] if event_column_index is not None else "",
            columns=column_profiles,
        ),
        value_rows,
    )


def counter_to_string(counter: Counter, limit: int = 8) -> str:
    return "; ".join(f"{key}={value}" for key, value in counter.most_common(limit))


def format_summary_rows(
    source_file: str | Path,
    value_rows: list[dict[str, object]],
    scope: str = "primary_record_id_column_index_0",
    random_seed: int = 42,
    examples_per_format: int = 5,
) -> list[dict[str, object]]:
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in value_rows:
        grouped[str(row["coded_format"])].append(row)

    rng = random.Random(random_seed)
    summary_rows = []
    for coded_format, rows in grouped.items():
        unique_original_values = sorted({str(row["original_value"]) for row in rows})
        unique_canonical_ids = sorted({str(row["canonical_record_id"]) for row in rows})
        example_count = min(examples_per_format, len(unique_original_values))
        examples = rng.sample(unique_original_values, example_count) if example_count else []
        summary_rows.append(
            {
                "source_file": str(source_file),
                "scope": scope,
                "coded_format": coded_format,
                "unique_id_values": len(unique_original_values),
                "unique_canonical_ids": len(unique_canonical_ids),
                "value_rows": len(rows),
                "row_occurrences": sum(int(row["row_count"]) for row in rows),
                "examples": "; ".join(examples),
            }
        )

    return sorted(
        summary_rows,
        key=lambda row: (-int(row["unique_id_values"]), str(row["coded_format"])),
    )


def column_summary_rows(profile: CsvProfile) -> list[dict[str, object]]:
    rows = []
    for column in profile.columns:
        rows.append(
            {
                "source_file": column.source_file,
                "column_index": column.column_index,
                "column_number": column.column_index + 1,
                "header": column.header,
                "score": column.score,
                "reason": column.reason,
                "total_rows": column.total_rows,
                "nonblank_rows": column.nonblank_rows,
                "unique_values": column.unique_values,
                "duplicated_values": column.duplicated_values,
                "rows_with_duplicated_values": column.rows_with_duplicated_values,
                "review_count": column.review_count,
                "class_counts": counter_to_string(column.class_counts),
                "coded_format_counts": counter_to_string(column.coded_format_counts),
                "shape_counts": counter_to_string(column.shape_counts),
                "suffix_counts": counter_to_string(column.suffix_counts),
                "examples": json.dumps(column.examples, ensure_ascii=False),
            }
        )
    return rows


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fieldnames = list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def print_profile(profile: CsvProfile) -> None:
    print(f"\n{profile.source_file}")
    print(f"  rows: {profile.total_rows}")
    print(f"  columns: {profile.total_columns}")
    if profile.event_column_index is not None:
        print(f"  event column: #{profile.event_column_index + 1} {profile.event_column_header!r}")
    else:
        print("  event column: not detected")
    print(f"  candidate ID columns: {len(profile.columns)}")
    for column in profile.columns:
        print(
            f"    #{column.column_index + 1} {column.header!r}: "
            f"unique={column.unique_values}, nonblank={column.nonblank_rows}, "
            f"duplicated_ids={column.duplicated_values}, review_rows={column.review_count}"
        )
        print(f"      classes: {counter_to_string(column.class_counts)}")
        print(f"      coded formats: {counter_to_string(column.coded_format_counts)}")
        print(f"      shapes: {counter_to_string(column.shape_counts)}")
        if column.suffix_counts:
            print(f"      suffixes: {counter_to_string(column.suffix_counts)}")


def print_format_summary(summary_rows: list[dict[str, object]]) -> None:
    print("  coded format summary:")
    for row in summary_rows:
        print(
            f"    {row['coded_format']}: "
            f"unique_values={row['unique_id_values']}, "
            f"row_occurrences={row['row_occurrences']}, "
            f"examples={row['examples']}"
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Discover record/subject ID columns and naming formats in CSV files."
    )
    parser.add_argument("csv", nargs="+", type=Path, help="CSV file(s) to profile")
    parser.add_argument(
        "--out-dir",
        type=Path,
        help="Optional directory for audit CSVs: *_record_id_columns.csv and *_record_id_values.csv",
    )
    parser.add_argument(
        "--all-candidate-columns",
        action="store_true",
        help="Scan every ID-like candidate column instead of only column index 0 'Record ID'/'record_id'.",
    )
    parser.add_argument("--max-examples", type=int, default=8, help="Number of examples to show per column")
    parser.add_argument(
        "--summary-examples",
        type=int,
        default=5,
        help="Number of random examples to include for each coded-format summary row",
    )
    parser.add_argument(
        "--random-seed",
        type=int,
        default=42,
        help="Seed used for deterministic summary example selection",
    )
    parser.add_argument("--json", action="store_true", help="Print a JSON summary instead of text")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summaries = []
    for csv_path in args.csv:
        profile, value_rows = profile_csv_with_value_rows(
            csv_path,
            args.max_examples,
            all_candidate_columns=args.all_candidate_columns,
        )
        summary_scope = "all_candidate_id_columns" if args.all_candidate_columns else "primary_record_id_column_index_0"
        summary_rows = format_summary_rows(
            profile.source_file,
            value_rows,
            scope=summary_scope,
            random_seed=args.random_seed,
            examples_per_format=args.summary_examples,
        )
        summaries.append(
            {
                "source_file": profile.source_file,
                "total_rows": profile.total_rows,
                "total_columns": profile.total_columns,
                "event_column_index": profile.event_column_index,
                "event_column_header": profile.event_column_header,
                "candidate_columns": column_summary_rows(profile),
                "format_summary": summary_rows,
            }
        )
        if args.out_dir:
            safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", csv_path.stem)
            write_csv(args.out_dir / f"{safe_stem}_record_id_columns.csv", column_summary_rows(profile))
            write_csv(args.out_dir / f"{safe_stem}_record_id_values.csv", value_rows)
            write_csv(args.out_dir / f"{safe_stem}_record_id_format_summary.csv", summary_rows)
        if not args.json:
            print_profile(profile)
            print_format_summary(summary_rows)

    if args.json:
        print(json.dumps(summaries, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
