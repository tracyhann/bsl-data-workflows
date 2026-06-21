"""Date recognizer and standardizer for REDCap-like table values."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date


DATE_ENTITY = "DATE"

MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

MONTH_PATTERN = "|".join(sorted(MONTHS, key=len, reverse=True))
MISSING_VALUES = {"", ".", "na", "n/a", "none", "null", "nan", "missing"}

TIME_SUFFIX = r"(?:[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\s*(?:am|pm))?)?"

ISO_RE = re.compile(
    rf"\b(?P<year>\d{{4}})[-/.](?P<month>\d{{1,2}})[-/.](?P<day>\d{{1,2}}){TIME_SUFFIX}\b",
    re.IGNORECASE,
)
COMPACT_ISO_RE = re.compile(r"\b(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})\b")
SLASH_RE = re.compile(
    rf"\b(?P<a>\d{{1,2}})[/-](?P<b>\d{{1,2}})[/-](?P<year>\d{{2,4}}){TIME_SUFFIX}\b",
    re.IGNORECASE,
)
MONTH_FIRST_RE = re.compile(
    rf"\b(?P<month>{MONTH_PATTERN})\.?\s+(?P<day>\d{{1,2}})(?:st|nd|rd|th)?[,]?\s+(?P<year>\d{{2,4}}){TIME_SUFFIX}\b",
    re.IGNORECASE,
)
DAY_FIRST_NAME_RE = re.compile(
    rf"\b(?P<day>\d{{1,2}})(?:st|nd|rd|th)?[-\s]+(?P<month>{MONTH_PATTERN})\.?[-,\s]+(?P<year>\d{{2,4}}){TIME_SUFFIX}\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class DateMatch:
    entity_type: str
    text: str
    start: int
    end: int
    score: float
    canonicalized: str
    parse_status: str


def _normalize_year(raw_year: str) -> int:
    year = int(raw_year)
    if len(raw_year) == 2:
        return 2000 + year if year <= 68 else 1900 + year
    return year


def _date_or_none(year: int, month: int, day: int) -> date | None:
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _iso_from_parts(year: str, month: str, day: str) -> str | None:
    parsed = _date_or_none(_normalize_year(year), int(month), int(day))
    return parsed.isoformat() if parsed else None


def _iso_from_numeric_month_day_year(a: str, b: str, year: str) -> str | None:
    first = int(a)
    second = int(b)

    if first > 31 or second > 31:
        return None
    if first > 12 and second <= 12:
        month, day = second, first
    else:
        month, day = first, second

    parsed = _date_or_none(_normalize_year(year), month, day)
    return parsed.isoformat() if parsed else None


def _month_number(raw_month: str) -> int:
    return MONTHS[raw_month.lower().rstrip(".")]


def _match_to_date(match: re.Match[str], kind: str) -> str | None:
    if kind in {"iso", "compact_iso"}:
        return _iso_from_parts(match.group("year"), match.group("month"), match.group("day"))
    if kind == "numeric":
        return _iso_from_numeric_month_day_year(match.group("a"), match.group("b"), match.group("year"))
    if kind == "month_first":
        return _iso_from_parts(str(_normalize_year(match.group("year"))), str(_month_number(match.group("month"))), match.group("day"))
    if kind == "day_first_name":
        return _iso_from_parts(str(_normalize_year(match.group("year"))), str(_month_number(match.group("month"))), match.group("day"))
    raise ValueError(f"Unsupported date pattern kind: {kind}")


def _candidate_matches(text: str):
    patterns = [
        ("iso", ISO_RE),
        ("compact_iso", COMPACT_ISO_RE),
        ("numeric", SLASH_RE),
        ("month_first", MONTH_FIRST_RE),
        ("day_first_name", DAY_FIRST_NAME_RE),
    ]
    for kind, pattern in patterns:
        for match in pattern.finditer(text):
            yield kind, match


def find_dates(value: object) -> list[DateMatch]:
    """Find parseable full dates and return their YYYY-MM-DD forms."""
    text = "" if value is None else str(value).strip()
    if text.lower() in MISSING_VALUES:
        return []

    matches: list[DateMatch] = []
    seen_spans: set[tuple[int, int]] = set()
    for kind, match in sorted(_candidate_matches(text), key=lambda item: item[1].start()):
        span = match.span()
        if span in seen_spans:
            continue
        canonicalized = _match_to_date(match, kind)
        if not canonicalized:
            continue
        seen_spans.add(span)
        matches.append(
            DateMatch(
                entity_type=DATE_ENTITY,
                text=match.group(0),
                start=match.start(),
                end=match.end(),
                score=0.9,
                canonicalized=canonicalized,
                parse_status="parsed",
            )
        )
    return matches


def standardize_date(value: object) -> str | None:
    """Return the first date in YYYY-MM-DD format, or None."""
    matches = find_dates(value)
    if not matches:
        return None
    return matches[0].canonicalized
