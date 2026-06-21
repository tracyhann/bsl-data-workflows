"""Custom REDCap subject ID recognizer and standardizer.

The Presidio dependency is optional. The regex-based recognizer in this module
is used directly by the runner and can also be registered with Presidio when
``presidio_analyzer`` is installed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass


REDCAP_SUBJECT_ID = "REDCAP_SUBJECT_ID"
MAX_REDCAP_SUBID_CELL_CHARS = 25

SUBID_PATTERN_TEXT = (
    r"(?i)\b"
    r"(?:irb[\s:_-]*)?"
    r"(?P<irb>\d{4,6})"
    r"[\s._-]*"
    r"(?P<token>subject|sub|s|b)"
    r"[\s._-]*"
    r"0*(?P<num>\d{1,4})"
    r"(?P<attached_suffix>[a-z][a-z0-9]*)?"
    r"(?:[._-]+(?P<separated_suffix>[a-z][a-z0-9]*))?"
    r"\b"
)
SUBID_RE = re.compile(SUBID_PATTERN_TEXT)


@dataclass(frozen=True)
class RedcapSubidMatch:
    entity_type: str
    text: str
    start: int
    end: int
    score: float
    canonicalized: str


def _canonicalize_parts(irb: str, subject_number: str) -> str | None:
    if not irb or not subject_number:
        return None
    try:
        subject_int = int(subject_number)
    except ValueError:
        return None
    if subject_int <= 0:
        return None
    return f"{irb}_s{subject_int:03d}"


def find_redcap_subids(value: object, max_cell_chars: int | None = MAX_REDCAP_SUBID_CELL_CHARS) -> list[RedcapSubidMatch]:
    """Find REDCap subject IDs in a cell-like value."""
    text = "" if value is None else str(value).strip()
    if max_cell_chars is not None and len(text) > max_cell_chars:
        return []

    matches: list[RedcapSubidMatch] = []
    for match in SUBID_RE.finditer(text):
        canonicalized = _canonicalize_parts(match.group("irb"), match.group("num"))
        if not canonicalized:
            continue
        matches.append(
            RedcapSubidMatch(
                entity_type=REDCAP_SUBJECT_ID,
                text=match.group(0),
                start=match.start(),
                end=match.end(),
                score=0.9,
                canonicalized=canonicalized,
            )
        )
    return matches


def standardize_redcap_subid(value: object) -> str | None:
    """Return the first standardized REDCap subject ID, or None."""
    matches = find_redcap_subids(value)
    if not matches:
        return None
    return matches[0].canonicalized


def build_presidio_analyzer():
    """Build a Presidio AnalyzerEngine with the REDCap subid recognizer.

    This function imports Presidio lazily so the rest of this package remains
    usable in environments where ``presidio_analyzer`` is not installed.
    """
    try:
        from presidio_analyzer import AnalyzerEngine, Pattern, PatternRecognizer
    except ImportError as exc:
        raise RuntimeError(
            "presidio_analyzer is not installed. Install Microsoft Presidio to "
            "use build_presidio_analyzer(), or use find_redcap_subids() directly."
        ) from exc

    pattern = Pattern(name="redcap_subject_id", regex=SUBID_PATTERN_TEXT, score=0.9)
    recognizer = PatternRecognizer(
        supported_entity=REDCAP_SUBJECT_ID,
        patterns=[pattern],
        context=["record", "subject", "subid", "participant", "redcap", "irb"],
    )
    analyzer = AnalyzerEngine()
    analyzer.registry.add_recognizer(recognizer)
    return analyzer
