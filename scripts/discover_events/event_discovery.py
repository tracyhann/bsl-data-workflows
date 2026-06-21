"""Discover REDCap event arms and per-arm event order."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable


EVENT_COLUMN_NAMES = {"event name", "event_name", "redcap_event_name", "unique_event_name"}
ARM_WORDS = {
    "first": 1,
    "one": 1,
    "second": 2,
    "two": 2,
    "third": 3,
    "three": 3,
    "fourth": 4,
    "four": 4,
    "fifth": 5,
    "five": 5,
    "sixth": 6,
    "six": 6,
    "seventh": 7,
    "seven": 7,
    "eighth": 8,
    "eight": 8,
    "ninth": 9,
    "nine": 9,
    "tenth": 10,
    "ten": 10,
}
ARM_WORD_PATTERN = "|".join(sorted(ARM_WORDS, key=len, reverse=True))
ARM_VALUE_PATTERN = rf"\d+|{ARM_WORD_PATTERN}"
ARM_SUFFIX_PATTERN = r"(?:(?<=\d)[a-z][a-z0-9]*|(?<=\d)\s+[a-z]\b)?"
ARM_RE = re.compile(
    rf"(?i)(?<![A-Za-z0-9])arm[\s_:-]*(?P<arm>{ARM_VALUE_PATTERN}){ARM_SUFFIX_PATTERN}"
)
PAREN_ARM_RE = re.compile(
    rf"(?i)\s*\(\s*arm[\s_:-]*(?:{ARM_VALUE_PATTERN}){ARM_SUFFIX_PATTERN}(?::[^)]*)?\)"
)


@dataclass(frozen=True)
class ParsedEvent:
    raw_event: str
    arm: str
    event_name: str


@dataclass(frozen=True)
class EventGroup:
    arm: str
    event_order: int
    event_name: str
    first_raw_event: str
    first_row: int
    count: int


def _normalized_header(header: str) -> str:
    return str(header).strip().lower()


def resolve_event_column_index(
    headers: list[str],
    column_name: str | None = None,
    column_index: int = 1,
) -> int:
    """Resolve the event column, defaulting to raw REDCap index 1."""
    if column_name:
        if column_name in headers:
            return headers.index(column_name)
        lowered = column_name.lower()
        for index, header in enumerate(headers):
            if header.lower() == lowered:
                return index
        raise ValueError(f"Column name {column_name!r} was not found.")

    if len(headers) > 1 and _normalized_header(headers[1]) in EVENT_COLUMN_NAMES:
        return 1
    for name in ("redcap_event_name", "Event Name", "event_name", "unique_event_name"):
        lowered = name.lower()
        for index, header in enumerate(headers):
            if header.lower() == lowered:
                return index

    if column_index < 0 or column_index >= len(headers):
        raise ValueError(f"Column index {column_index} is out of range for {len(headers)} columns.")
    return column_index


def _arm_value(raw_arm: str) -> str:
    text = raw_arm.lower()
    if text.isdigit():
        return str(int(text))
    return str(ARM_WORDS[text])


def parse_arm(raw_event: object) -> str:
    text = "" if raw_event is None else str(raw_event)
    match = ARM_RE.search(text)
    if not match:
        return "NA"
    return _arm_value(match.group("arm"))


def _clean_event_name(raw_event: str) -> str:
    cleaned = PAREN_ARM_RE.sub("", raw_event)
    cleaned = ARM_RE.sub("", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"[_-]+$", "", cleaned).strip()
    cleaned = re.sub(r"^[_-]+", "", cleaned).strip()
    cleaned = re.sub(r"__+", "_", cleaned)
    return cleaned or raw_event.strip()


def parse_event_value(raw_event: object) -> ParsedEvent:
    text = "" if raw_event is None else str(raw_event).strip()
    return ParsedEvent(raw_event=text, arm=parse_arm(text), event_name=_clean_event_name(text))


def discover_event_groups(raw_events: Iterable[object]) -> list[EventGroup]:
    counts: dict[tuple[str, str], int] = {}
    first_seen: dict[tuple[str, str], tuple[int, str]] = {}

    for row_index, raw_event in enumerate(raw_events, start=1):
        parsed = parse_event_value(raw_event)
        key = (parsed.arm, parsed.event_name)
        counts[key] = counts.get(key, 0) + 1
        if key not in first_seen:
            first_seen[key] = (row_index, parsed.raw_event)

    arm_order = sorted(
        {key[0] for key in first_seen},
        key=lambda arm: (arm == "NA", int(arm) if arm.isdigit() else 10**9, arm),
    )
    groups: list[EventGroup] = []
    per_arm_order: dict[str, int] = {}
    ordered_keys: list[tuple[str, str]] = []
    for arm in arm_order:
        arm_keys = [key for key in first_seen if key[0] == arm]
        ordered_keys.extend(sorted(arm_keys, key=lambda key: first_seen[key][0]))

    for key in ordered_keys:
        arm, event_name = key
        per_arm_order[arm] = per_arm_order.get(arm, 0) + 1
        first_row, first_raw_event = first_seen[key]
        groups.append(
            EventGroup(
                arm=arm,
                event_order=per_arm_order[arm],
                event_name=event_name,
                first_raw_event=first_raw_event,
                first_row=first_row,
                count=counts[key],
            )
        )
    return groups
