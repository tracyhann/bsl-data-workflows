"""Discover REDCap instrument blocks from column headers."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from typing import Iterable


INITIAL_METADATA_COLUMNS = {
    "record_id",
    "record id",
    "event name",
    "event_name",
    "redcap_event_name",
    "survey identifier",
    "survey_identifier",
    "redcap_survey_identifier",
}
COMMON_SUFFIXES = {
    "recordid",
    "record",
    "id",
    "date",
    "visitnumber",
    "visit",
    "score",
    "total",
    "details",
    "timestamp",
}


@dataclass(frozen=True)
class InstrumentBlock:
    instrument_order: int
    instrument_key: str
    start_index: int
    end_index: int
    column_count: int
    start_column: str
    stop_column: str
    dominant_prefix: str
    prefix_evidence: str
    suffix_evidence: str
    has_timestamp_start: bool


def _normalize_header(header: str) -> str:
    return str(header).strip().lower()


def is_complete_column(column_name: object) -> bool:
    """Return True only for terminal REDCap instrument status columns."""
    normalized = _normalize_header(str(column_name))
    return normalized.endswith("_complete") or normalized == "complete?"


def _instrument_key_from_stop(stop_column: str, block_columns: list[str]) -> str:
    normalized = str(stop_column).strip()
    if normalized.lower().endswith("_complete"):
        return normalized[: -len("_complete")]
    return "complete_status"


def _initial_data_start(headers: list[str]) -> int:
    index = 0
    while index < len(headers) and _normalize_header(headers[index]) in INITIAL_METADATA_COLUMNS:
        index += 1
    return index


def _tokens(column_name: str) -> list[str]:
    return [token for token in re.split(r"[_\s:/()-]+", column_name.strip().lower()) if token]


def _first_token(column_name: str) -> str:
    tokens = _tokens(column_name)
    return tokens[0] if tokens else ""


def _last_token(column_name: str) -> str:
    tokens = _tokens(column_name)
    return tokens[-1] if tokens else ""


def _format_counter(counter: Counter[str], limit: int = 5) -> str:
    return "; ".join(f"{name}:{count}" for name, count in counter.most_common(limit))


def _dominant_prefix(columns: Iterable[str]) -> str:
    counter = Counter(_first_token(column) for column in columns if _first_token(column))
    if not counter:
        return "NA"
    return counter.most_common(1)[0][0]


def _prefix_evidence(columns: Iterable[str]) -> str:
    counter = Counter(_first_token(column) for column in columns if _first_token(column))
    return _format_counter(counter)


def _suffix_evidence(columns: Iterable[str]) -> str:
    counter = Counter(
        token
        for column in columns
        if (token := _last_token(column)) and token not in COMMON_SUFFIXES
    )
    return _format_counter(counter)


def discover_instruments(headers: list[str]) -> list[InstrumentBlock]:
    """Discover contiguous instrument blocks closed by REDCap complete columns."""
    start_index = _initial_data_start(headers)
    blocks: list[InstrumentBlock] = []

    for index, header in enumerate(headers):
        if index < start_index or not is_complete_column(header):
            continue

        block_columns = headers[start_index : index + 1]
        if not block_columns:
            start_index = index + 1
            continue

        content_columns = block_columns[:-1]
        instrument_key = _instrument_key_from_stop(header, block_columns)
        blocks.append(
            InstrumentBlock(
                instrument_order=len(blocks) + 1,
                instrument_key=instrument_key,
                start_index=start_index,
                end_index=index,
                column_count=len(block_columns),
                start_column=block_columns[0],
                stop_column=header,
                dominant_prefix=_dominant_prefix(content_columns),
                prefix_evidence=_prefix_evidence(content_columns),
                suffix_evidence=_suffix_evidence(content_columns),
                has_timestamp_start=block_columns[0].strip().lower().endswith("_timestamp"),
            )
        )
        start_index = index + 1

    return blocks
