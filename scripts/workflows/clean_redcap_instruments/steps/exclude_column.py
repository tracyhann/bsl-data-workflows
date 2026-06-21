#!/usr/bin/env python3
"""Detect sensitive REDCap columns that should be excluded from cleaned workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Sequence


DEFAULT_EXCLUDE_KEYWORDS = ("email", "address", "contact", "mrn", "phone")


@dataclass(frozen=True)
class ExcludedColumn:
    column_index: int
    column_name: str
    column_label: str
    keyword: str


def _normalized_tokens(value: object) -> list[str]:
    text = str(value or "").lower()
    if "e-mail" in text or "e mail" in text or "e_mail" in text:
        text = re.sub(r"e[\s_-]*mail", "email", text)
    return [token for token in re.split(r"[^a-z0-9]+", text) if token]


def _matches_keyword(value: object, keyword: str) -> bool:
    tokens = _normalized_tokens(value)
    compact_value = "".join(tokens)
    compact_keyword = "".join(_normalized_tokens(keyword))
    if not compact_keyword:
        return False
    return compact_keyword in tokens or compact_keyword in compact_value


def flag_sensitive_columns(
    raw_headers: Sequence[object],
    label_headers: Sequence[object] | None = None,
    keywords: Sequence[str] = DEFAULT_EXCLUDE_KEYWORDS,
) -> dict[int, ExcludedColumn]:
    label_headers = label_headers or []
    flagged: dict[int, ExcludedColumn] = {}
    for index, raw_header in enumerate(raw_headers):
        label_header = label_headers[index] if index < len(label_headers) else ""
        text_values = [raw_header, label_header]
        for keyword in keywords:
            if any(_matches_keyword(value, keyword) for value in text_values):
                flagged[index] = ExcludedColumn(
                    column_index=index,
                    column_name=str(raw_header or ""),
                    column_label=str(label_header or ""),
                    keyword=keyword.lower(),
                )
                break
    return flagged
