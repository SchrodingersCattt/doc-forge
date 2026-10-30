"""Shared conservative policy for automatic table grouping."""
from __future__ import annotations

import re


def is_grouping_column(header: str) -> bool:
    """Allow label spans only in explicitly named grouping columns."""
    normalized = re.sub(r"[^A-Za-z]+", " ", header).strip().lower()
    safe = {"level", "group", "category", "class", "type", "family",
            "site", "site level", "factor", "block"}
    return normalized in safe or normalized.endswith(" level")
