"""Explicit Markdown table layout markers.

Page span and vertical grouping are never inferred from the column count or
from repeated cell text. A manuscript opts in with a directive line or a
caption marker. The historical defaults remain column span, automatic
grouping, and fixed column fitting.
"""

from __future__ import annotations

import re
from typing import Mapping

from .blocks import Block


TABLE_SPANS = frozenset({"column", "page"})
TABLE_GROUPINGS = frozenset({"auto", "off"})
TABLE_WIDTHS = frozenset({"fixed", "proportional"})
_MARKER_VALUES = {
    "span": TABLE_SPANS,
    "grouping": TABLE_GROUPINGS,
    "widths": TABLE_WIDTHS,
}
_MARKER_KEYS = frozenset(_MARKER_VALUES)


def _reject(key: str, value: str) -> ValueError:
    allowed = ", ".join(sorted(_MARKER_VALUES[key]))
    return ValueError(f"Unsupported table {key} value: {value}. Expected one of: {allowed}.")


def split_caption_markers(text: str) -> tuple[str, dict[str, str]]:
    """Split ``Caption.|span=page|grouping=off`` into text plus markers.

    A pipe that is not a known ``key=value`` marker stays in the caption.
    """

    parts = [part.strip() for part in text.split("|")]
    if len(parts) == 1:
        return text.strip(), {}
    kept = [parts[0]]
    found: dict[str, str] = {}
    for part in parts[1:]:
        if "=" not in part:
            kept.append(part)
            continue
        key, value = (piece.strip().lower() for piece in part.split("=", 1))
        if key not in _MARKER_KEYS:
            kept.append(part)
            continue
        if value not in _MARKER_VALUES[key]:
            raise _reject(key, value)
        if key in found and found[key] != value:
            raise ValueError(f"Conflicting table {key} markers")
        found[key] = value
    return " | ".join(kept).strip(), found


def layout_directive(line: str) -> tuple[str, str] | None:
    """Return ``(key, value)`` for ``Table span: page`` and its siblings."""

    match = re.match(
        r"(?i)^table\s+(?P<key>span|grouping|widths)\s*:\s*(?P<value>\S+)\s*$",
        line.strip(),
    )
    if match is None:
        return None
    key = match.group("key").lower()
    value = match.group("value").lower()
    if value not in _MARKER_VALUES[key]:
        raise _reject(key, value)
    return key, value


def with_layout_options(block: Block, markers: Mapping[str, str]) -> Block:
    """Return *block* with layout markers merged into its options."""

    if not markers:
        return block
    mapping = dict(block.options)
    mapping.update(markers)
    return Block(
        kind=block.kind,
        text=block.text,
        level=block.level,
        rows=block.rows,
        language=block.language,
        path=block.path,
        options=tuple(sorted(mapping.items())),
        alignments=block.alignments,
    )


def option_value(block: Block | None, key: str) -> str:
    if block is None:
        return ""
    for name, value in block.options:
        if name == key:
            return value
    return ""


def resolved_table_option(table: Block | None, caption: Block | None, key: str, default: str) -> str:
    """Resolve one layout marker from the table, its caption, or the default."""

    if key not in _MARKER_VALUES:
        raise ValueError(f"Unknown table layout option: {key}")
    values = [value for value in (option_value(table, key), option_value(caption, key)) if value]
    unique = set(values)
    if len(unique) > 1:
        raise ValueError(f"Conflicting table {key} markers")
    resolved = values[0] if values else default
    if resolved not in _MARKER_VALUES[key]:
        raise _reject(key, resolved)
    return resolved


def grouping_is_auto(options: tuple[tuple[str, str], ...] | Mapping[str, str], default: str) -> bool:
    """Return whether automatic vertical grouping applies to one table."""

    if isinstance(options, Mapping):
        requested = str(options.get("grouping", "") or "")
    else:
        requested = ""
        for key, value in options:
            if key == "grouping":
                requested = value
                break
    mode = requested or default
    if mode not in TABLE_GROUPINGS:
        raise ValueError("table_grouping must be one of: auto, off")
    return mode == "auto"


__all__ = [
    "TABLE_GROUPINGS",
    "TABLE_SPANS",
    "TABLE_WIDTHS",
    "grouping_is_auto",
    "layout_directive",
    "option_value",
    "resolved_table_option",
    "split_caption_markers",
    "with_layout_options",
]
