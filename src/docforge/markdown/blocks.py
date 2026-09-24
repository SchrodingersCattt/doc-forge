"""Block-level document model shared by the Markdown parser and renderer."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Block:
    """One structural unit of a Markdown document.

    Attributes:
        kind: heading | paragraph | ordered | bullet | reference | code |
            equation | table_caption | table | quote | image | separator
        text: textual payload (empty for tables)
        level: heading/list nesting level (0-based for lists, 1+ for headings)
        language: fenced-code info string
        path: resolved image path for image blocks
        rows: table cell matrix (row-major) for table blocks
    """

    kind: str
    text: str = ""
    level: int = 0
    rows: tuple[tuple[str, ...], ...] = ()
    language: str = ""
    path: str = ""
    options: tuple[tuple[str, str], ...] = ()

    def as_dict(self) -> dict:
        return {
            "kind": self.kind,
            "text": self.text,
            "level": self.level,
            "rows": [list(row) for row in self.rows],
            "language": self.language,
            "path": self.path,
            "options": [list(item) for item in self.options],
        }


@dataclass(frozen=True)
class SectionSource:
    """A named Markdown section with an ordering slot.

    Used by template-driven workflows that must
    rebuild the report body from a fixed set of Markdown files while keeping
    the official cover, tables, attachments, and commitment pages intact.
    """

    order: int
    heading: str
    blocks: tuple[Block, ...]
    path: Path


__all__ = ["Block", "SectionSource"]


def _unused() -> None:  # pragma: no cover - satisfies lint-only placeholder
    field  # imported for future block metadata use
