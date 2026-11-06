"""Explicit, table-scoped Markdown table notes.

Table notes are not Word page footnotes.  They are body paragraphs associated
with the immediately preceding Markdown table.  Two recognition modes exist:

* ``legacy`` (default): historical behavior — ``Table note:`` / ``Table foot:``
  always become ``table_foot`` blocks, and a line that begins with ``[a]``
  immediately after a table or another note is also treated as a note.
* ``explicit``: only ``Table note:`` / ``Table foot:`` declarations become
  notes, and only while the association window after a table is open.  An
  unmarked ``[a]`` paragraph remains ordinary prose.  Orphan declarations do
  not attach to an earlier unrelated table.

Escape a literal directive-looking line by prefixing a backslash::

    \\Table note: this stays an ordinary paragraph.

Superscript markers such as ``<sup>[a]</sup>`` inside table cells remain
authored inline text.  A matching ``Table note: [a] ...`` declaration makes
the relationship explicit for diagnostics; unmatched superscripts are left
alone.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, Sequence

from .blocks import Block

TABLE_NOTE_MODES = frozenset({"legacy", "explicit"})
TABLE_NOTE_DIAGNOSTIC_POLICIES = frozenset({"warn", "error", "ignore"})

TABLE_NOTE_DIRECTIVE_RE = re.compile(r"(?i)^table\s+(?:note|foot)\s*:\s*(.*)$")
TABLE_NOTE_ESCAPE_RE = re.compile(r"(?i)^\\table\s+(?:note|foot)\s*:\s*(.*)$")
LEGACY_BRACKET_NOTE_RE = re.compile(r"^\[([A-Za-z0-9]+)\]\s+(.*)$")
NOTE_LABEL_PREFIX_RE = re.compile(r"^\[([^\]]+)\]\s*(.*)$")
# Candidate note markers authored as HTML superscript inside table cells.
SUP_NOTE_REF_RE = re.compile(r"<sup>\s*\[([^\]]+)\]\s*</sup>", re.IGNORECASE)

__all__ = [
    "TableNoteDiagnostic",
    "TABLE_NOTE_MODES",
    "TABLE_NOTE_DIAGNOSTIC_POLICIES",
    "TABLE_NOTE_DIRECTIVE_RE",
    "TABLE_NOTE_ESCAPE_RE",
    "LEGACY_BRACKET_NOTE_RE",
    "extract_note_label",
    "table_cell_note_labels",
    "option_value",
    "with_options",
    "associate_table_notes",
    "diagnose_table_notes",
    "apply_table_note_diagnostics",
]


@dataclass(frozen=True)
class TableNoteDiagnostic:
    """One table-note recognition or label finding."""

    code: str
    message: str
    severity: str = "warning"
    table_id: str = ""
    label: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "code": self.code,
            "message": self.message,
            "severity": self.severity,
            "table_id": self.table_id,
            "label": self.label,
        }


def extract_note_label(text: str) -> tuple[str, str]:
    """Return ``(label, remainder)`` when *text* begins with ``[label]``."""

    match = NOTE_LABEL_PREFIX_RE.match((text or "").strip())
    if match is None:
        return "", (text or "").strip()
    return match.group(1).strip(), match.group(2).strip()


def table_cell_note_labels(rows: Sequence[Sequence[str]]) -> tuple[str, ...]:
    """Collect unique superscript ``[label]`` markers from table cells, in order."""

    seen: set[str] = set()
    labels: list[str] = []
    for row in rows:
        for cell in row:
            for match in SUP_NOTE_REF_RE.finditer(cell or ""):
                label = match.group(1).strip()
                if not label or label in seen:
                    continue
                seen.add(label)
                labels.append(label)
    return tuple(labels)


def option_value(block: Block, key: str, default: str = "") -> str:
    for name, value in block.options:
        if name == key:
            return value
    return default


def with_options(block: Block, **updates: str) -> Block:
    """Return a copy of *block* with selected option keys replaced."""

    mapping = dict(block.options)
    for key, value in updates.items():
        if value == "":
            mapping.pop(key, None)
        else:
            mapping[key] = value
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


def associate_table_notes(blocks: Sequence[Block]) -> list[Block]:
    """Attach ``table_id`` / ``label`` options to notes in an open table window.

    A table opens an association window.  Consecutive ``table_foot`` blocks stay
    in that window.  Any other block closes it.  Notes already marked
    ``orphan=true`` are left unassociated.
    """

    result: list[Block] = []
    open_table_id = ""
    table_counter = 0
    for block in blocks:
        if block.kind == "table":
            table_counter += 1
            open_table_id = str(table_counter)
            result.append(with_options(block, table_id=open_table_id))
            continue
        if block.kind == "table_foot":
            if option_value(block, "orphan") == "true" or not open_table_id:
                result.append(
                    with_options(
                        block,
                        table_id="",
                        orphan="true",
                        label=option_value(block, "label") or extract_note_label(block.text)[0],
                    )
                )
                open_table_id = ""
                continue
            label = option_value(block, "label") or extract_note_label(block.text)[0]
            result.append(with_options(block, table_id=open_table_id, label=label, orphan="false"))
            continue
        open_table_id = ""
        result.append(block)
    return result


def diagnose_table_notes(blocks: Sequence[Block]) -> list[TableNoteDiagnostic]:
    """Report orphan notes and missing/duplicate labels within each table scope."""

    findings: list[TableNoteDiagnostic] = []
    notes_by_table: dict[str, list[Block]] = {}
    tables: dict[str, Block] = {}
    for block in blocks:
        if block.kind == "table":
            table_id = option_value(block, "table_id")
            if table_id:
                tables[table_id] = block
            continue
        if block.kind == "paragraph" and option_value(block, "table_note_orphan") == "true":
            findings.append(
                TableNoteDiagnostic(
                    code="table_note.orphan",
                    message=f"table note is not associated with a preceding table: {block.text[:80]}",
                    label=option_value(block, "label") or extract_note_label(block.text)[0],
                )
            )
            continue
        if block.kind != "table_foot":
            continue
        if option_value(block, "orphan") == "true" or not option_value(block, "table_id"):
            findings.append(
                TableNoteDiagnostic(
                    code="table_note.orphan",
                    message=f"table note is not associated with a preceding table: {block.text[:80]}",
                    label=option_value(block, "label") or extract_note_label(block.text)[0],
                )
            )
            continue
        table_id = option_value(block, "table_id")
        notes_by_table.setdefault(table_id, []).append(block)

    for table_id, notes in notes_by_table.items():
        labels = [option_value(note, "label") or extract_note_label(note.text)[0] for note in notes]
        seen: dict[str, int] = {}
        for label in labels:
            if not label:
                # Unlabelled notes are valid presentation; no diagnostic.
                continue
            seen[label] = seen.get(label, 0) + 1
        for label, count in seen.items():
            if count > 1:
                findings.append(
                    TableNoteDiagnostic(
                        code="table_note.duplicate_label",
                        message=f"table {table_id} declares note label [{label}] {count} times",
                        table_id=table_id,
                        label=label,
                    )
                )
        table = tables.get(table_id)
        if table is None:
            continue
        declared = {label for label in labels if label}
        for label in table_cell_note_labels(table.rows):
            if label not in declared:
                findings.append(
                    TableNoteDiagnostic(
                        code="table_note.missing_label",
                        message=(
                            f"table {table_id} references [{label}] in a cell "
                            "without a matching Table note declaration"
                        ),
                        table_id=table_id,
                        label=label,
                    )
                )
    return findings


def apply_table_note_diagnostics(
    findings: Iterable[TableNoteDiagnostic],
    *,
    policy: str = "warn",
) -> tuple[TableNoteDiagnostic, ...]:
    """Apply *policy* to diagnostics: ignore, warn (return), or error (raise)."""

    if policy not in TABLE_NOTE_DIAGNOSTIC_POLICIES:
        raise ValueError("table_note_diagnostics must be one of: warn, error, ignore")
    items = tuple(findings)
    if policy == "ignore" or not items:
        return () if policy == "ignore" else items
    if policy == "error":
        summary = "; ".join(item.message for item in items)
        raise ValueError(f"table note diagnostics failed: {summary}")
    return items
