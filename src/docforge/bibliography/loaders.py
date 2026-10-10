"""Bibliography loaders for JSON and BibTeX sources."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .model import BibliographyEntry


def normalize_entries(
    entries: dict[str, BibliographyEntry | dict[str, Any] | str],
    *,
    source_format: str = "unknown",
) -> dict[str, BibliographyEntry]:
    result: dict[str, BibliographyEntry] = {}
    for key, value in entries.items():
        if not isinstance(key, str):
            raise ValueError(f"Bibliography entry key must be a string: {key!r}")
        if key.startswith("_"):
            continue
        result[key] = value if isinstance(value, BibliographyEntry) else BibliographyEntry.from_mapping(key, value, source_format=source_format)
    return result


def load_json(path: Path) -> dict[str, BibliographyEntry]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError(f"Bibliography must be a JSON object: {path}")
    return normalize_entries(payload, source_format="json")


def _bib_atom(text: str, index: int, macros: dict[str, str] | None = None) -> tuple[str, int]:
    """Read one BibTeX value atom and return it with its next cursor."""
    length = len(text)
    while index < length and text[index].isspace():
        index += 1
    if index >= length:
        return "", index
    if text[index] == "{":
        depth = 1
        index += 1
        begin = index
        pieces: list[str] = []
        while index < length and depth:
            character = text[index]
            if character == "\\":
                # Escaped braces are literal content.
                if index + 1 < length:
                    pieces.append(text[index : index + 2])
                    index += 2
                    continue
            if character == "{":
                depth += 1
            elif character == "}":
                depth -= 1
                if depth == 0:
                    break
            pieces.append(character)
            index += 1
        if depth:
            raise ValueError("Unclosed BibTeX field value")
        return "".join(pieces), index + 1
    if text[index] == '"':
        index += 1
        pieces: list[str] = []
        while index < length:
            character = text[index]
            if character == "\\" and index + 1 < length:
                pieces.append(text[index : index + 2])
                index += 2
                continue
            if character == '"':
                return "".join(pieces), index + 1
            pieces.append(character)
            index += 1
        raise ValueError("Unclosed BibTeX quoted field value")
    begin = index
    while index < length and text[index] not in ",}\n#":
        index += 1
    atom = text[begin:index].strip()
    if macros is not None:
        atom = macros.get(atom.lower(), atom)
    return atom, index


def _bib_value(text: str, index: int, macros: dict[str, str] | None = None) -> tuple[str, int]:
    """Read a BibTeX value, including ``#``-concatenated atoms.

    A braced or quoted atom may itself contain ``#``.  Only a hash outside an
    atom starts BibTeX concatenation, and a trailing hash is retained as
    literal text so malformed or legacy unbraced values are not silently
    truncated.
    """
    first, index = _bib_atom(text, index, macros)
    pieces = [first]
    while True:
        probe = index
        while probe < len(text) and text[probe].isspace():
            probe += 1
        if probe >= len(text) or text[probe] != "#":
            return "".join(pieces).strip(), index
        atom_start = probe + 1
        atom, next_index = _bib_atom(text, atom_start, macros)
        # A hash without a following atom is literal.  This keeps values such
        # as an unbraced ``C#`` field intact instead of dropping the suffix.
        if not atom and next_index == atom_start:
            return ("".join(pieces) + text[index:next_index]).strip(), next_index
        pieces.append(atom)
        index = next_index


def _bib_records(text: str) -> list[tuple[str, str, dict[str, str]]]:
    """Parse BibTeX entries with balanced braces and quoted values."""
    records: list[tuple[str, str, dict[str, str]]] = []
    macros: dict[str, str] = {}
    cursor = 0
    while True:
        marker = text.find("@", cursor)
        if marker < 0:
            break
        cursor = marker + 1
        type_start = cursor
        while cursor < len(text) and (text[cursor].isalnum() or text[cursor] in "_-"):
            cursor += 1
        entry_type = text[type_start:cursor]
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        if not entry_type or cursor >= len(text) or text[cursor] not in "{(":
            continue
        opener = text[cursor]
        closer = "}" if opener == "{" else ")"
        cursor += 1
        if entry_type.lower() == "string":
            # String declarations are definitions, not bibliography records.
            # Parse them before ordinary entry-key handling so a later value
            # such as ``journal # \" Letters\"`` is resolved in one pass.
            while cursor < len(text) and text[cursor].isspace():
                cursor += 1
            name_start = cursor
            while cursor < len(text) and (text[cursor].isalnum() or text[cursor] in "_:-"):
                cursor += 1
            macro_name = text[name_start:cursor].strip().lower()
            while cursor < len(text) and text[cursor].isspace():
                cursor += 1
            if macro_name and cursor < len(text) and text[cursor] == "=":
                value, cursor = _bib_value(text, cursor + 1, macros)
                macros[macro_name] = re.sub(r"\s+", " ", value).strip()
            while cursor < len(text) and text[cursor] != closer:
                cursor += 1
            if cursor < len(text):
                cursor += 1
            continue
        key_start = cursor
        while cursor < len(text) and text[cursor] not in "," + closer:
            cursor += 1
        key = text[key_start:cursor].strip()
        if cursor >= len(text) or text[cursor] == closer:
            continue
        cursor += 1
        fields: dict[str, str] = {"_type": entry_type}
        while cursor < len(text):
            while cursor < len(text) and (text[cursor].isspace() or text[cursor] == ","):
                cursor += 1
            if cursor >= len(text) or text[cursor] == closer:
                cursor += 1
                break
            name_start = cursor
            while cursor < len(text) and (text[cursor].isalnum() or text[cursor] in "_:-"):
                cursor += 1
            name = text[name_start:cursor].strip().lower()
            while cursor < len(text) and text[cursor].isspace():
                cursor += 1
            if not name or cursor >= len(text) or text[cursor] != "=":
                # Skip malformed field material to the next comma/entry end.
                while cursor < len(text) and text[cursor] not in "," + closer:
                    cursor += 1
                continue
            value, cursor = _bib_value(text, cursor + 1, macros)
            fields[name] = re.sub(r"\s+", " ", value).strip()
            while cursor < len(text) and text[cursor].isspace():
                cursor += 1
            if cursor < len(text) and text[cursor] == ",":
                cursor += 1
        if key:
            records.append((entry_type, key, fields))
    return records


def parse_bib_text(text: str) -> dict[str, BibliographyEntry]:
    entries: dict[str, BibliographyEntry] = {}
    for entry_type, key, fields in _bib_records(text):
        entries[key] = BibliographyEntry.from_mapping(key, fields, source_format="bibtex")
    return entries


def load_bib(path: Path) -> dict[str, BibliographyEntry]:
    return parse_bib_text(path.read_text(encoding="utf-8"))


def load(path: Path) -> dict[str, BibliographyEntry]:
    suffix = path.suffix.lower()
    if suffix in {".json", ".jsonl"}:
        return load_json(path)
    if suffix in {".bib", ".bibtex"}:
        return load_bib(path)
    raise ValueError(f"Unsupported bibliography format: {path.suffix or '<none>'}")

