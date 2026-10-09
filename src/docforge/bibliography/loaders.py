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


def _bib_value(text: str, index: int) -> tuple[str, int]:
    """Read one BibTeX value and return it with its next cursor position."""
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
        return "".join(pieces).strip(), index + 1
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
                return "".join(pieces).strip(), index + 1
            pieces.append(character)
            index += 1
        raise ValueError("Unclosed BibTeX quoted field value")
    begin = index
    while index < length and text[index] not in ",}\n":
        index += 1
    return text[begin:index].strip(), index


def _bib_records(text: str) -> list[tuple[str, str, dict[str, str]]]:
    """Parse BibTeX entries with balanced braces and quoted values."""
    records: list[tuple[str, str, dict[str, str]]] = []
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
            value, cursor = _bib_value(text, cursor + 1)
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

