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


def parse_bib_text(text: str) -> dict[str, BibliographyEntry]:
    entries: dict[str, BibliographyEntry] = {}
    for match in re.finditer(r"@(\w+)\{([^,]+),\s*(.*?)\n\}", text, re.DOTALL):
        entry_type, key, body = match.group(1), match.group(2).strip(), match.group(3)
        fields: dict[str, Any] = {"_type": entry_type}
        for field in re.finditer(r"(\w+)\s*=\s*\{(.*?)\}(?:\s*,|\s*$)", body, re.DOTALL):
            fields[field.group(1).lower()] = re.sub(r"\s+", " ", field.group(2).strip())
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

