"""Check that numbers and DOIs in a draft occur in a supplied source pack."""

from __future__ import annotations

import re
from pathlib import Path

DOI_RE = re.compile(r"\b10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", re.IGNORECASE)
NUMBER_RE = re.compile(
    r"(?i)(?:p\s*[<=>]\s*0?\.\d+)"
    r"|(?:n\s*=\s*\d+)"
    r"|(?:\d{1,3}(?:,\d{3})+(?:\.\d+)?)"
    r"|(?:\d+\.\d+)"
    r"|(?:\d+\s*(?:%|％|例|个|名|人|次|倍|℃|°C|K|mg|kg|g|mL|ml|mm|cm|nm))"
    r"|(?:\d{3,})"
)
TEXT_SUFFIXES = {".txt", ".md", ".csv", ".json", ".tex", ".tsv"}


def canon_number(raw: str) -> str | None:
    match = re.search(r"\d+(?:\.\d+)?", raw.replace(",", ""))
    if not match:
        return None
    number = match.group(0)
    if "." in number:
        number = number.rstrip("0").rstrip(".")
    return number


def attested_canons(source_text: str) -> set[str]:
    found = set()
    for match in NUMBER_RE.finditer(source_text):
        canon = canon_number(match.group(0))
        if canon:
            found.add(canon)
    return found


def attested_dois(source_text: str) -> set[str]:
    return {match.group(0).lower().rstrip(".,;)") for match in DOI_RE.finditer(source_text)}


def repair_unattested(text: str, source_text: str) -> tuple[str, list[str]]:
    """Replace numbers and DOIs that do not occur in source_text with [未提供]."""

    notes: list[str] = []
    dois = attested_dois(source_text)
    canons = attested_canons(source_text)

    def doi_repl(match: re.Match[str]) -> str:
        token = match.group(0).rstrip(".,;)")
        if token.lower() in dois:
            return match.group(0)
        notes.append(f"replaced unattested DOI {token} with [未提供]")
        return "[未提供]"

    replaced = DOI_RE.sub(doi_repl, text)

    def number_repl(match: re.Match[str]) -> str:
        raw = match.group(0)
        canon = canon_number(raw)
        if canon and canon in canons:
            return raw
        if canon:
            notes.append(f"replaced unattested value {raw.strip()} with [未提供]")
            return "[未提供]"
        return raw

    return NUMBER_RE.sub(number_repl, replaced), notes


def read_sources(paths: list[Path]) -> str:
    chunks: list[str] = []
    for path in paths:
        if path.is_dir():
            files = [item for item in path.rglob("*") if item.is_file() and item.suffix.lower() in TEXT_SUFFIXES]
            for child in sorted(files):
                chunks.append(child.read_text(encoding="utf-8", errors="replace"))
        elif path.is_file():
            chunks.append(path.read_text(encoding="utf-8", errors="replace"))
    return "\n".join(chunks)
