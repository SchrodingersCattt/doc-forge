"""BibTeX parsing and citation resolution for the TeX converter."""

from __future__ import annotations

import re
from pathlib import Path


def parse_bib(path: Path) -> dict[str, dict]:
    """Parse a BibTeX file into {key: field-dict} (fields lower-cased)."""
    text = path.read_text(encoding="utf-8")
    entries: dict[str, dict] = {}
    for m in re.finditer(r"@(\w+)\{(\w+),\s*(.*?)\n\}", text, re.DOTALL):
        entry_type, key, body = m.group(1), m.group(2), m.group(3)
        fields: dict[str, str] = {"_type": entry_type}
        for fm in re.finditer(r"(\w+)\s*=\s*\{(.*?)\}(?:\s*,|\s*$)", body, re.DOTALL):
            fname = fm.group(1).lower()
            fval = re.sub(r"\s+", " ", fm.group(2).strip())
            fields[fname] = fval
        entries[key] = fields
    return entries


def _author_short(raw: str) -> str:
    raw = raw.replace(" and others", " et al.")
    parts = [a.strip() for a in raw.split(" and ") if a.strip()]
    if not parts:
        return raw
    first = parts[0].split(",")[0].strip()
    if len(parts) == 1:
        return first
    if len(parts) == 2:
        if parts[1].strip().startswith("et al"):
            return f"{first} et al."
        return f"{first} & {parts[1].split(',')[0].strip()}"
    return f"{first} et al."


class CitationResolver:
    """Assign stable numbers to cited keys in first-citation order."""

    def __init__(
        self,
        bib: dict[str, dict],
        inherited_numbers: dict[str, int] | None = None,
    ) -> None:
        self.bib = bib
        self.numbers: dict[str, int] = dict(inherited_numbers or {})

    def number_for(self, key: str) -> int:
        if key not in self.numbers:
            self.numbers[key] = len(self.numbers) + 1
        return self.numbers[key]

    def resolve(self, key: str) -> str:
        return f"[{self.number_for(key)}]"

    def format_bibliography(self) -> str:
        """Render the 'References' section from the resolver's citation order."""
        ordered = sorted(self.numbers.items(), key=lambda item: item[1])
        lines = ["## References"]
        for key, number in ordered:
            entry = self.bib.get(key)
            if entry is None:
                lines.append(f"{number}. {key} (missing BibTeX entry)")
                continue
            authors = _author_short(entry.get("author", entry.get("_type", "unknown")))
            year = entry.get("year", "")
            title = entry.get("title", key)
            journal = entry.get("journal", entry.get("booktitle", ""))
            volume = entry.get("volume", "")
            pages = entry.get("pages", "")
            location = ", ".join(part for part in (journal, volume, pages) if part)
            lines.append(f"{number}. {authors} ({year}). {title}. {location}".rstrip())
        return "\n".join(lines)


__all__ = ["parse_bib", "CitationResolver", "_author_short"]