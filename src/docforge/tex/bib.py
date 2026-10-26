"""BibTeX parsing and TeX citation resolution."""
from __future__ import annotations
import re
from pathlib import Path
from .tokenize import plain_tex
from ..bibliography import CitationResolver as SharedCitationResolver
from ..bibliography.loaders import _bib_records

def parse_bib(path: Path) -> dict[str, dict]:
    """Parse a BibTeX file into ``{key: field-dict}``.

    The shared loader scans balanced field braces, so titles containing
    protected words or commas are not truncated at their first closing brace.
    """
    text = path.read_text(encoding="utf-8")
    return {
        key: fields
        for _entry_type, key, fields in _bib_records(text)
    }



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


class CitationResolver(SharedCitationResolver):
    """Use the shared numbering logic and TeX-aware reference formatting."""

    def citation_parts(self, keys: str, use_inherited: bool = True) -> list[tuple[str, str | None]]:
        """Collapse consecutive labels and link each range to its first reference.

        Two or more consecutive numbers become an en-dash range:
        ``5,6`` is ``5–6`` and ``2,5,6,7,8`` is ``2,5–8``.
        """
        labels = self.labels_for(keys, use_inherited=use_inherited)
        parts: list[tuple[str, str | None]] = []
        for index, group in enumerate(_group_labels(labels)):
            if index:
                parts.append((",", None))
            if len(group) == 1:
                parts.append((group[0][2], group[0][2]))
            else:
                parts.append((f"{group[0][2]}–{group[-1][2]}", group[0][2]))
        return parts

    def reference_items(self, *, profile: str = "plain") -> list[tuple[str, str]]:
        if profile != "plain":
            return super().reference_items(profile=profile)
        items: list[tuple[str, str]] = []
        for key in self.order:
            if key not in self.entries:
                continue
            entry = self.entries[key]
            authors = _author_short(str(entry.get("author", ""))).rstrip(".")
            title = plain_tex(str(entry.get("title", ""))).rstrip(".")
            journal = plain_tex(str(entry.get("journal", entry.get("booktitle", ""))))
            volume = str(entry.get("volume", ""))
            pages = plain_tex(str(entry.get("pages", "")))
            year = str(entry.get("year", ""))
            doi = str(entry.get("doi", ""))
            parts = [f"{authors}.", f"{title}."]
            if journal:
                venue = journal + (f" {volume}" if volume else "")
                venue += f", {pages}" if pages else ""
                parts.append(f"{venue} ({year}).")
            elif year:
                parts.append(f"({year}).")
            if doi:
                parts.append(f"https://doi.org/{doi}")
            items.append((f"{self.local_prefix}{self._key2num[key]}", " ".join(parts)))
        return items


def _citation_sort_key(label: str) -> tuple[int, int, str]:
    m = re.match(r"^([A-Za-z]*)(\d+)$", label)
    if not m:
        return (2, 0, label)
    prefix, num = m.group(1), int(m.group(2))
    if not prefix:
        return (0, num, label)
    return (1, num, prefix)


def _group_labels(labels: list[str]) -> list[list[tuple[str, int, str]]]:
    if not labels:
        return []
    unique = sorted(set(labels), key=_citation_sort_key)
    parsed: list[tuple[str, int, str]] = []
    for label in unique:
        m = re.match(r"^([A-Za-z]*)(\d+)$", label)
        if m:
            parsed.append((m.group(1), int(m.group(2)), label))
        else:
            parsed.append((label, -1, label))
    groups: list[list[tuple[str, int, str]]] = [[parsed[0]]]
    for item in parsed[1:]:
        prev = groups[-1][-1]
        if item[0] == prev[0] and item[1] == prev[1] + 1 and item[1] > 0:
            groups[-1].append(item)
        else:
            groups.append([item])
    return groups


def _compress_labels(labels: list[str]) -> str:
    groups = _group_labels(labels)
    if not groups:
        return ""
    parts = []
    for g in groups:
        if len(g) <= 2:
            parts.extend(item[2] for item in g)
        else:
            parts.append(f"{g[0][2]}–{g[-1][2]}")
    return "[" + ",".join(parts) + "]"


def _resolve_cites_in_text(text: str, resolver: CitationResolver) -> str:
    text = re.sub(r"\\mainref\{([^}]+)\}", lambda m: resolver.resolve(m.group(1)), text)
    return re.sub(r"\\cite\{([^}]+)\}", lambda m: resolver.resolve(m.group(1), use_inherited=False), text)

__all__=["parse_bib","CitationResolver","_author_short","_compress_labels"]
