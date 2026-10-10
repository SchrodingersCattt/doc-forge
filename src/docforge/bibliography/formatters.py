"""Pluggable bibliography formatting profiles."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from .model import BibliographyEntry

Formatter = Callable[[BibliographyEntry, str], str]


def _authors(entry: BibliographyEntry) -> str:
    value: Any = entry.get("authors", entry.get("author", ""))
    if isinstance(value, (list, tuple)):
        items = [str(item).strip() for item in value if str(item).strip()]
        if len(items) > 1:
            return ", ".join(items[:-1]) + ", and " + items[-1]
        return "".join(items)
    return str(value).replace(" and others", " et al.").strip()


def _sentence(text: str) -> str:
    return text if text.endswith(".") else f"{text}."



def _venue(entry: BibliographyEntry) -> str:
    """Return the most specific publication or event venue available."""
    for name in (
        "journal", "booktitle", "eventtitle", "conference", "event",
        "publisher", "institution", "school", "organization", "howpublished",
    ):
        value = entry.get(name, "")
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""

def plain(entry: BibliographyEntry, label: str = "") -> str:
    raw = entry.get("raw")
    if raw:
        return str(raw)
    authors = _authors(entry) or "unknown"
    title = str(entry.get("title", "")).strip()
    venue = _venue(entry)
    year = str(entry.get("year", "")).strip()
    volume = str(entry.get("volume", "")).strip()
    pages = str(entry.get("pages", entry.get("locator", ""))).strip()
    doi = str(entry.get("doi", "")).strip()
    parts = [_sentence(authors)]
    if title:
        parts.append(_sentence(title))
    if venue:
        suffix = f" {volume}" if volume else ""
        if pages:
            suffix += f", {pages}"
        if year:
            suffix += f" ({year})"
        parts.append(f"{venue}{suffix}.")
    elif year:
        parts.append(f"({year}).")
    if doi:
        parts.append(f"https://doi.org/{doi.removeprefix('https://doi.org/').removeprefix('doi:')}")
    return " ".join(parts)


def markdown(entry: BibliographyEntry, label: str = "") -> str:
    raw = entry.get("raw")
    if raw:
        return str(raw)
    authors = _authors(entry)
    if not authors:
        raise ValueError(f"Bibliography record {entry.key!r} requires authors/author")
    year = entry.get("year")
    if not isinstance(year, int) or isinstance(year, bool) or year < 0:
        raise ValueError(f"Bibliography record {entry.key!r} year must be a non-negative integer")
    venue = _venue(entry)
    if not venue:
        raise ValueError(f"Bibliography record {entry.key!r} requires a venue (journal, booktitle, or publisher)")
    parts = [_sentence(authors)]
    title = entry.get("title")
    if title:
        parts.append(_sentence(str(title).strip()))
    journal_part = f"*{venue}*, **{year}**"
    if entry.get("volume") not in (None, ""):
        journal_part += f", *{str(entry.get('volume')).strip()}*"
    if entry.get("issue") not in (None, ""):
        journal_part += f" ({str(entry.get('issue')).strip()})"
    locator = entry.get("locator", entry.get("pages"))
    if locator:
        journal_part += f", {str(locator).strip()}"
    parts.append(journal_part + ".")
    doi = entry.get("doi")
    if doi and entry.get("volume") in (None, "") and entry.get("issue") in (None, "") and not locator:
        parts.append(f"DOI: {str(doi).removeprefix('https://doi.org/').removeprefix('doi:').strip()}")
    return " ".join(parts)


FORMATTERS: dict[str, Formatter] = {"plain": plain, "markdown": markdown}


def format_entry(entry: BibliographyEntry, *, profile: str = "plain", label: str = "") -> str:
    try:
        formatter = FORMATTERS[profile]
    except KeyError as exc:
        raise ValueError(f"Unknown bibliography formatter profile: {profile}") from exc
    return formatter(entry, label)
