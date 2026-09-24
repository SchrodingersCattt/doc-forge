"""Shared bibliography models, loaders, citation resolution, and formatters."""

from .formatters import FORMATTERS, format_entry
from .loaders import load, load_bib, load_json, normalize_entries, parse_bib_text
from .model import BibliographyEntry
from .resolver import CitationResolver, NumberingPolicy

__all__ = [
    "BibliographyEntry",
    "CitationResolver",
    "NumberingPolicy",
    "FORMATTERS",
    "format_entry",
    "load",
    "load_bib",
    "load_json",
    "normalize_entries",
    "parse_bib_text",
]
