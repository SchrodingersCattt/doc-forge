"""Project-neutral CI gates for text and TeX source repositories."""

from .abbreviations import AbbreviationFinding, check_abbreviations, load_abbreviation_whitelist
from .references import ReferenceIssue, audit_references

__all__ = [
    "AbbreviationFinding",
    "ReferenceIssue",
    "audit_references",
    "check_abbreviations",
    "load_abbreviation_whitelist",
]

