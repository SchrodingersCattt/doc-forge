"""Project-neutral CI gates for text and TeX source repositories."""

from .abbreviations import AbbreviationFinding, check_abbreviations, load_abbreviation_whitelist
from .references import ReferenceIssue, audit_references
from .style import StyleFinding, baseline_payload, check_style, load_style_config, unresolved

__all__ = [
    "AbbreviationFinding",
    "ReferenceIssue",
    "audit_references",
    "check_abbreviations",
    "load_abbreviation_whitelist",
    "StyleFinding",
    "baseline_payload",
    "check_style",
    "load_style_config",
    "unresolved",
]
