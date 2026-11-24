"""Project-neutral CI gates for text and TeX source repositories."""

from .abbreviations import AbbreviationFinding, check_abbreviations, load_abbreviation_whitelist
from .evidence import read_sources, repair_unattested
from .figures import check_figure_freshness, file_sha256
from .references import ReferenceIssue, audit_references
from .render import counts_as_pass, layout_gate
from .style import StyleFinding, baseline_payload, check_style, load_style_config, unresolved

__all__ = [
    "AbbreviationFinding",
    "ReferenceIssue",
    "audit_references",
    "counts_as_pass",
    "layout_gate",
    "check_abbreviations",
    "load_abbreviation_whitelist",
    "StyleFinding",
    "baseline_payload",
    "check_style",
    "load_style_config",
    "unresolved",
    "read_sources",
    "repair_unattested",
    "check_figure_freshness",
    "file_sha256",
]
