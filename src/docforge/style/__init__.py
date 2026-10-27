"""Programmable style profiles and the DOCX style engine."""

from .engine import restyle_docx, resolved_paragraph_styles
from .profile import (
    STYLE_SCHEMA,
    FontSpec,
    PageSpec,
    ParagraphSpec,
    RoleSpec,
    StylePolicy,
    StyleProfile,
    StyleProfileError,
    load_profile,
)

__all__ = [
    "STYLE_SCHEMA", "FontSpec", "PageSpec", "ParagraphSpec", "RoleSpec", "StylePolicy",
    "StyleProfile", "StyleProfileError", "load_profile", "restyle_docx", "resolved_paragraph_styles",
]
