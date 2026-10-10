"""Reversible DOCX <-> Markdown bundles (see docs/docxmd-syntax.md)."""

from .reader import ExportOptions, export_docx
from .verify import compare_docx, compare_pdfs, redline_changes, render_pdf, render_pdfs, roundtrip
from .writer import build_docx

__all__ = [
    "ExportOptions", "build_docx", "compare_docx", "compare_pdfs", "export_docx",
    "redline_changes", "render_pdf", "render_pdfs", "roundtrip",
]
