"""LaTeX manuscript conversion, citation handling, and DOCX export."""
from .tokenize import Span, TableCell, tokenize_tex, spans_to_plain, plain_tex
from .bib import parse_bib, CitationResolver
from .convert import (
    ConversionContext, build_label_map, build_validated_docx, collect_citation_numbers, convert_files,
    convert_tex, latex_to_docx, validate_docx_package,
)
from .docx_export import TexExportResult, docx_to_tex

__all__ = [
    "Span", "TableCell", "tokenize_tex", "spans_to_plain", "plain_tex",
    "parse_bib", "CitationResolver", "ConversionContext", "TexExportResult",
    "docx_to_tex", "build_label_map", "build_validated_docx", "collect_citation_numbers", "convert_files",
    "convert_tex", "latex_to_docx", "validate_docx_package",
]
