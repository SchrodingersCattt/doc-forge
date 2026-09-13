"""LaTeX manuscript -> DOCX conversion."""

from .tokenize import Span, TableCell, tokenize_tex, spans_to_plain, plain_tex
from .bib import parse_bib, CitationResolver
from .convert import latex_to_docx, convert_files, build_label_map
from .docx_export import TexExportResult, docx_to_tex

__all__ = [
    "Span",
    "TableCell",
    "tokenize_tex",
    "spans_to_plain",
    "plain_tex",
    "parse_bib",
    "CitationResolver",
    "TexExportResult",
    "docx_to_tex",
    "latex_to_docx",
    "convert_files",
    "build_label_map",
]
