"""Compatibility exports for the project-agnostic TeX converter."""
from .converter import (
    ConversionContext, build_validated_docx, collect_citation_numbers, convert_files, convert_tex,
    latex_to_docx, validate_docx_package,
)

# Keep the established label-map entry point available to package consumers.
from .converter import _build_label_map as build_label_map

__all__ = [
    "ConversionContext", "build_label_map", "build_validated_docx", "collect_citation_numbers",
    "convert_files", "convert_tex", "latex_to_docx", "validate_docx_package",
]
