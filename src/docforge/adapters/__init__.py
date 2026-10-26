"""Format readers, writers, and the legacy conversion entry points."""

from .legacy import md2docx_namespace, run_md2docx, run_tex2docx, tex2docx_namespace
from .protocol import (
    ConversionError,
    FormatCapabilities,
    LossReport,
    ReadResult,
    SourceContext,
    WriteResult,
)

__all__ = [
    "ConversionError", "FormatCapabilities", "LossReport", "ReadResult", "SourceContext",
    "WriteResult", "md2docx_namespace", "run_md2docx", "run_tex2docx", "tex2docx_namespace",
]
