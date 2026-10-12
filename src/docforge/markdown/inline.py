"""Shared visible inline Markdown codec.

The reversible ``docxmd`` bundle owns the lossless parser.  Legacy Markdown
assembly imports the same display and matching helpers through this module so
captions, source matching, and DOCX text never grow separate parsers.
"""

from ..docxmd.inline import DisplayRun, Run, display, is_caption, is_filename_alt, key

__all__ = ["DisplayRun", "Run", "display", "is_caption", "is_filename_alt", "key"]
