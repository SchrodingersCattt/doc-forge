"""Strict mathematical LaTeX to native Word OMML conversion."""

from .pandoc import MathConversionError, latex_to_omml

__all__ = ["MathConversionError", "latex_to_omml"]