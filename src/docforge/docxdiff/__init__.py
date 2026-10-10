"""Word tracked-changes diffing with comment preservation."""

from .package import Package
from .redline import SparseRedlineError, create_tracked_docx

__all__ = ["Package", "SparseRedlineError", "create_tracked_docx"]
