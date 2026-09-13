"""Shared output-path safety checks for docforge exporters."""

from __future__ import annotations

from pathlib import Path


def validate_output_path(path: Path | None, *, label: str = "output") -> None:
    """Reject publish paths containing the forbidden ``final`` token."""
    if path is None:
        return
    if "final" in path.name.casefold():
        raise ValueError(f"{label} path contains forbidden token 'final': {path}")


__all__ = ["validate_output_path"]
