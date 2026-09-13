"""Small subprocess helpers for commands that require an installed Pandoc.

Pandoc is an external project prerequisite.  This module deliberately does
not discover, install, or replace it; callers invoke the executable named
``pandoc`` and receive an actionable error when it is unavailable.
"""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Sequence


def run_pandoc(
    input_path: Path,
    output_path: Path,
    *,
    output_format: str,
    extract_media_root: Path | None = None,
    track_changes: str = "accept",
    extra_args: Sequence[str] = (),
) -> tuple[str, ...]:
    """Run Pandoc for a DOCX input and return the exact command tuple.

    ``extract_media_root`` is passed directly to Pandoc's
    ``--extract-media`` option.  Pandoc writes the extracted files below a
    ``media`` directory in that root.
    """

    if track_changes not in {"accept", "reject", "all"}:
        raise ValueError("track_changes must be accept, reject, or all")
    command = [
        "pandoc",
        str(input_path),
        "-f",
        "docx",
        "-t",
        output_format,
        "--wrap=none",
        f"--track-changes={track_changes}",
    ]
    if extract_media_root is not None:
        command.extend(["--extract-media", str(extract_media_root)])
    command.extend(str(arg) for arg in extra_args)
    command.extend(["-o", str(output_path)])

    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except FileNotFoundError as exc:
        raise RuntimeError(
            "pandoc is required for DOCX conversion; install Pandoc and "
            "make the `pandoc` command available on PATH"
        ) from exc
    if completed.returncode:
        details = completed.stderr.strip() or completed.stdout.strip() or "no diagnostic output"
        raise RuntimeError(f"pandoc failed with exit code {completed.returncode}: {details}")
    return tuple(command)


__all__ = ["run_pandoc"]
