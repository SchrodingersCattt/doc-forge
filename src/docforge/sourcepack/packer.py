"""Lightweight, auditable TeX source ZIP from an explicit allowlist."""

from __future__ import annotations

import csv
import hashlib
import io
import subprocess
import zipfile
from datetime import datetime
from pathlib import Path

from ..output import validate_output_path


def sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def git_short_hash(root: Path) -> str:
    try:
        return subprocess.check_output(
            ["git", "-C", str(root), "rev-parse", "--short", "HEAD"],
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        return "nogit"


def package_files(
    root: Path,
    *,
    allowlist: list[str],
    glob_dirs: list[Path] | None = None,
    readme_parts: dict[str, str] | None = None,
    output: Path | None = None,
) -> Path:
    """Zip *allowlist* paths (relative to *root*) with a SHA-256 manifest CSV.

    *glob_dirs* (optional) are directories whose regular files are included
    recursively; *readme_parts* maps filename → text placed at the archive
    root (a README plus any sidecar text). Returns the output path.
    """
    files: list[Path] = [root / relative for relative in allowlist]
    for directory in glob_dirs or []:
        files.extend(p for p in directory.rglob("*") if p.is_file())
    missing = [path for path in files if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Missing package inputs: {missing}")
    files = sorted(set(files))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = output or root / "dist" / f"tex_source_{timestamp}_{git_short_hash(root)}.zip"
    output = output.resolve()
    validate_output_path(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    manifest_rows: list[tuple[str, int, str]] = []
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in files:
            relative = path.relative_to(root).as_posix()
            payload = path.read_bytes()
            archive.writestr(relative, payload)
            manifest_rows.append((relative, len(payload), sha256_bytes(payload)))
        for name, text in (readme_parts or {}).items():
            payload = text.encode("utf-8")
            archive.writestr(name, payload)
            manifest_rows.append((name, len(payload), sha256_bytes(payload)))

        buffer = io.StringIO()
        writer = csv.writer(buffer, lineterminator="\n")
        writer.writerow(["Path", "Size_bytes", "SHA256"])
        writer.writerows(manifest_rows)
        archive.writestr("SOURCE_PACKAGE_MANIFEST.csv", buffer.getvalue().encode("utf-8"))
    return output


__all__ = ["package_files", "sha256_bytes", "git_short_hash"]