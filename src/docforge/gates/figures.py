"""Check that a figure manifest still matches the exported bytes."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 16), b""):
            digest.update(chunk)
    return digest.hexdigest()


def check_figure_freshness(manifest: dict[str, Any], root: Path) -> list[str]:
    """Return issues when a listed path is missing or its sha256 differs."""

    files = manifest.get("files")
    if not isinstance(files, list):
        return ["manifest files must be a list"]
    issues: list[str] = []
    for item in files:
        if not isinstance(item, dict):
            issues.append("manifest entry must be an object")
            continue
        rel = str(item.get("path") or "")
        expect = str(item.get("sha256") or "")
        path = root / rel
        if not rel or not path.is_file():
            issues.append(f"missing {rel or '(empty path)'}")
            continue
        if file_sha256(path) != expect:
            issues.append(f"stale {rel}")
    return issues
