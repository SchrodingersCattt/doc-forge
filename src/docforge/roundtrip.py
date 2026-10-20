"""Portable cross-format round-trip contract primitives.

Adapters for DOCX, Markdown, presentations, spreadsheets, and TeX can share
these safety and reporting rules without sharing their package internals.  The
contract keeps exact reuse separate from semantic rebuilds: byte equality is
required only for an unchanged source, while changed models are compared by a
small normalized projection.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Mapping, Protocol, Sequence


INVARIANTS: tuple[str, ...] = (
    "RT-OPEN", "RT-EXACT", "RT-CONTENT", "RT-STRUCTURE", "RT-STYLE",
    "RT-GEOMETRY", "RT-REL", "RT-MEDIA", "RT-IDEMP", "RT-UNSUPPORTED",
)
_COUNT_KEY_RE = re.compile(r"^[a-z_]{1,40}$")


class RoundTripMode(str, Enum):
    EXACT_REUSE = "EXACT_REUSE"
    SOURCE_PATCH = "SOURCE_PATCH"
    REBUILD = "REBUILD"


class RoundTripError(ValueError):
    """Base class for contract failures with a stable public code."""

    code = "ROUNDTRIP_ERROR"

    def __init__(self, message: str, *, code: str | None = None) -> None:
        super().__init__(message)
        if code:
            self.code = code


class SnapshotMismatchError(RoundTripError):
    code = "SNAPSHOT_SHA_MISMATCH"


class PathCollisionError(RoundTripError):
    code = "PATH_COLLISION"


@dataclass(frozen=True)
class SourceSnapshot:
    """Opaque source snapshot metadata used before an exact-copy operation."""

    path: Path
    sha256: str

    @classmethod
    def from_file(cls, path: Path) -> "SourceSnapshot":
        path = Path(path)
        return cls(path, sha256_file(path))

    def verify(self) -> None:
        path = Path(self.path)
        if not path.is_file() or sha256_file(path) != self.sha256:
            raise SnapshotMismatchError("source snapshot is missing or its SHA-256 does not match")


@dataclass(frozen=True)
class InvariantResult:
    """One redacted invariant result suitable for a public manifest."""

    code: str
    status: str
    counts: Mapping[str, int] = field(default_factory=dict)
    diagnostics: tuple[Mapping[str, Any], ...] = ()

    def __post_init__(self) -> None:
        if self.code not in INVARIANTS:
            raise ValueError(f"unknown round-trip invariant: {self.code}")
        if self.status not in {"pass", "fail", "skip", "unsupported"}:
            raise ValueError("invariant status must be pass, fail, skip, or unsupported")

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "status": self.status,
            "counts": {str(key): int(value) for key, value in sorted(self.counts.items())},
            "diagnostics": [redact_diagnostic(item) for item in self.diagnostics],
        }


class RoundTripAdapter(Protocol):
    """Minimal adapter surface owned by each format implementation."""

    def probe(self, source: Path) -> str: ...
    def export(self, source: Path, workdir: Path) -> Any: ...
    def build(self, model: Any, output: Path) -> None: ...
    def project(self, path: Path) -> Mapping[str, Any]: ...
    def compare(self, before: Mapping[str, Any], after: Mapping[str, Any]) -> Mapping[str, "InvariantResult"]: ...
    def unsupported_features(self, path: Path) -> Sequence[str]: ...


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def ensure_distinct_paths(*paths: Path | None, source: Path | None = None) -> None:
    """Reject resolved source/output/report collisions before publication."""

    resolved = [Path(path).resolve() for path in paths if path is not None]
    if len(resolved) != len(set(resolved)):
        raise PathCollisionError("round-trip output paths must be distinct")
    if source is not None:
        source_resolved = Path(source).resolve()
        if source_resolved in resolved:
            raise PathCollisionError("round-trip source collides with an output path")


def confined_path(root: Path, relative: str | os.PathLike[str], *, label: str = "path") -> Path:
    """Resolve a manifest path while rejecting traversal and symlink escapes."""

    root = Path(root).resolve()
    candidate = (root / Path(relative)).resolve()
    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise RoundTripError(f"{label} escapes the workdir", code="PATH_OUTSIDE_WORKDIR") from exc
    current = root
    for part in candidate.relative_to(root).parts:
        current /= part
        if current.is_symlink():
            raise RoundTripError(f"{label} resolves through a symlink", code="PATH_OUTSIDE_WORKDIR")
    return candidate


def _redact_counts(value: Any) -> dict[str, int]:
    """Keep only small, stable integer counters from a public receipt."""

    if not isinstance(value, Mapping):
        return {}
    return {
        str(key): item
        for key, item in value.items()
        if isinstance(key, str)
        and _COUNT_KEY_RE.fullmatch(key)
        and isinstance(item, int)
        and not isinstance(item, bool)
    }


def redact_diagnostic(value: Any, *, _field: str | None = None) -> Any:
    """Keep public diagnostics stable and free of paths, text, XML, and hashes."""

    allowed = {"category", "ordinal", "reason", "counts", "code", "status", "severity"}
    if isinstance(value, Mapping):
        result: dict[str, Any] = {}
        for key, item in value.items():
            name = str(key)
            if name not in allowed:
                continue
            result[name] = _redact_counts(item) if name == "counts" else redact_diagnostic(item, _field=name)
        return result
    if isinstance(value, (list, tuple)):
        return [redact_diagnostic(item, _field=_field) for item in value]
    if isinstance(value, (int, float, bool)) or value is None:
        return value
    if isinstance(value, str):
        cleaned = re.sub(r"[^A-Za-z0-9_.:-]+", "_", value)[:80]
        if _field == "reason" and not re.fullmatch(r"[A-Z][A-Z0-9_.:-]{1,79}", cleaned):
            return "REDACTED"
        if _field == "category" and cleaned not in {
            "content", "structure", "style", "geometry", "relationships", "media", "other",
        }:
            return "other"
        if _field == "status" and cleaned not in {"pass", "fail", "skip", "unsupported"}:
            return "unsupported"
        if _field == "severity" and cleaned not in {"info", "warning", "error"}:
            return "info"
        if _field == "mode" and cleaned not in {
            "EXACT_REUSE", "SOURCE_PATCH", "REBUILD", "exact-reuse", "source-patch", "rebuild",
        }:
            return "REDACTED"
        if _field == "format" and cleaned not in {"docx", "docm", "dotx", "dotm", "pptx", "xlsx", "tex", "markdown"}:
            return "unknown"
        if _field == "code" and cleaned not in set(INVARIANTS) | {
            "PATH_COLLISION", "PATH_OUTSIDE_WORKDIR", "SNAPSHOT_SHA_MISMATCH",
            "SOURCE_SHA_MISMATCH", "EXACT_SHA_MISMATCH",
        }:
            return "REDACTED"
        return cleaned
    return None


def redact_report(report: Mapping[str, Any]) -> dict[str, Any]:
    """Return the portable public manifest subset for a contract run."""

    allowed = {"artifact_id", "format", "mode", "invariants", "counts", "unsupported_features", "diagnostics"}
    result: dict[str, Any] = {}
    for key, value in report.items():
        if key not in allowed:
            continue
        if key == "invariants" and isinstance(value, Mapping):
            invariants: dict[str, Any] = {}
            for code, item in value.items():
                code = str(code)
                if code not in INVARIANTS:
                    continue
                if isinstance(item, Mapping):
                    projected: dict[str, Any] = {"code": code}
                    if "status" in item:
                        projected["status"] = redact_diagnostic(item["status"], _field="status")
                    if "counts" in item:
                        projected["counts"] = _redact_counts(item["counts"])
                    if "diagnostics" in item:
                        projected["diagnostics"] = redact_diagnostic(item["diagnostics"], _field="diagnostics")
                    invariants[code] = projected
                else:
                    invariants[code] = {
                        "code": code,
                        "status": redact_diagnostic(item, _field="status"),
                        "counts": {},
                    }
            result[key] = invariants
        elif key == "counts":
            result[key] = _redact_counts(value)
        elif key == "unsupported_features":
            result[key] = ["RT-UNSUPPORTED"] if value else []
        else:
            result[key] = redact_diagnostic(value, _field=key)
    return result


def compare_projection(before: Mapping[str, Any], after: Mapping[str, Any]) -> dict[str, InvariantResult]:
    """Compare normalized adapter projections by the shared invariant names."""

    results: dict[str, InvariantResult] = {}
    for code, key in (
        ("RT-CONTENT", "content"), ("RT-STRUCTURE", "structure"),
        ("RT-STYLE", "style"), ("RT-GEOMETRY", "geometry"),
        ("RT-REL", "relationships"), ("RT-MEDIA", "media"),
    ):
        left, right = before.get(key), after.get(key)
        results[code] = InvariantResult(code, "pass" if left == right else "fail", {"equal": int(left == right)})
    results["RT-IDEMP"] = InvariantResult("RT-IDEMP", "pass" if before == after else "fail")
    return results


def exact_reuse(source: Path, output: Path, *, snapshot: SourceSnapshot | None = None, force: bool = False) -> InvariantResult:
    """Copy an unchanged source after snapshot verification and collision checks."""

    source, output = Path(source), Path(output)
    ensure_distinct_paths(output, source=source)
    if snapshot is not None:
        snapshot.verify()
        # The snapshot is the unchanged source baseline.  A valid snapshot
        # alone is insufficient when the live source was replaced after the
        # bundle was exported; refuse to label that copy as an exact reuse.
        if sha256_file(source) != snapshot.sha256:
            raise SnapshotMismatchError("source package differs from its snapshot")
    if output.exists() and not force:
        raise FileExistsError(f"output exists; pass force to overwrite: {output.name}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, output)
    if sha256_file(source) != sha256_file(output):
        raise RoundTripError("exact-copy output SHA-256 differs", code="EXACT_SHA_MISMATCH")
    return InvariantResult("RT-EXACT", "pass", {"equal": 1})


__all__ = [
    "INVARIANTS", "RoundTripMode", "RoundTripError", "SnapshotMismatchError", "PathCollisionError",
    "SourceSnapshot", "InvariantResult", "RoundTripAdapter", "sha256_file", "ensure_distinct_paths",
    "confined_path", "redact_diagnostic", "redact_report", "compare_projection", "exact_reuse",
]
