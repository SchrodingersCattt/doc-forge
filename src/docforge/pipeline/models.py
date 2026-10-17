"""Versioned, project-neutral models for auditable document pipelines.

The pipeline package intentionally contains no manuscript-specific knowledge.
Consumers describe their documents and register small stages, gates, and
validators.  The models in this module are JSON friendly so a manifest can be
written without leaking absolute workspace paths or machine-specific state.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import uuid
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, Mapping, MutableMapping


PIPELINE_SCHEMA = "docforge.pipeline.v1"


def _json(value: Any) -> Any:
    """Return a deterministic JSON-compatible value."""

    if isinstance(value, Path):
        return "<absolute-path>" if value.is_absolute() else value.as_posix()
    if hasattr(value, "to_dict"):
        return value.to_dict()
    if isinstance(value, Mapping):
        return {str(k): _json(v) for k, v in sorted(value.items(), key=lambda item: str(item[0]))}
    if isinstance(value, (tuple, list, set, frozenset)):
        values = [_json(item) for item in value]
        return sorted(values, key=lambda item: json.dumps(item, sort_keys=True, default=str)) if isinstance(value, (set, frozenset)) else values
    if isinstance(value, float) and value != value:
        return "NaN"
    if isinstance(value, str):
        return _redact_text(value)
    return value


_POSIX_PATH_RE = re.compile(r"(?<![A-Za-z0-9_:/])/(?:[^/\s]+/)*[^/\s]+")
_WINDOWS_PATH_RE = re.compile(r"(?<![A-Za-z0-9_])[A-Za-z]:[\\/](?:[^\\/\s]+[\\/]?)*")


def _redact_text(value: str) -> str:
    """Redact absolute paths even when embedded in exception text."""

    if os.path.isabs(value):
        return "<absolute-path>"
    value = _POSIX_PATH_RE.sub("<absolute-path>", value)
    return _WINDOWS_PATH_RE.sub("<absolute-path>", value)


def canonical_json(value: Any) -> str:
    return json.dumps(_json(value), ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _fingerprint_details(value: Any) -> Any:
    """Drop unstable source coordinates while retaining diagnostic content."""

    if isinstance(value, Mapping):
        unstable = {"line", "column", "start_line", "end_line", "start_column", "end_column", "exception_message", "error_message"}
        return {str(key): _fingerprint_details(item) for key, item in value.items() if str(key) not in unstable}
    if isinstance(value, (list, tuple)):
        return [_fingerprint_details(item) for item in value]
    return _json(value)


def sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _portable_path(value: str | os.PathLike[str], *, label: str = "path") -> str:
    raw = str(value).replace("\\", "/")
    path = PurePosixPath(raw)
    if path.is_absolute() or re.match(r"^[A-Za-z]:/", raw):
        raise ValueError(f"{label} must be workspace-relative: {value}")
    parts: list[str] = []
    for part in path.parts:
        if part in {"", "."}:
            continue
        if part == "..":
            if not parts:
                raise ValueError(f"{label} must be workspace-relative: {value}")
            parts.pop()
        else:
            parts.append(part)
    normalized = "/".join(parts)
    if not normalized:
        raise ValueError(f"{label} must not be empty")
    return normalized


@dataclass(frozen=True)
class SourceRef:
    """A portable source location attached to an artifact or diagnostic."""

    artifact_id: str
    line: int | None = None
    column: int | None = None
    end_line: int | None = None
    end_column: int | None = None

    def __post_init__(self) -> None:
        if not str(self.artifact_id).strip():
            raise ValueError("source artifact_id must not be empty")
        for name in ("line", "column", "end_line", "end_column"):
            value = getattr(self, name)
            if value is not None and (not isinstance(value, int) or value < 1):
                raise ValueError(f"source {name} must be a positive integer")

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"artifact_id": self.artifact_id}
        for name in ("line", "column", "end_line", "end_column"):
            if getattr(self, name) is not None:
                payload[name] = getattr(self, name)
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, Any] | str) -> "SourceRef":
        if isinstance(value, str):
            return cls(value)
        return cls(
            artifact_id=str(value.get("artifact_id", value.get("id", ""))),
            line=value.get("line"), column=value.get("column"),
            end_line=value.get("end_line"), end_column=value.get("end_column"),
        )


@dataclass(frozen=True)
class Diagnostic:
    """A stable, content-addressed finding from a stage or validator."""

    severity: str
    code: str
    message: str
    source: SourceRef | None = None
    details: Mapping[str, Any] = field(default_factory=dict)
    fingerprint: str = field(init=False)

    def __post_init__(self) -> None:
        severity = str(self.severity).lower()
        if severity == "warn":
            severity = "warning"
        if severity not in {"info", "warning", "error"}:
            raise ValueError("diagnostic severity must be info, warning, or error")
        if not str(self.code).strip():
            raise ValueError("diagnostic code must not be empty")
        message = _redact_text(str(self.message))
        if not message.strip():
            raise ValueError("diagnostic message must not be empty")
        object.__setattr__(self, "severity", severity)
        object.__setattr__(self, "message", message)
        # Do not include line numbers, absolute paths, or timestamps in the
        # fingerprint.  A source artifact identity is stable across workspaces.
        basis = {"severity": severity, "code": self.code, "message": self.message, "details": _fingerprint_details(self.details)}
        if self.source is not None:
            basis["source"] = {"artifact_id": self.source.artifact_id}
        object.__setattr__(self, "fingerprint", sha256_bytes(canonical_json(basis).encode("utf-8"))[:32])

    def to_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "code": self.code,
            "message": self.message,
            "source": self.source.to_dict() if self.source else None,
            "details": _json(self.details),
            "fingerprint": self.fingerprint,
        }

    @property
    def location(self) -> SourceRef | None:
        return self.source

    @classmethod
    def from_dict(cls, value: Mapping[str, Any] | str, *, severity: str = "error", code: str = "pipeline.error") -> "Diagnostic":
        if isinstance(value, str):
            return cls(severity, code, value)
        source = value.get("source")
        return cls(
            str(value.get("severity", severity)), str(value.get("code", code)), str(value.get("message", "")),
            SourceRef.from_dict(source) if source else None, value.get("details", {}),
        )


@dataclass
class ArtifactRef:
    """Identity and integrity metadata for one workspace-relative artifact."""

    id: str
    role: str
    path: str
    media_type: str = "application/octet-stream"
    schema: str | None = None
    size: int | None = None
    sha256: str | None = None
    producer_stage: str | None = None
    metadata: MutableMapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not str(self.id).strip() or not str(self.role).strip():
            raise ValueError("artifact id and role must not be empty")
        self.path = _portable_path(self.path, label=f"artifact {self.id} path")
        if self.size is not None and (not isinstance(self.size, int) or self.size < 0):
            raise ValueError("artifact size must be a non-negative integer")
        if self.sha256 is not None and not re.fullmatch(r"[0-9a-f]{64}", self.sha256):
            raise ValueError("artifact sha256 must be a lowercase SHA-256 digest")

    @classmethod
    def from_path(cls, id: str, role: str, path: Path, *, root: Path, media_type: str = "application/octet-stream", schema: str | None = None, producer_stage: str | None = None, metadata: Mapping[str, Any] | None = None) -> "ArtifactRef":
        resolved = path.resolve()
        root_resolved = root.resolve()
        try:
            relative = resolved.relative_to(root_resolved).as_posix()
        except ValueError as exc:
            raise ValueError(f"artifact path escapes workspace: {path}") from exc
        if not resolved.is_file():
            raise FileNotFoundError(path)
        return cls(id, role, relative, media_type, schema, resolved.stat().st_size, sha256_file(resolved), producer_stage, dict(metadata or {}))

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"id": self.id, "role": self.role, "path": self.path, "media_type": self.media_type}
        for key in ("schema", "size", "sha256", "producer_stage"):
            value = getattr(self, key)
            if value is not None:
                payload[key] = value
        if self.metadata:
            payload["metadata"] = _json(self.metadata)
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "ArtifactRef":
        return cls(
            id=str(value.get("id", "")), role=str(value.get("role", "")), path=str(value.get("path", "")),
            media_type=str(value.get("media_type", value.get("mime_type", "application/octet-stream"))),
            schema=value.get("schema"), size=value.get("size", value.get("size_bytes")),
            sha256=value.get("sha256"), producer_stage=value.get("producer_stage", value.get("producer")),
            metadata=dict(value.get("metadata", {})),
        )

    @property
    def workspace_relative_path(self) -> str:
        """Explicit alias used by the public contract and manifests."""

        return self.path


def _diagnostics(values: Iterable[Diagnostic | Mapping[str, Any] | str], *, default_code: str) -> tuple[Diagnostic, ...]:
    result: list[Diagnostic] = []
    for value in values:
        result.append(value if isinstance(value, Diagnostic) else Diagnostic.from_dict(value, code=default_code))
    return tuple(result)


@dataclass
class ValidationReport:
    status: str
    diagnostics: tuple[Diagnostic, ...] = ()
    metrics: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.status = str(self.status).lower()
        if self.status == "warning":
            self.status = "warn"
        if self.status not in {"pass", "warn", "fail", "unavailable"}:
            raise ValueError("validation status must be pass, warn, fail, or unavailable")
        self.diagnostics = _diagnostics(self.diagnostics, default_code="validation.finding")

    @property
    def ok(self) -> bool:
        return self.status in {"pass", "warn"}

    @property
    def passed(self) -> bool:
        return self.ok

    def to_dict(self) -> dict[str, Any]:
        return {"status": self.status, "diagnostics": [item.to_dict() for item in self.diagnostics], "metrics": _json(self.metrics)}

    @classmethod
    def from_value(cls, value: Any, *, unavailable: bool = False) -> "ValidationReport":
        if isinstance(value, cls):
            return value
        if value is None:
            return cls("unavailable" if unavailable else "pass")
        if isinstance(value, bool):
            return cls("pass" if value else "fail")
        if isinstance(value, Mapping):
            status = value.get("status", "pass")
            return cls(str(status), tuple(value.get("diagnostics", ())), value.get("metrics", {}))
        raise TypeError(f"unsupported validation report: {type(value).__name__}")


@dataclass
class StageResult:
    artifacts: tuple[ArtifactRef, ...] = ()
    diagnostics: tuple[Diagnostic, ...] = ()
    provenance: Mapping[str, Any] = field(default_factory=dict)
    metrics: Mapping[str, Any] = field(default_factory=dict)
    cache_key: str | None = None
    status: str = "pass"

    def __post_init__(self) -> None:
        self.artifacts = tuple(item if isinstance(item, ArtifactRef) else ArtifactRef.from_dict(item) for item in self.artifacts)
        self.diagnostics = _diagnostics(self.diagnostics, default_code="stage.finding")
        self.status = str(self.status).lower()
        if self.status == "warning":
            self.status = "warn"
        if self.status not in {"pass", "warn", "fail"}:
            raise ValueError("stage status must be pass, warn, or fail")
        if any(item.severity == "error" for item in self.diagnostics):
            self.status = "fail"

    @property
    def ok(self) -> bool:
        return self.status in {"pass", "warn"}

    @property
    def outputs(self) -> tuple[ArtifactRef, ...]:
        return self.artifacts

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "artifacts": [item.to_dict() for item in self.artifacts],
            "diagnostics": [item.to_dict() for item in self.diagnostics],
            "provenance": _json(self.provenance), "metrics": _json(self.metrics), "cache_key": self.cache_key,
        }


@dataclass(frozen=True)
class DocumentSpec:
    id: str
    source: Mapping[str, Any]
    outputs: tuple[str, ...] = ()
    depends_on: tuple[str, ...] = ()
    numbering: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "DocumentSpec":
        return cls(str(value.get("id", "")), dict(value.get("source", {})), tuple(value.get("outputs", ())), tuple(value.get("depends_on", ())), dict(value.get("numbering", {})))

    def to_dict(self) -> dict[str, Any]:
        return {"id": self.id, "source": _json(self.source), "outputs": list(self.outputs), "depends_on": list(self.depends_on), "numbering": _json(self.numbering)}


@dataclass(frozen=True)
class StageSpec:
    id: str
    kind: str
    depends_on: tuple[str, ...] = ()
    consumes: tuple[str, ...] = ()
    produces: tuple[str, ...] = ()
    parameters: Mapping[str, Any] = field(default_factory=dict)
    required: bool = True

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "StageSpec":
        return cls(str(value.get("id", "")), str(value.get("kind", "")), tuple(value.get("depends_on", ())), tuple(value.get("consumes", ())), tuple(value.get("produces", ())), dict(value.get("parameters", value.get("config", {}))), bool(value.get("required", True)))

    def to_dict(self) -> dict[str, Any]:
        return {"id": self.id, "kind": self.kind, "depends_on": list(self.depends_on), "consumes": list(self.consumes), "produces": list(self.produces), "parameters": _json(self.parameters), "required": self.required}


@dataclass(frozen=True)
class GateSpec:
    id: str
    kind: str
    required: bool = True
    parameters: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "GateSpec":
        return cls(str(value.get("id", "")), str(value.get("kind", "")), bool(value.get("required", True)), dict(value.get("parameters", value.get("config", {}))))

    def to_dict(self) -> dict[str, Any]:
        return {"id": self.id, "kind": self.kind, "required": self.required, "parameters": _json(self.parameters)}


@dataclass(frozen=True)
class ReleasePolicy:
    required_artifacts: tuple[str, ...] = ()
    atomic: bool = True
    retain_previous: bool = True
    warnings_as_errors: bool = False

    @classmethod
    def from_dict(cls, value: Mapping[str, Any] | None) -> "ReleasePolicy":
        value = value or {}
        return cls(tuple(value.get("required_artifacts", ())), bool(value.get("atomic", True)), bool(value.get("retain_previous", True)), bool(value.get("warnings_as_errors", False)))

    def to_dict(self) -> dict[str, Any]:
        return {"required_artifacts": list(self.required_artifacts), "atomic": self.atomic, "retain_previous": self.retain_previous, "warnings_as_errors": self.warnings_as_errors}


@dataclass(frozen=True)
class PipelineSpec:
    schema: str = PIPELINE_SCHEMA
    documents: tuple[DocumentSpec, ...] = ()
    stages: tuple[StageSpec, ...] = ()
    gates: tuple[GateSpec, ...] = ()
    validators: tuple[GateSpec, ...] = ()
    release: ReleasePolicy = field(default_factory=ReleasePolicy)
    policy: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if self.schema != PIPELINE_SCHEMA:
            raise ValueError(f"unsupported pipeline schema: {self.schema}")

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "PipelineSpec":
        schema = str(value.get("schema", ""))
        return cls(
            schema=schema,
            documents=tuple(item if isinstance(item, DocumentSpec) else DocumentSpec.from_dict(item) for item in value.get("documents", ())),
            stages=tuple(item if isinstance(item, StageSpec) else StageSpec.from_dict(item) for item in value.get("stages", ())),
            gates=tuple(item if isinstance(item, GateSpec) else GateSpec.from_dict(item) for item in value.get("gates", ())),
            validators=tuple(item if isinstance(item, GateSpec) else GateSpec.from_dict(item) for item in value.get("validators", ())),
            release=ReleasePolicy.from_dict(value.get("release")), policy=dict(value.get("policy", {})),
        )

    @classmethod
    def from_file(cls, path: Path | str) -> "PipelineSpec":
        path = Path(path)
        if path.suffix.casefold() in {".yaml", ".yml"}:
            try:
                import yaml  # type: ignore
            except ImportError as exc:
                raise RuntimeError("YAML pipeline specs require the optional PyYAML dependency") from exc
            payload = yaml.safe_load(path.read_text(encoding="utf-8"))
        else:
            payload = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, Mapping):
            raise ValueError("pipeline spec must contain a mapping")
        return cls.from_dict(payload)

    @classmethod
    def from_json(cls, payload: str) -> "PipelineSpec":
        value = json.loads(payload)
        if not isinstance(value, Mapping):
            raise ValueError("pipeline JSON must contain a mapping")
        return cls.from_dict(value)

    def to_dict(self) -> dict[str, Any]:
        return {"schema": self.schema, "documents": [item.to_dict() for item in self.documents], "stages": [item.to_dict() for item in self.stages], "gates": [item.to_dict() for item in self.gates], "validators": [item.to_dict() for item in self.validators], "release": self.release.to_dict(), "policy": _json(self.policy)}

    def to_json(self, *, indent: int | None = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent) + "\n"

    def semantic_hash(self) -> str:
        return sha256_bytes(canonical_json(self.to_dict()).encode("utf-8"))


@dataclass
class StageContext:
    spec: PipelineSpec
    workspace: Path
    source_root: Path
    stage_dir: Path
    stage: StageSpec
    inputs: Mapping[str, ArtifactRef]
    artifacts: Mapping[str, ArtifactRef]
    parameters: Mapping[str, Any] = field(default_factory=dict)
    run_id: str = ""

    def source_path(self, relative: str | os.PathLike[str]) -> Path:
        return confined_path(self.source_root, relative, allow_missing=False)

    def output_path(self, relative: str | os.PathLike[str]) -> Path:
        return confined_path(self.stage_dir, relative, allow_missing=True)

    def artifact_path(self, artifact: ArtifactRef | str) -> Path:
        ref = self.artifacts[artifact] if isinstance(artifact, str) else artifact
        return self.output_path(ref.path)


def confined_path(root: Path, relative: str | os.PathLike[str], *, allow_missing: bool = True, reject_symlink: bool = True) -> Path:
    """Resolve *relative* under *root* and reject traversal/symlink escapes."""

    root = root.resolve()
    portable = _portable_path(relative)
    if reject_symlink:
        current = root
        for part in Path(portable).parts:
            current = current / part
            if current.is_symlink():
                # A link is rejected even when it happens to point back inside
                # the workspace: link targets are not stable across machines.
                raise ValueError(f"symlink paths are not allowed: {relative}")
    candidate = (root / portable).resolve(strict=False)
    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"path escapes workspace: {relative}") from exc
    if not allow_missing and not candidate.exists():
        raise FileNotFoundError(candidate)
    return candidate


@dataclass
class ArtifactStore:
    """Confined artifact index shared by stages and validators."""

    root: Path
    _items: dict[str, ArtifactRef] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.root = Path(self.root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)

    def add(self, artifact: ArtifactRef, *, verify: bool = True) -> ArtifactRef:
        if artifact.id in self._items:
            raise ValueError(f"duplicate artifact id: {artifact.id}")
        path = confined_path(self.root, artifact.path, allow_missing=not verify)
        if verify:
            if not path.is_file() or path.is_symlink():
                raise ValueError(f"artifact is not a regular file: {artifact.path}")
            artifact = ArtifactRef.from_path(artifact.id, artifact.role, path, root=self.root, media_type=artifact.media_type, schema=artifact.schema, producer_stage=artifact.producer_stage, metadata=artifact.metadata)
        self._items[artifact.id] = artifact
        return artifact

    def get(self, artifact_id: str) -> ArtifactRef:
        return self._items[artifact_id]

    def path(self, artifact: ArtifactRef | str) -> Path:
        ref = self.get(artifact) if isinstance(artifact, str) else artifact
        return confined_path(self.root, ref.path, allow_missing=False)

    def values(self) -> tuple[ArtifactRef, ...]:
        return tuple(self._items.values())

    def as_dict(self) -> dict[str, ArtifactRef]:
        return dict(self._items)


@dataclass(frozen=True)
class Workspace:
    """Portable roots used by a pipeline run."""

    root: Path
    output: Path | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "root", Path(self.root).resolve())
        if self.output is not None:
            object.__setattr__(self, "output", Path(self.output).resolve())

    def resolve(self, relative: str | os.PathLike[str]) -> Path:
        return confined_path(self.root, relative, allow_missing=False)


def new_run_id() -> str:
    return uuid.uuid4().hex


__all__ = [
    "PIPELINE_SCHEMA", "ArtifactRef", "SourceRef", "Diagnostic", "ValidationReport", "StageResult",
    "DocumentSpec", "StageSpec", "GateSpec", "ReleasePolicy", "PipelineSpec", "StageContext", "ArtifactStore", "Workspace",
    "canonical_json", "sha256_file", "sha256_bytes", "confined_path", "new_run_id",
]
