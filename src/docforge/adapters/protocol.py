"""Format adapter contracts, diagnostics, and fidelity plans.

These types sit beside the pipeline runner. They reuse pipeline diagnostics
so a conversion report and a pipeline report share one finding schema.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Mapping, Protocol

from ..pipeline.models import Diagnostic, ValidationReport
from ..semantic import SemanticDocument


Fidelity = Literal["semantic", "preserve"]
LossPolicy = Literal["error", "warn", "drop"]
ExecutionMode = Literal["EXACT_REUSE", "SOURCE_PATCH", "REBUILD"]


class ConversionError(ValueError):
    """A conversion or restyle that must not publish an output."""

    def __init__(self, message: str, *, code: str, diagnostics: tuple[Diagnostic, ...] = ()) -> None:
        super().__init__(message)
        self.code = code
        self.diagnostics = diagnostics


@dataclass(frozen=True)
class FormatCapabilities:
    format_id: str
    node_kinds: tuple[str, ...]
    citations: bool
    media: bool
    annotations: bool
    preservation: tuple[str, ...]
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "format_id": self.format_id,
            "node_kinds": list(self.node_kinds),
            "citations": self.citations,
            "media": self.media,
            "annotations": self.annotations,
            "preservation": list(self.preservation),
            "notes": list(self.notes),
        }


@dataclass(frozen=True)
class SourceContext:
    """Verified source snapshot metadata. Not a second document model."""

    format_id: str
    sha256: str
    filename: str
    anchors: tuple[tuple[str, str], ...] = ()
    opaque_parts: tuple[str, ...] = ()
    retained_features: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "format_id": self.format_id,
            "sha256": self.sha256,
            "filename": self.filename,
            "anchors": [list(item) for item in self.anchors],
            "opaque_parts": list(self.opaque_parts),
            "retained_features": list(self.retained_features),
        }


@dataclass(frozen=True)
class ReadOptions:
    fidelity: Fidelity = "semantic"
    loss_policy: LossPolicy = "warn"


@dataclass(frozen=True)
class ReadResult:
    document: SemanticDocument
    source: SourceContext | None
    diagnostics: tuple[Diagnostic, ...]
    capabilities: FormatCapabilities
    origin_format: str
    origin_name: str
    content_hash: str
    origin: Path | None = None

    def unchanged(self, document: SemanticDocument | None = None) -> bool:
        current = self.document if document is None else document
        return current.semantic_hash() == self.content_hash


@dataclass(frozen=True)
class WriteOptions:
    fidelity: Fidelity = "semantic"
    loss_policy: LossPolicy = "warn"
    force: bool = False
    style: object | None = None


@dataclass(frozen=True)
class LossReport:
    policy: str
    diagnostics: tuple[Diagnostic, ...] = ()
    dropped: tuple[str, ...] = ()

    @property
    def errors(self) -> tuple[Diagnostic, ...]:
        return tuple(item for item in self.diagnostics if item.severity == "error")

    def allows_publication(self) -> bool:
        if self.policy == "drop":
            return True
        if self.policy == "warn":
            return True
        return not self.errors

    def to_dict(self) -> dict[str, object]:
        return {
            "policy": self.policy,
            "dropped": list(self.dropped),
            "diagnostics": [item.to_dict() for item in self.diagnostics],
        }


@dataclass(frozen=True)
class WritePlan:
    format_id: str
    execution_mode: ExecutionMode
    document: SemanticDocument
    source: SourceContext | None
    options: WriteOptions
    loss_report: LossReport
    origin: Path | None = None
    origin_format: str = ""
    legacy: Mapping[str, object] = field(default_factory=dict)


@dataclass(frozen=True)
class WriteResult:
    output: Path
    execution_mode: ExecutionMode
    loss_report: LossReport
    validation: ValidationReport
    manifest: Path | None = None

    def to_dict(self) -> dict[str, object]:
        return {
            "output_name": self.output.name,
            "execution_mode": self.execution_mode,
            "loss_report": self.loss_report.to_dict(),
            "validation": self.validation.to_dict(),
            "manifest_name": self.manifest.name if self.manifest else None,
        }


class FormatReader(Protocol):
    format_id: str

    def probe(self, source: Path) -> FormatCapabilities: ...

    def read(self, source: Path, *, options: ReadOptions) -> ReadResult: ...


class FormatWriter(Protocol):
    format_id: str

    def plan(self, document: SemanticDocument, *, source: SourceContext | None, options: WriteOptions) -> WritePlan: ...

    def write(self, plan: WritePlan, output: Path) -> WriteResult: ...


def finding(severity: str, code: str, message: str, **details: object) -> Diagnostic:
    return Diagnostic(severity, code, message, details=details)


def default_loss_policy(fidelity: Fidelity, explicit: LossPolicy | None) -> LossPolicy:
    if explicit is not None:
        return explicit
    return "error" if fidelity == "preserve" else "warn"


def ensure_publishable(report: LossReport) -> None:
    if report.allows_publication():
        return
    raise ConversionError(
        "conversion loss policy rejected unsupported constructs",
        code="LOSS_POLICY",
        diagnostics=report.errors,
    )


__all__ = [
    "ConversionError", "ExecutionMode", "Fidelity", "FormatCapabilities", "FormatReader",
    "FormatWriter", "LossPolicy", "LossReport", "ReadOptions", "ReadResult", "SourceContext",
    "WriteOptions", "WritePlan", "WriteResult", "default_loss_policy", "ensure_publishable", "finding",
]
