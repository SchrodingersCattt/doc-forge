"""Reusable and auditable manuscript pipeline primitives.

The public surface is deliberately small: versioned specs and data models,
protocols for domain stages/gates/validators, an explicit registry, and a
runner that owns safety checks, manifests, and publication.
"""

from .models import (
    PIPELINE_SCHEMA,
    ArtifactRef,
    ArtifactStore,
    Diagnostic,
    DocumentSpec,
    GateSpec,
    PipelineSpec,
    ReleasePolicy,
    SourceRef,
    StageContext,
    StageResult,
    StageSpec,
    ValidationReport,
    Workspace,
    canonical_json,
    confined_path,
    new_run_id,
    sha256_bytes,
    sha256_file,
)
from .runner import (
    CallableStage,
    Gate,
    PipelinePreflightError,
    PipelineResult,
    PipelineRunner,
    Registry,
    Stage,
    Validator,
)
from .adapters import ArtifactPresenceValidator, CallableGate, CallableValidator, CopyArtifactStage
from ..output import (
    INVALID_FORMAT,
    UNSUPPORTED_FORMAT,
    FormatError,
    InvalidFormatError,
    InputFormat,
    InvalidFormat,
    UnsupportedFormat,
    UnsupportedFormatError,
    detect_format,
    _detect_format,
)

PipelineRegistry = Registry

__all__ = [
    "PIPELINE_SCHEMA", "ArtifactRef", "ArtifactStore", "Diagnostic", "DocumentSpec", "GateSpec", "PipelineSpec",
    "ReleasePolicy", "SourceRef", "StageContext", "StageResult", "StageSpec", "ValidationReport", "Workspace",
    "Stage", "Gate", "Validator", "Registry", "PipelineRegistry", "CallableStage", "PipelineRunner",
    "PipelineResult", "PipelinePreflightError", "canonical_json", "confined_path", "new_run_id", "sha256_bytes", "sha256_file",
    "CallableGate", "CallableValidator", "CopyArtifactStage", "ArtifactPresenceValidator",
    "INVALID_FORMAT", "UNSUPPORTED_FORMAT", "FormatError", "InvalidFormatError", "InvalidFormat", "InputFormat",
    "UnsupportedFormat", "UnsupportedFormatError", "detect_format", "_detect_format",
]
