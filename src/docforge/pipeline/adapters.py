"""Small generic adapters for existing docforge operations.

These adapters deliberately do not import Markdown, TeX, DOCX, or plotting
implementations.  Consumers can wrap those functions with the same stable
stage/validator interfaces while keeping project-specific configuration out of
the pipeline core.
"""

from __future__ import annotations

import shutil
from collections.abc import Callable
from typing import Any

from .models import ArtifactRef, StageContext, StageResult, ValidationReport


class CallableGate:
    """Adapt a ``context -> ValidationReport`` function to a gate."""

    def __init__(self, name: str, function: Callable[[StageContext], Any]) -> None:
        self.name, self.function = name, function

    def validate(self, context: StageContext) -> ValidationReport:
        return ValidationReport.from_value(self.function(context))


class CallableValidator(CallableGate):
    """Alias with validator terminology for readability in registries."""


class CopyArtifactStage:
    """Copy one declared source file into a staged artifact."""

    def __init__(self, *, name: str, source: str, artifact_id: str, role: str, output: str | None = None, media_type: str = "application/octet-stream", schema: str | None = None) -> None:
        self.name = name
        self.source = source
        self.artifact_id = artifact_id
        self.role = role
        self.output = output or source
        self.media_type = media_type
        self.schema = schema

    def run(self, context: StageContext) -> StageResult:
        source = context.source_path(self.source)
        output = context.output_path(self.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, output)
        return StageResult((ArtifactRef(self.artifact_id, self.role, self.output, self.media_type, self.schema),), provenance={"source": self.source})


class ArtifactPresenceValidator:
    """Validate that required artifact ids or roles are present."""

    def __init__(self, required: tuple[str, ...] | list[str], *, name: str = "artifact-presence") -> None:
        self.required = tuple(required)
        self.name = name

    def validate(self, context: StageContext) -> ValidationReport:
        available = tuple(context.artifacts.values())
        missing = [key for key in self.required if not any(item.id == key or item.role == key for item in available)]
        if missing:
            from .models import Diagnostic
            return ValidationReport("fail", tuple(Diagnostic("error", "artifact.missing", f"required artifact is missing: {key}", details={"artifact": key}) for key in missing))
        return ValidationReport("pass", metrics={"checked": len(self.required)})


__all__ = ["CallableGate", "CallableValidator", "CopyArtifactStage", "ArtifactPresenceValidator"]
