"""Execution, validation, and publication for :mod:`docforge.pipeline`."""

from __future__ import annotations

import json
import inspect
import os
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, Protocol

from .models import (
    ArtifactRef,
    Diagnostic,
    GateSpec,
    PipelineSpec,
    ReleasePolicy,
    StageContext,
    StageResult,
    StageSpec,
    ValidationReport,
    _portable_path,
    canonical_json,
    confined_path,
    new_run_id,
    sha256_bytes,
    sha256_file,
)


class Stage(Protocol):
    name: str
    consumes: tuple[str, ...]
    produces: tuple[str, ...]

    def run(self, context: StageContext) -> StageResult:
        ...


class Gate(Protocol):
    name: str

    def validate(self, context: StageContext) -> ValidationReport:
        ...


class Validator(Protocol):
    name: str

    def validate(self, context: StageContext) -> ValidationReport:
        ...


class PipelinePreflightError(ValueError):
    """Raised when a spec cannot be safely executed."""


@dataclass
class Registry:
    """Explicit extension registry for stages, gates, and validators."""

    stages: dict[str, Any] = field(default_factory=dict)
    gates: dict[str, Any] = field(default_factory=dict)
    validators: dict[str, Any] = field(default_factory=dict)

    def register_stage(self, kind: str, stage: Any) -> "Registry":
        return self._register(self.stages, kind, stage, "stage")

    def register_gate(self, kind: str, gate: Any) -> "Registry":
        return self._register(self.gates, kind, gate, "gate")

    def register_validator(self, kind: str, validator: Any) -> "Registry":
        return self._register(self.validators, kind, validator, "validator")

    def _register(self, target: dict[str, Any], kind: str, value: Any, label: str) -> "Registry":
        key = str(kind).strip()
        if not key:
            raise ValueError(f"{label} registration kind must not be empty")
        if key in target:
            raise ValueError(f"duplicate {label} registration: {key}")
        target[key] = value
        return self

    def resolve(self, kind: str, category: str) -> Any:
        values = {"stage": self.stages, "gate": self.gates, "validator": self.validators}.get(category)
        if values is None:
            raise ValueError(f"unknown registry category: {category}")
        return values.get(kind)


class CallableStage:
    """Small adapter for registering a function as a pipeline stage."""

    def __init__(self, name: str, function: Callable[[StageContext], Any]) -> None:
        self.name = name
        self.function = function
        self.consumes = ()
        self.produces = ()

    def run(self, context: StageContext) -> StageResult:
        value = self.function(context)
        if isinstance(value, StageResult):
            return value
        if value is None:
            return StageResult()
        if isinstance(value, Mapping):
            return StageResult(
                artifacts=tuple(value.get("artifacts", ())),
                diagnostics=tuple(value.get("diagnostics", ())),
                provenance=value.get("provenance", {}), metrics=value.get("metrics", {}),
                cache_key=value.get("cache_key"), status=value.get("status", "pass"),
            )
        raise TypeError(f"stage {self.name} returned {type(value).__name__}, expected StageResult or mapping")


class CallableValidator:
    """Adapt a ``context -> ValidationReport`` callback."""

    def __init__(self, name: str, function: Callable[[StageContext], Any]) -> None:
        self.name = name
        self.function = function

    def validate(self, context: StageContext) -> ValidationReport:
        return ValidationReport.from_value(self.function(context))


@dataclass
class PipelineResult:
    success: bool
    artifacts: tuple[ArtifactRef, ...] = ()
    stage_results: Mapping[str, StageResult] = field(default_factory=dict)
    diagnostics: tuple[Diagnostic, ...] = ()
    manifest: Path | None = None
    report: str = ""
    publication_status: str = "not-published"
    cache: Mapping[str, Any] = field(default_factory=dict)
    toolchain: Mapping[str, Any] = field(default_factory=dict)
    run_id: str = ""

    @property
    def human_report(self) -> str:
        return self.report

    @property
    def manifest_path(self) -> Path | None:
        return self.manifest

    def to_dict(self) -> dict[str, Any]:
        return {
            "success": self.success,
            "artifacts": [item.to_dict() for item in self.artifacts],
            "stage_results": {key: value.to_dict() for key, value in self.stage_results.items()},
            "diagnostics": [item.to_dict() for item in self.diagnostics],
            "manifest": self.manifest.as_posix() if self.manifest else None,
            "report": self.report,
            "publication_status": self.publication_status,
            "cache": dict(self.cache),
            "toolchain": dict(self.toolchain),
            "run_id": self.run_id,
        }


def _instance(value: Any, *, kind: str) -> Any:
    if value is None:
        return None
    if isinstance(value, type):
        return value()
    if callable(value) and not hasattr(value, "run") and not hasattr(value, "validate"):
        try:
            signature = inspect.signature(value)
            required = [item for item in signature.parameters.values() if item.default is inspect.Parameter.empty and item.kind in (inspect.Parameter.POSITIONAL_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD)]
        except (TypeError, ValueError):
            required = []
        if required:
            return CallableStage(kind, value) if kind == "stage" else CallableValidator(kind, value)
        return value()
    return value


class PipelineRunner:
    """Run a :class:`PipelineSpec` with deterministic audit artifacts."""

    def __init__(self, registry: Registry | None = None, *, toolchain: Mapping[str, Any] | None = None) -> None:
        self.registry = registry or Registry()
        self.toolchain = dict(toolchain or {})

    def preflight(self, spec: PipelineSpec | Mapping[str, Any]) -> tuple[StageSpec, ...]:
        if not isinstance(spec, PipelineSpec):
            spec = PipelineSpec.from_dict(spec)
        errors: list[str] = []
        if spec.schema != "docforge.pipeline.v1":
            errors.append(f"unsupported pipeline schema: {spec.schema}")
        self._unique(errors, "document", [item.id for item in spec.documents])
        self._unique(errors, "stage", [item.id for item in spec.stages])
        self._unique(errors, "gate", [item.id for item in (*spec.gates, *spec.validators)])
        stage_ids = {item.id for item in spec.stages}
        known_kinds = set(self.registry.stages)
        for item in spec.stages:
            if item.kind not in known_kinds:
                errors.append(f"unknown stage registration: {item.kind}")
            errors.extend(f"stage {item.id} depends on unknown stage {dependency}" for dependency in item.depends_on if dependency not in stage_ids)
            if item.id in item.depends_on:
                errors.append(f"stage {item.id} depends on itself")
        for category, definitions, registry in (("gate", spec.gates, self.registry.gates), ("validator", spec.validators, self.registry.validators)):
            for item in definitions:
                if item.kind not in registry:
                    errors.append(f"unknown {category} registration: {item.kind}")
        for document in spec.documents:
            source = document.source
            if not isinstance(source, Mapping):
                errors.append(f"document {document.id} source must be a mapping")
            path = source.get("path") if isinstance(source, Mapping) else None
            if path is not None:
                try:
                    _portable_path(str(path), label=f"document {document.id} source path")
                except (TypeError, ValueError):
                    errors.append(f"document {document.id} source path must be relative: {path}")
        for name in spec.release.required_artifacts:
            if not str(name).strip():
                errors.append("release required_artifacts contains an empty value")
        produced_by: dict[str, str] = {}
        for stage in spec.stages:
            for output in stage.produces:
                produced_by.setdefault(str(output), stage.id)
        document_artifacts: set[str] = set()
        for document in spec.documents:
            document_artifacts.add(document.id)
            document_artifacts.update(document.outputs)
            if isinstance(document.source, Mapping):
                for key in ("id", "artifact_id", "path"):
                    value = document.source.get(key)
                    if value:
                        document_artifacts.add(str(value))
        dependency_map = {stage.id: set(stage.depends_on) for stage in spec.stages}

        def is_transitive_dependency(stage_id: str, producer_id: str) -> bool:
            pending = list(dependency_map.get(stage_id, ()))
            visited: set[str] = set()
            while pending:
                candidate = pending.pop()
                if candidate in visited:
                    continue
                visited.add(candidate)
                if candidate == producer_id:
                    return True
                pending.extend(dependency_map.get(candidate, ()))
            return False

        for stage in spec.stages:
            for consumed in stage.consumes:
                producer = produced_by.get(str(consumed))
                if producer is None:
                    if str(consumed) not in document_artifacts:
                        errors.append(f"stage {stage.id} consumes undeclared artifact: {consumed}")
                elif not is_transitive_dependency(stage.id, producer):
                    errors.append(f"stage {stage.id} consumes {consumed} without depending on producer {producer}")
        order = self._topological(spec.stages, errors)
        # ``produces`` is a declarative output namespace.  Duplicate names or
        # paths are rejected before any user stage can write to disk.
        produced: dict[str, str] = {}
        for stage in spec.stages:
            for output in stage.produces:
                owner = produced.setdefault(str(output), stage.id)
                if owner != stage.id:
                    errors.append(f"duplicate stage output {output!r}: {owner} and {stage.id}")
        if errors:
            raise PipelinePreflightError("pipeline preflight failed: " + "; ".join(errors))
        return tuple(order)

    @staticmethod
    def _unique(errors: list[str], label: str, values: Iterable[str]) -> None:
        seen: set[str] = set()
        for value in values:
            if not str(value).strip():
                errors.append(f"{label} id must not be empty")
            elif value in seen:
                errors.append(f"duplicate {label} id: {value}")
            seen.add(value)

    @staticmethod
    def _topological(stages: tuple[StageSpec, ...], errors: list[str]) -> list[StageSpec]:
        by_id = {item.id: item for item in stages}
        pending = {item.id: set(item.depends_on) for item in stages}
        result: list[StageSpec] = []
        while pending:
            ready = sorted((item for item, dependencies in pending.items() if not dependencies), key=lambda item: next(index for index, stage in enumerate(stages) if stage.id == item))
            if not ready:
                errors.append("stage dependency cycle detected")
                break
            for item in ready:
                result.append(by_id[item])
                pending.pop(item)
                for dependencies in pending.values():
                    dependencies.discard(item)
        return result

    def run(self, spec: PipelineSpec | Mapping[str, Any], workspace: Path, *, output_dir: Path | None = None, publish: bool = True, run_id: str | None = None) -> PipelineResult:
        if not isinstance(spec, PipelineSpec):
            spec = PipelineSpec.from_dict(spec)
        order = self.preflight(spec)
        source_root = Path(workspace).resolve()
        if not source_root.is_dir():
            raise FileNotFoundError(source_root)
        target = Path(output_dir or source_root / "dist").resolve()
        if target == source_root or target in source_root.parents:
            raise ValueError("output_dir must not equal or contain the workspace")
        if target.exists() and target.is_symlink():
            raise ValueError("output_dir must not be a symlink")
        rid = run_id or new_run_id()
        parent = target.parent
        parent.mkdir(parents=True, exist_ok=True)
        stage_root = Path(tempfile.mkdtemp(prefix=f".docforge-{rid}-", dir=str(parent)))
        all_diagnostics: list[Diagnostic] = []
        results: dict[str, StageResult] = {}
        artifacts: dict[str, ArtifactRef] = {}
        cache_hits = 0
        failed = False
        try:
            for stage in order:
                if failed:
                    results[stage.id] = StageResult(status="fail", diagnostics=(Diagnostic("error", "pipeline.stage_blocked", f"stage {stage.id} was blocked by an earlier failure"),))
                    continue
                implementation = _instance(self.registry.stages.get(stage.kind), kind="stage")
                if implementation is None:
                    result = StageResult(status="fail", diagnostics=(Diagnostic("error", "pipeline.unknown_stage", f"stage registration unavailable: {stage.kind}"),))
                    results[stage.id] = result
                    all_diagnostics.extend(result.diagnostics)
                    failed = True
                    continue
                inputs = self._select_inputs(stage, artifacts)
                if inputs is None:
                    result = StageResult(status="fail", diagnostics=(Diagnostic("error", "pipeline.missing_input", f"stage {stage.id} is missing a declared input"),))
                    results[stage.id] = result
                    all_diagnostics.extend(result.diagnostics)
                    failed = True
                    continue
                cache_key = self._stage_cache_key(stage, inputs)
                context = StageContext(spec, stage_root, source_root, stage_root, stage, inputs, artifacts, stage.parameters, rid)
                try:
                    value = implementation.run(context)
                    result = value if isinstance(value, StageResult) else CallableStage(stage.id, lambda _context: value).run(context)
                except Exception as exc:  # noqa: BLE001 - recorded as an auditable diagnostic
                    result = StageResult(status="fail", diagnostics=(Diagnostic("error", "pipeline.stage_exception", f"stage {stage.id} failed", details={"stage": stage.id, "exception_type": type(exc).__name__, "exception_message": str(exc)}),), cache_key=cache_key)
                try:
                    result = self._finalize_stage(result, stage, stage_root, artifacts, cache_key=cache_key)
                except Exception as exc:  # noqa: BLE001 - publication remains blocked and the error is recorded
                    result = StageResult(status="fail", diagnostics=(Diagnostic("error", "pipeline.artifact_output_invalid", f"stage {stage.id} produced an invalid artifact", details={"stage": stage.id, "exception_type": type(exc).__name__, "exception_message": str(exc)}),), cache_key=cache_key)
                results[stage.id] = result
                all_diagnostics.extend(result.diagnostics)
                for artifact in result.artifacts:
                    artifacts[artifact.id] = artifact
                if result.cache_key:
                    cache_hits += int(bool(result.provenance.get("cache_hit", False)))
                if not result.ok or (spec.release.warnings_as_errors and any(item.severity == "warning" for item in result.diagnostics)):
                    failed = True
            context = StageContext(spec, stage_root, source_root, stage_root, StageSpec("__pipeline__", "__pipeline__"), {}, artifacts, {}, rid)
            for category, definitions, registry in (("gate", spec.gates, self.registry.gates), ("validator", spec.validators, self.registry.validators)):
                for definition in definitions:
                    implementation = _instance(registry.get(definition.kind), kind=category)
                    report = self._run_validation(implementation, context, definition, category)
                    all_diagnostics.extend(report.diagnostics)
                    if report.status == "fail" or (report.status == "unavailable" and definition.required) or (spec.release.warnings_as_errors and report.status == "warn"):
                        failed = True
            artifact_list = tuple(sorted(artifacts.values(), key=lambda item: (item.role, item.id)))
            integrity_diagnostics = self._validate_artifacts(artifact_list, stage_root)
            all_diagnostics.extend(integrity_diagnostics)
            if integrity_diagnostics:
                failed = True
            missing = [role for role in spec.release.required_artifacts if not any(item.id == role or item.role == role for item in artifact_list)]
            all_diagnostics.extend(Diagnostic("error", "pipeline.required_artifact_missing", f"required artifact is missing: {role}") for role in missing)
            if missing:
                failed = True
            manifest_payload = self._manifest(spec, rid, artifact_list, results, all_diagnostics, failed, self.toolchain)
            manifest_path = stage_root / "manifest.json"
            report_text = self._report(manifest_payload)
            report_path = stage_root / "report.txt"
            checksums_path = stage_root / "checksums.sha256"
            manifest_path.write_text(json.dumps(manifest_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            report_path.write_text(report_text, encoding="utf-8")
            checksums_path.write_text(self._checksums(stage_root), encoding="utf-8")
            success = not failed and not any(item.severity == "error" for item in all_diagnostics)
            publication = "not-published"
            published_manifest = manifest_path
            if success and publish:
                self._publish(stage_root, target, retain_previous=spec.release.retain_previous)
                publication = "published"
                published_manifest = target / "manifest.json"
                stage_root = target
            elif success:
                publication = "staged"
            else:
                # Keep a portable failure receipt for audit/debugging while
                # leaving the previous publication untouched.
                failed_root = target.parent / f".docforge-failed-{rid}"
                if failed_root.exists():
                    if not self._has_publication_marker(failed_root):
                        raise ValueError(f"refusing to replace unmarked failure receipt: {failed_root.name}")
                    shutil.rmtree(failed_root)
                os.replace(stage_root, failed_root)
                publication = "failed"
                published_manifest = failed_root / "manifest.json"
            return PipelineResult(success, artifact_list, results, tuple(all_diagnostics), published_manifest, report_text, publication, {"stage_cache_hits": cache_hits}, self.toolchain, rid)
        except Exception:
            shutil.rmtree(stage_root, ignore_errors=True)
            raise

    @staticmethod
    def _select_inputs(stage: StageSpec, artifacts: Mapping[str, ArtifactRef]) -> dict[str, ArtifactRef] | None:
        if not stage.consumes:
            return {}
        result: dict[str, ArtifactRef] = {}
        for key in stage.consumes:
            found = next((item for item in artifacts.values() if item.id == key or item.role == key), None)
            if found is None:
                return None
            result[key] = found
        return result

    def _stage_cache_key(self, stage: StageSpec, inputs: Mapping[str, ArtifactRef]) -> str:
        return sha256_bytes(canonical_json({
            "stage": stage.to_dict(),
            "inputs": [{"id": item.id, "role": item.role, "sha256": item.sha256} for item in inputs.values()],
            "toolchain": self.toolchain,
        }).encode("utf-8"))

    @staticmethod
    def _finalize_stage(result: StageResult, stage: StageSpec, root: Path, existing: Mapping[str, ArtifactRef], *, cache_key: str) -> StageResult:
        outputs: list[ArtifactRef] = []
        seen_paths = {item.path for item in existing.values()}
        seen_ids = set(existing)
        declared = {str(item) for item in stage.produces}
        for item in result.artifacts:
            path = confined_path(root, item.path, allow_missing=False)
            if path.is_symlink():
                raise ValueError(f"artifact output must not be a symlink: {item.path}")
            normalized_path = path.relative_to(root.resolve()).as_posix()
            if normalized_path in {"manifest.json", "report.txt", "checksums.sha256"}:
                raise ValueError(f"artifact output collides with reserved audit file: {normalized_path}")
            if item.id not in declared and item.role not in declared:
                raise ValueError(f"artifact {item.id} is not declared by stage {stage.id}")
            if item.id in seen_ids:
                raise ValueError(f"duplicate artifact id: {item.id}")
            if normalized_path in seen_paths:
                raise ValueError(f"duplicate artifact output path: {normalized_path}")
            seen_ids.add(item.id)
            seen_paths.add(normalized_path)
            outputs.append(ArtifactRef.from_path(item.id, item.role, path, root=root, media_type=item.media_type, schema=item.schema, producer_stage=stage.id, metadata=item.metadata))
        provenance = dict(result.provenance)
        provenance.setdefault("stage", stage.id)
        return StageResult(tuple(outputs), result.diagnostics, provenance, result.metrics, cache_key, result.status)

    @staticmethod
    def _validate_artifacts(artifacts: Iterable[ArtifactRef], root: Path) -> tuple[Diagnostic, ...]:
        diagnostics: list[Diagnostic] = []
        for artifact in artifacts:
            try:
                path = confined_path(root, artifact.path, allow_missing=False)
                if path.is_symlink() or not path.is_file():
                    raise ValueError("artifact is not a regular file")
                size, digest = path.stat().st_size, sha256_file(path)
                if artifact.size != size or artifact.sha256 != digest:
                    diagnostics.append(Diagnostic("error", "pipeline.artifact_changed", f"artifact changed after its producing stage: {artifact.id}", details={"artifact_id": artifact.id}))
            except (FileNotFoundError, ValueError) as exc:
                diagnostics.append(Diagnostic("error", "pipeline.artifact_invalid", f"artifact {artifact.id} is invalid: {exc}", details={"artifact_id": artifact.id}))
        return tuple(diagnostics)

    @staticmethod
    def _run_validation(implementation: Any, context: StageContext, definition: GateSpec, category: str) -> ValidationReport:
        if implementation is None:
            severity = "error" if definition.required else "warning"
            code = f"pipeline.{category}_unavailable"
            return ValidationReport("unavailable", (Diagnostic(severity, code, f"{category} registration unavailable: {definition.kind}"),))
        try:
            value = implementation.validate(context)
            report = ValidationReport.from_value(value)
            if report.status == "unavailable" and not report.diagnostics:
                severity = "error" if definition.required else "warning"
                report.diagnostics = (Diagnostic(severity, f"pipeline.{category}_unavailable", f"{category} {definition.id} is unavailable"),)
            return report
        except Exception as exc:  # noqa: BLE001
            return ValidationReport("fail", (Diagnostic("error", f"pipeline.{category}_exception", f"{category} {definition.id} failed", details={"exception_type": type(exc).__name__, "exception_message": str(exc)}),))

    @staticmethod
    def _manifest(spec: PipelineSpec, run_id: str, artifacts: tuple[ArtifactRef, ...], results: Mapping[str, StageResult], diagnostics: Iterable[Diagnostic], failed: bool, toolchain: Mapping[str, Any] | None = None) -> dict[str, Any]:
        payload = {
            "schema": "docforge.pipeline.manifest.v1", "pipeline_schema": spec.schema, "run_id": run_id,
            "status": "fail" if failed else "pass", "spec_hash": spec.semantic_hash(),
            "stages": {key: value.to_dict() for key, value in results.items()},
            "artifacts": [item.to_dict() for item in artifacts],
            "diagnostics": [item.to_dict() for item in diagnostics],
            "release": spec.release.to_dict(), "toolchain": dict(toolchain or {}), "claim_boundaries": spec.policy.get("claim_boundaries", []),
            "limitations": spec.policy.get("limitations", []), "dependencies": spec.policy.get("dependencies", []),
        }
        semantic = dict(payload)
        semantic.pop("run_id", None)
        payload["semantic_hash"] = sha256_bytes(canonical_json(semantic).encode("utf-8"))
        return payload

    @staticmethod
    def _report(manifest: Mapping[str, Any]) -> str:
        diagnostics = manifest.get("diagnostics", [])
        lines = [f"docforge pipeline {manifest.get('status', 'unknown').upper()}", f"run_id: {manifest.get('run_id', '')}", f"artifacts: {len(manifest.get('artifacts', []))}"]
        if diagnostics:
            lines.append("diagnostics:")
            lines.extend(f"- {item.get('severity', 'error')} {item.get('code', '')}: {item.get('message', '')}" for item in diagnostics)
        else:
            lines.append("diagnostics: none")
        return "\n".join(lines) + "\n"

    @staticmethod
    def _checksums(root: Path) -> str:
        rows: list[str] = []
        for path in sorted(item for item in root.rglob("*") if item.is_file() and item.name not in {"checksums.sha256"}):
            rows.append(f"{sha256_file(path)}  {path.relative_to(root).as_posix()}")
        return "\n".join(rows) + ("\n" if rows else "")

    @staticmethod
    def _has_publication_marker(path: Path) -> bool:
        """Return true only for a directory produced by this pipeline."""

        if not path.is_dir() or path.is_symlink():
            return False
        marker = path / "manifest.json"
        if not marker.is_file() or marker.is_symlink():
            return False
        try:
            payload = json.loads(marker.read_text(encoding="utf-8"))
        except (OSError, ValueError, UnicodeError):
            return False
        return isinstance(payload, Mapping) and payload.get("schema") == "docforge.pipeline.manifest.v1"

    @classmethod
    def _publish(cls, stage_root: Path, target: Path, *, retain_previous: bool) -> None:
        if target.exists() and target.is_symlink():
            raise ValueError("output_dir must not be a symlink")
        backup = target.with_name(target.name + ".previous")
        for existing, label in ((target, "output_dir"), (backup, "previous publication")):
            if existing.exists() or existing.is_symlink():
                if not cls._has_publication_marker(existing):
                    raise ValueError(f"refusing to replace unmarked {label}: {existing}")
        if backup.exists():
            shutil.rmtree(backup)
        if target.exists():
            os.replace(target, backup)
        try:
            os.replace(stage_root, target)
        except Exception:
            if target.exists():
                shutil.rmtree(target, ignore_errors=True)
            if backup.exists():
                os.replace(backup, target)
            raise
        if not retain_previous and backup.exists():
            if not cls._has_publication_marker(backup):
                raise ValueError(f"refusing to remove unmarked previous publication: {backup}")
            shutil.rmtree(backup)


__all__ = ["Stage", "Gate", "Validator", "Registry", "CallableStage", "PipelineRunner", "PipelineResult", "PipelinePreflightError"]
