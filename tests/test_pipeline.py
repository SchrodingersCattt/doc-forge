from __future__ import annotations

import json
from pathlib import Path

import pytest

from docforge.pipeline import (
    ArtifactRef,
    Diagnostic,
    GateSpec,
    PipelinePreflightError,
    PipelineRunner,
    PipelineSpec,
    Registry,
    SourceRef,
    StageResult,
    ValidationReport,
    confined_path,
)


class WriteStage:
    def __init__(self, filename: str, value: str) -> None:
        self.filename, self.value = filename, value

    def run(self, context):
        output = context.output_path(self.filename)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(self.value, encoding="utf-8")
        return StageResult((ArtifactRef("output", "document", self.filename, media_type="text/plain"),))


class Gate:
    def validate(self, context):
        return ValidationReport("pass", metrics={"checked": len(context.artifacts)})


def spec(*, stages=None, gates=None, release=None):
    return PipelineSpec.from_dict({
        "schema": "docforge.pipeline.v1",
        "stages": stages or [{"id": "write", "kind": "writer", "produces": ["document"]}],
        "gates": gates or [],
        "release": release or {"required_artifacts": ["document"], "atomic": True},
    })


def test_pipeline_publishes_manifest_and_hashes(tmp_path: Path):
    registry = Registry().register_stage("writer", WriteStage("nested/doc.txt", "hello"))
    registry.register_gate("style", Gate())
    pipeline = PipelineRunner(registry)
    result = pipeline.run(spec(gates=[{"id": "style", "kind": "style"}]), tmp_path, output_dir=tmp_path / "published", run_id="test-run")

    assert result.success
    assert result.publication_status == "published"
    assert result.manifest == tmp_path / "published" / "manifest.json"
    payload = json.loads(result.manifest.read_text(encoding="utf-8"))
    assert payload["schema"] == "docforge.pipeline.manifest.v1"
    assert payload["artifacts"][0]["path"] == "nested/doc.txt"
    assert len(payload["artifacts"][0]["sha256"]) == 64
    assert (tmp_path / "published" / "checksums.sha256").is_file()


def test_preflight_rejects_cycles_unknowns_and_duplicate_outputs(tmp_path: Path):
    registry = Registry().register_stage("writer", WriteStage("x", "x"))
    runner = PipelineRunner(registry)
    cyclic = spec(stages=[
        {"id": "a", "kind": "writer", "depends_on": ["b"], "produces": ["x"]},
        {"id": "b", "kind": "writer", "depends_on": ["a"], "produces": ["y"]},
    ])
    with pytest.raises(PipelinePreflightError, match="cycle"):
        runner.preflight(cyclic)
    duplicate = spec(stages=[
        {"id": "a", "kind": "writer", "produces": ["x"]},
        {"id": "b", "kind": "writer", "produces": ["x"]},
    ])
    with pytest.raises(PipelinePreflightError, match="duplicate stage output"):
        runner.preflight(duplicate)


def test_failed_stage_keeps_previous_publication(tmp_path: Path):
    registry = Registry().register_stage("writer", WriteStage("doc.txt", "new"))
    target = tmp_path / "published"
    target.mkdir()
    (target / "doc.txt").write_text("old", encoding="utf-8")

    class Failing:
        def run(self, context):
            return StageResult(status="fail", diagnostics=(Diagnostic("error", "domain.bad", "bad input"),))

    registry = Registry().register_stage("failing", Failing())
    result = PipelineRunner(registry).run(spec(stages=[{"id": "fail", "kind": "failing", "produces": []}]), tmp_path, output_dir=target)
    assert not result.success
    assert result.publication_status == "failed"
    assert result.manifest and result.manifest.is_file()
    assert (target / "doc.txt").read_text(encoding="utf-8") == "old"


def test_optional_unavailable_validator_warns_and_required_fails(tmp_path: Path):
    registry = Registry().register_stage("writer", WriteStage("doc.txt", "ok"))
    registry.register_gate("visual", None)
    optional = spec(gates=[{"id": "visual", "kind": "visual", "required": False}])
    result = PipelineRunner(registry).run(optional, tmp_path, output_dir=tmp_path / "optional")
    assert result.success
    assert any(item.code == "pipeline.gate_unavailable" and item.severity == "warning" for item in result.diagnostics)

    required = spec(gates=[{"id": "visual", "kind": "visual", "required": True}])
    result = PipelineRunner(registry).run(required, tmp_path, output_dir=tmp_path / "required")
    assert not result.success
    assert any(item.code == "pipeline.gate_unavailable" and item.severity == "error" for item in result.diagnostics)


def test_diagnostic_fingerprint_ignores_unstable_location():
    one = Diagnostic("warning", "style.emdash", "replace dash", details={"path": "a.md"}, source=SourceRef("source", line=2))
    two = Diagnostic("warning", "style.emdash", "replace dash", details={"path": "a.md"}, source=SourceRef("source", line=99))
    assert one.fingerprint == two.fingerprint


def test_confined_path_rejects_traversal_and_symlinks(tmp_path: Path):
    with pytest.raises(ValueError):
        confined_path(tmp_path, "../escape.txt")
    outside = tmp_path.parent / "outside.txt"
    outside.write_text("x", encoding="utf-8")
    link = tmp_path / "link"
    link.symlink_to(outside)
    with pytest.raises(ValueError, match="symlink"):
        confined_path(tmp_path, "link")
