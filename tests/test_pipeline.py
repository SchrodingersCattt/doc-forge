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


def test_exception_message_redaction_preserves_closing_quotes():
    """Absolute paths inside quoted exception text must not swallow the quote."""
    diagnostic = Diagnostic(
        "error",
        "pipeline.stage_exception",
        "stage failed",
        details={
            "exception_message": "No such file: '/workspace/secret/input.md'",
            "windows_message": r"No such file: 'C:\Users\me\secret\input.md'",
            "double_quoted_posix": 'No such file: "/workspace/secret/input.md"',
            "double_quoted_windows": r'No such file: "C:\Users\me\secret\input.md"',
        },
    )
    details = diagnostic.to_dict()["details"]
    assert details["exception_message"] == "No such file: '<absolute-path>'"
    assert details["windows_message"] == "No such file: '<absolute-path>'"
    assert details["double_quoted_posix"] == 'No such file: "<absolute-path>"'
    assert details["double_quoted_windows"] == 'No such file: "<absolute-path>"'


def test_confined_path_rejects_traversal_and_symlinks(tmp_path: Path):
    with pytest.raises(ValueError):
        confined_path(tmp_path, "../escape.txt")
    outside = tmp_path.parent / "outside.txt"
    outside.write_text("x", encoding="utf-8")
    link = tmp_path / "link"
    link.symlink_to(outside)
    with pytest.raises(ValueError, match="symlink"):
        confined_path(tmp_path, "link")


def test_publication_rejects_workspace_ancestor_and_unmarked_target(tmp_path: Path):
    registry = Registry().register_stage("writer", WriteStage("doc.txt", "new"))
    workspace = tmp_path / "project" / "workspace"
    workspace.mkdir(parents=True)
    with pytest.raises(ValueError, match="contain the workspace"):
        PipelineRunner(registry).run(spec(), workspace, output_dir=workspace.parent)
    assert workspace.is_dir()

    target = tmp_path / "notes"
    target.mkdir()
    (target / "keep.txt").write_text("keep", encoding="utf-8")
    with pytest.raises(ValueError, match="unmarked"):
        PipelineRunner(registry).run(spec(), workspace, output_dir=target)
    assert (target / "keep.txt").read_text(encoding="utf-8") == "keep"


def test_publication_replaces_only_marked_previous_output(tmp_path: Path):
    registry = Registry().register_stage("writer", WriteStage("doc.txt", "new"))
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    target = tmp_path / "published"
    first = PipelineRunner(registry).run(spec(), workspace, output_dir=target, run_id="first")
    assert first.success
    (target / "doc.txt").write_text("updated", encoding="utf-8")
    second = PipelineRunner(registry).run(spec(), workspace, output_dir=target, run_id="second")
    assert second.success
    assert (target / "manifest.json").is_file()


def test_preflight_rejects_undeclared_and_non_dependency_inputs():
    registry = Registry().register_stage("writer", WriteStage("doc.txt", "new"))
    undeclared = spec(stages=[{"id": "consumer", "kind": "writer", "consumes": ["missing"], "produces": ["document"]}])
    with pytest.raises(PipelinePreflightError, match="undeclared artifact"):
        PipelineRunner(registry).preflight(undeclared)
    detached = spec(stages=[
        {"id": "producer", "kind": "writer", "produces": ["input"]},
        {"id": "consumer", "kind": "writer", "consumes": ["input"], "produces": ["document"]},
    ])
    with pytest.raises(PipelinePreflightError, match="without depending"):
        PipelineRunner(registry).preflight(detached)
    bad_source = PipelineSpec.from_dict({
        "schema": "docforge.pipeline.v1",
        "documents": [{"id": "main", "source": {"path": "a/../../x.md"}}],
    })
    with pytest.raises(PipelinePreflightError, match="source path"):
        PipelineRunner().preflight(bad_source)


def test_runtime_rejects_undeclared_and_duplicate_artifacts(tmp_path: Path):
    class BadStage:
        def run(self, context):
            output = context.output_path("bad.txt")
            output.write_text("bad", encoding="utf-8")
            return StageResult((ArtifactRef("wrong", "wrong", "bad.txt"),))

    registry = Registry().register_stage("bad", BadStage())
    result = PipelineRunner(registry).run(spec(stages=[{"id": "bad", "kind": "bad", "produces": ["declared"]}], release={"required_artifacts": []}), tmp_path, output_dir=tmp_path / "out")
    assert not result.success
    assert any(item.code == "pipeline.artifact_output_invalid" for item in result.diagnostics)

    class DuplicateStage:
        def run(self, context):
            for name in ("a.txt", "b.txt"):
                context.output_path(name).write_text(name, encoding="utf-8")
            return StageResult((ArtifactRef("same", "a", "a.txt"), ArtifactRef("same", "b", "b.txt")))

    registry = Registry().register_stage("duplicate", DuplicateStage())
    result = PipelineRunner(registry).run(spec(stages=[{"id": "dup", "kind": "duplicate", "produces": ["a", "b"]}], release={"required_artifacts": []}), tmp_path, output_dir=tmp_path / "duplicate")
    assert not result.success
    assert any("duplicate artifact id" in item.details.get("exception_message", "") for item in result.diagnostics)

    class ReservedStage:
        def run(self, context):
            context.output_path("sub/../manifest.json").write_text("bad", encoding="utf-8")
            return StageResult((ArtifactRef("manifest", "declared", "sub/../manifest.json"),))

    registry = Registry().register_stage("reserved", ReservedStage())
    result = PipelineRunner(registry).run(spec(stages=[{"id": "reserved", "kind": "reserved", "produces": ["declared"]}], release={"required_artifacts": []}), tmp_path, output_dir=tmp_path / "reserved")
    assert not result.success
    assert any("reserved audit file" in item.details.get("exception_message", "") for item in result.diagnostics)


def test_failure_manifest_paths_and_fingerprints_are_stable(tmp_path: Path):
    class MissingStage:
        def run(self, context):
            return StageResult((ArtifactRef("missing", "document", "nope.txt"),))

    pipeline = PipelineRunner(Registry().register_stage("missing", MissingStage()))
    pipeline_spec = spec(stages=[{"id": "missing", "kind": "missing", "produces": ["document"]}])
    fingerprints = []
    for root in (tmp_path / "one", tmp_path / "two"):
        root.mkdir()
        result = pipeline.run(pipeline_spec, root, output_dir=root.parent / f"out-{root.name}", run_id="stable")
        payload = json.loads(result.manifest.read_text(encoding="utf-8"))
        encoded = json.dumps(payload, sort_keys=True)
        assert str(root) not in encoded
        fingerprints.append(next(item.fingerprint for item in result.diagnostics if item.code == "pipeline.artifact_output_invalid"))
    assert fingerprints[0] == fingerprints[1]


def test_cache_key_uses_consumed_input_hashes(tmp_path: Path):
    class Producer:
        def __init__(self, value):
            self.value = value
        def run(self, context):
            context.output_path("input.txt").write_text(self.value, encoding="utf-8")
            return StageResult((ArtifactRef("input", "input", "input.txt"),))

    class Consumer:
        def run(self, context):
            context.output_path("output.txt").write_text("ok", encoding="utf-8")
            return StageResult((ArtifactRef("output", "document", "output.txt"),))

    pipeline_spec = spec(stages=[
        {"id": "produce", "kind": "produce", "produces": ["input"]},
        {"id": "consume", "kind": "consume", "depends_on": ["produce"], "consumes": ["input"], "produces": ["document"]},
    ])
    registry = Registry().register_stage("produce", Producer("one")).register_stage("consume", Consumer())
    (tmp_path / "one").mkdir()
    (tmp_path / "two").mkdir()
    first = PipelineRunner(registry).run(pipeline_spec, tmp_path / "one", output_dir=tmp_path / "out-one")
    registry.stages["produce"] = Producer("two")
    second = PipelineRunner(registry).run(pipeline_spec, tmp_path / "two", output_dir=tmp_path / "out-two")
    assert first.stage_results["consume"].cache_key != second.stage_results["consume"].cache_key


def test_cli_builtin_copy_adapter(tmp_path: Path):
    from docforge.cli import main

    workspace = tmp_path / "workspace"
    workspace.mkdir()
    (workspace / "input.txt").write_text("hello", encoding="utf-8")
    config = {
        "schema": "docforge.pipeline.v1",
        "stages": [{"id": "copy", "kind": "copy-artifact", "produces": ["document"], "parameters": {"source": "input.txt", "artifact_id": "document", "role": "document", "output": "copy.txt", "media_type": "text/plain"}}],
        "release": {"required_artifacts": ["document"]},
    }
    spec_path = workspace / "pipeline.json"
    spec_path.write_text(json.dumps(config), encoding="utf-8")
    assert main(["pipeline", str(spec_path), "--workspace", str(workspace), "--output-dir", str(tmp_path / "published")]) == 0
    assert (tmp_path / "published" / "copy.txt").read_text(encoding="utf-8") == "hello"


def test_callable_stage_and_gate_registrations_are_adapted(tmp_path: Path):
    def write(context):
        context.output_path("callable.txt").write_text("ok", encoding="utf-8")
        return {"artifacts": [ArtifactRef("document", "document", "callable.txt")]}

    registry = Registry().register_stage("callable", write).register_gate("quality", lambda context: {"status": "pass"})
    pipeline_spec = spec(stages=[{"id": "callable", "kind": "callable", "produces": ["document"]}], gates=[{"id": "quality", "kind": "quality"}])
    result = PipelineRunner(registry).run(pipeline_spec, tmp_path, output_dir=tmp_path / "callable")
    assert result.success
