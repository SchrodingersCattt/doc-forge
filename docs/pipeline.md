# Auditable pipelines

`docforge.pipeline` is the project-neutral orchestration layer.  It does not
know a journal, a manuscript layout, or a consumer repository.  A consumer
declares a `docforge.pipeline.v1` spec and registers the small stages, gates,
and validators that are specific to its workflow.

```python
from pathlib import Path

from docforge.pipeline import ArtifactRef, PipelineRunner, PipelineSpec, Registry, StageResult


class SourceStage:
    def run(self, context):
        output = context.output_path("source.txt")
        output.write_text(context.source_path("manuscript.md").read_text())
        return StageResult((ArtifactRef("source", "source", "source.txt", media_type="text/plain"),))


spec = PipelineSpec.from_dict({
    "schema": "docforge.pipeline.v1",
    "stages": [{"id": "source", "kind": "source-adapter", "produces": ["source"]}],
    "release": {"required_artifacts": ["source"], "atomic": True},
})
registry = Registry().register_stage("source-adapter", SourceStage())
result = PipelineRunner(registry).run(spec, Path("workspace"), output_dir=Path("dist"))
assert result.success
```

Stages write only through `StageContext.output_path`; input files are read with
`source_path`.  The runner rejects traversal and symlink paths, checks output
collisions and hashes every artifact.  It executes a topologically sorted DAG,
then runs gates and validators.  Required unavailable validators fail the run;
optional unavailable validators produce an explicit warning receipt.

Successful runs atomically replace the publication directory and write
`manifest.json`, `checksums.sha256`, and `report.txt` alongside outputs.  The
previous publication is retained in a sibling `.previous` directory by
default.  Failed runs never replace the previous publication and retain a
portable `.docforge-failed-<run-id>` receipt for review.

The manifest schema is `docforge.pipeline.manifest.v1`.  Paths are always
workspace-relative, and its `semantic_hash` excludes the run id so identical
inputs can be compared across runs.  Diagnostic fingerprints use stable codes,
messages, artifact identities, and details; line and column numbers are kept
for display but do not change a fingerprint.

JSON specs are available in the core install.  YAML is supported when the
optional `PyYAML` package is installed:

```bash
docforge pipeline pipeline.json --workspace . --output-dir dist
docforge pipeline pipeline.yaml --workspace . --no-publish
```

The registry is the stable extension point.  Existing Markdown, TeX, DOCX,
source-package, figure, and domain checks can be wrapped as stages or
validators without copying path, hash, manifest, or publication code into a
consumer workflow.
