"""Transactional, project-neutral one-pass DOCX assembly.

The regular template assembly API is intentionally composable.  Consumers that
want one command which either publishes a complete, audited document or leaves
the destination untouched can use :func:`build_docx_contract` instead.  The
helper performs semantic-style preflight, builds in a sibling temporary
directory, validates the result, writes all audit sidecars, and commits the
three files together.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

from docx import Document

from ..output import validate_output_path
from .blocks import Block
from .launcher import parse_markdown
from .template import (
    AssemblyResult,
    assemble_markdown_template,
    discover_template_styles,
    sha256_file,
    verify_template_output,
    write_assembly_sidecars,
)


@dataclass(frozen=True)
class ContractResult:
    """Files and audit information published by a contract build."""

    output: Path
    manifest: Path
    checksum: Path
    roles: Mapping[str, str]
    blocks: tuple[Mapping[str, Any], ...]


def _role(block: Block) -> str:
    if block.kind == "heading":
        return f"heading_{max(1, min(block.level, 3))}"
    return {
        "paragraph": "body",
        "quote": "body",
        "ordered": "body",
        "bullet": "body",
        "code": "body",
        "equation": "body",
        "table": "body",
        "separator": "body",
        "image": "caption",
        "table_caption": "caption",
        "reference": "reference",
    }.get(block.kind, "body")


def _block_audit(inputs: Sequence[Path], *, keep_comments: bool) -> tuple[dict[str, Any], ...]:
    records: list[dict[str, Any]] = []
    for source_index, source in enumerate(inputs):
        blocks = parse_markdown(source, strip_comments=not keep_comments)
        for index, block in enumerate(blocks):
            records.append(
                {
                    "source_index": source_index,
                    "source": source.name,
                    "index": index,
                    "kind": block.kind,
                    "role": _role(block),
                    "decision": "render",
                    "text_sha256": hashlib.sha256(block.text.encode("utf-8")).hexdigest(),
                }
            )
    return tuple(records)


def _install(paths: Sequence[tuple[Path, Path]], *, force: bool) -> None:
    """Install staged files and restore previous files if installation fails."""
    if not force and any(destination.exists() for _source, destination in paths):
        existing = next(destination for _source, destination in paths if destination.exists())
        raise FileExistsError(f"Output exists; pass --force to overwrite: {existing}")
    backup_root = Path(tempfile.mkdtemp(prefix=".docforge-contract-backup-", dir=paths[0][1].parent))
    backups: dict[Path, Path] = {}
    installed: list[Path] = []
    try:
        for _source, destination in paths:
            if destination.exists():
                backup = backup_root / str(len(backups))
                backup.parent.mkdir(parents=True, exist_ok=True)
                os.replace(destination, backup)
                backups[destination] = backup
        for source, destination in paths:
            destination.parent.mkdir(parents=True, exist_ok=True)
            os.replace(source, destination)
            installed.append(destination)
    except Exception:
        for destination in installed:
            destination.unlink(missing_ok=True)
        for destination, backup in backups.items():
            os.replace(backup, destination)
        raise
    finally:
        shutil.rmtree(backup_root, ignore_errors=True)


def build_docx_contract(
    inputs: Sequence[Path],
    *,
    template_path: Path,
    output: Path,
    metadata_path: Path,
    bibliography_path: Path | None = None,
    citation_base_path: Path | None = None,
    force: bool = False,
    command: Sequence[str] | None = None,
    **options: Any,
) -> ContractResult:
    """Build and publish one audited DOCX as an all-or-nothing operation.

    ``options`` are the rendering options accepted by
    :func:`assemble_markdown_template`.  The template must define every
    semantic role; this preflight is deliberately stricter than the historical
    ``md2docx`` compatibility path, which permits a minimal Normal-only
    template.
    """
    source_paths = tuple(Path(path) for path in inputs)
    template_path, output, metadata_path = Path(template_path), Path(output), Path(metadata_path)
    if not source_paths:
        raise ValueError("Provide at least one Markdown input")
    missing = [path for path in source_paths if not path.is_file()]
    if missing:
        raise FileNotFoundError(missing[0])
    validate_output_path(output)
    if not template_path.is_file():
        raise FileNotFoundError(template_path)
    if not metadata_path.is_file():
        raise FileNotFoundError(metadata_path)

    # Strict discovery is the contract's key preflight.  It runs before any
    # destination is touched, so a malformed template cannot leave a stale or
    # partially generated DOCX that looks usable to a caller.
    roles = discover_template_styles(Document(template_path), strict=True)
    blocks = _block_audit(source_paths, keep_comments=bool(options.get("keep_comments", False)))
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".docforge-contract-", dir=output.parent) as temporary:
        stage = Path(temporary)
        staged_output = stage / output.name
        assembly_options = dict(options)
        assembly_options.pop("keep_comments", None)
        result: AssemblyResult = assemble_markdown_template(
            source_paths,
            template_path=template_path,
            output=staged_output,
            metadata_path=metadata_path,
            bibliography_path=bibliography_path,
            citation_base_path=citation_base_path,
            keep_comments=bool(options.get("keep_comments", False)),
            force=True,
            **assembly_options,
        )
        # Run the public verifier once more at the contract boundary.  This
        # makes the one-pass helper safe even if assembly internals evolve.
        verification = verify_template_output(result.output)
        staged_manifest, staged_checksum = write_assembly_sidecars(
            result,
            inputs=source_paths,
            template_path=template_path,
            metadata_path=metadata_path,
            bibliography_path=bibliography_path,
            citation_base_path=citation_base_path,
            command=command or ["docforge", "contract"],
        )
        manifest_data = json.loads(staged_manifest.read_text(encoding="utf-8"))
        manifest_data["contract"] = {
            "schema": "docforge.docx-contract.v1",
            "atomic": True,
            "semantic_roles": dict(roles),
            "blocks": list(blocks),
            "verification": verification,
            "output_sha256": sha256_file(staged_output),
        }
        staged_manifest.write_text(json.dumps(manifest_data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        manifest = output.with_suffix(".manifest.json")
        checksum = output.with_suffix(".sha256")
        _install(
            [(staged_output, output), (staged_manifest, manifest), (staged_checksum, checksum)],
            force=force,
        )
    return ContractResult(output, manifest, checksum, roles, blocks)


__all__ = ["ContractResult", "build_docx_contract"]
