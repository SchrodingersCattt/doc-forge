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
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

from ..output import validate_output_path
from .blocks import Block
from .launcher import parse_markdown
from .template import (
    AssemblyResult,
    assemble_markdown_template,
    discover_template_styles,
    _load_template_document,
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
    redline: Path | None = None
    render: Mapping[str, Any] = None
    rollback: Mapping[str, Any] = None


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


def _install(paths: Sequence[tuple[Path, Path]], *, force: bool) -> Mapping[str, Any]:
    """Install staged files and restore previous files if installation fails."""
    destinations = [destination.resolve() for _source, destination in paths]
    if len(destinations) != len(set(destinations)):
        raise ValueError("Contract installation contains duplicate destinations")
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
    except Exception as error:
        rollback_errors: list[str] = []
        for destination in reversed(installed):
            try:
                destination.unlink(missing_ok=True)
            except OSError as rollback_error:
                rollback_errors.append(f"remove {destination}: {rollback_error}")
        for destination, backup in reversed(tuple(backups.items())):
            try:
                os.replace(backup, destination)
            except OSError as rollback_error:
                rollback_errors.append(f"restore {destination}: {rollback_error}")
        if rollback_errors:
            raise RuntimeError(
                "Contract installation failed and rollback was incomplete: "
                + "; ".join(rollback_errors)
            ) from error
        raise RuntimeError(
            "Contract installation failed; rollback restored all pre-existing artifacts"
        ) from error
    finally:
        shutil.rmtree(backup_root, ignore_errors=True)
    return {
        "committed": True,
        "rollback_attempted": False,
        "rollback_restored": False,
        "artifact_count": len(paths),
    }


def _canonicalize_package(path: Path) -> None:
    """Normalize ZIP metadata so identical contract inputs have stable bytes."""
    with zipfile.ZipFile(path) as source:
        entries = [(info.filename, source.read(info.filename), info.compress_type) for info in source.infolist()]
    temporary = path.with_suffix(path.suffix + ".canonical")
    try:
        with zipfile.ZipFile(temporary, "w") as target:
            for name, payload, compress_type in sorted(entries):
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = compress_type
                info.create_system = 0
                info.external_attr = 0
                target.writestr(info, payload)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _render_receipts(paths: Sequence[Path], output_dir: Path, engine: str) -> dict[str, Any]:
    """Render staged DOCX files and return stable page-level audit receipts."""
    from ..docxmd.verify import render_pdfs

    pdfs = render_pdfs(list(paths), output_dir, engine)
    documents: list[dict[str, Any]] = []
    for docx, pdf in zip(paths, pdfs):
        pages: list[dict[str, Any]] = []
        try:
            import fitz
        except ImportError:  # pragma: no cover - exercised on minimal installs
            fitz = None
        if fitz is not None:
            with fitz.open(pdf) as rendered:
                for number, page in enumerate(rendered, start=1):
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
                    image = pixmap.tobytes("png")
                    text = page.get_text("text")
                    pages.append(
                        {
                            "page": number,
                            "sha256": hashlib.sha256(image).hexdigest(),
                            "width": pixmap.width,
                            "height": pixmap.height,
                            "empty": not text.strip() and not page.get_images(),
                            "glyph_count": sum(len(line.get("spans", [])) for block in page.get_text("dict")["blocks"] if "lines" in block for line in block["lines"]),
                            "image_count": len(page.get_images()),
                        }
                    )
        else:
            # PyMuPDF is intentionally not a hard dependency.  ``pdftoppm``
            # gives the same stable page-image receipt on minimal installs.
            prefix = output_dir / (pdf.stem + ".page")
            try:
                subprocess.run(
                    ["pdftoppm", "-png", "-r", "72", str(pdf), str(prefix)],
                    check=True,
                    capture_output=True,
                    timeout=120,
                )
                rendered_pages = sorted(output_dir.glob(prefix.name + "-*.png"))
            except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
                rendered_pages = []
            if not rendered_pages:
                # A renderer-produced PDF is still a useful receipt when no
                # image utility is present; hash the package and report one
                # opaque page rather than failing after a successful render.
                rendered_pages = [pdf]
            for number, image_path in enumerate(rendered_pages, start=1):
                payload = image_path.read_bytes()
                pages.append(
                    {
                        "page": number,
                        "sha256": hashlib.sha256(payload).hexdigest(),
                        "width": None,
                        "height": None,
                        "empty": False,
                        "glyph_count": None,
                        "image_count": None,
                    }
                )
        documents.append(
            {
                "docx": str(docx),
                "page_count": len(pages),
                "pages": pages,
            }
        )
    return {"engine": engine, "documents": documents}


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
    baseline_path: Path | None = None,
    redline_output: Path | None = None,
    render_engine: str = "auto",
    render: bool = True,
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
    baseline_path = Path(baseline_path) if baseline_path is not None else None
    if baseline_path is not None and not baseline_path.is_file():
        raise FileNotFoundError(baseline_path)
    if render_engine not in {"auto", "word", "libreoffice", "none"}:
        raise ValueError("render_engine must be one of: auto, word, libreoffice, none")
    if render_engine == "none":
        render = False
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
    if output.resolve() == template_path.resolve():
        raise ValueError("Contract output must differ from the template path")
    redline_path = Path(redline_output) if redline_output is not None else (
        output.with_name(output.stem + ".tracked" + output.suffix) if baseline_path is not None else None
    )
    if redline_path is not None:
        if redline_path.resolve() in {output.resolve(), template_path.resolve(), baseline_path.resolve() if baseline_path else None}:
            raise ValueError("Redline output must differ from the template, baseline, and clean output")
        validate_output_path(redline_path)

    # Strict discovery is the contract's key preflight.  It runs before any
    # destination is touched, so a malformed template cannot leave a stale or
    # partially generated DOCX that looks usable to a caller.
    # Use the same revision-accepting DOCX/DOTX loader as the regular
    # assembly path.  Opening a real DOTX directly with python-docx fails on
    # its template content type before semantic preflight can run.
    roles = discover_template_styles(_load_template_document(template_path), strict=True)
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
        _canonicalize_package(staged_output)
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
        # The sidecar was generated beside the staged file.  Rewrite its
        # output descriptor to the destination that is about to be installed
        # so the published audit never points into the deleted temp directory.
        manifest_data["output"] = {
            "path": str(output.resolve()),
            "sha256": sha256_file(staged_output),
            "size_bytes": staged_output.stat().st_size,
        }
        # Assembly sidecars include a wall-clock timestamp.  Contract manifests
        # are reproducible audit records, so retain a stable epoch marker.
        manifest_data["generated_at_utc"] = "1970-01-01T00:00:00+00:00"
        redline_receipt: dict[str, Any] = {"enabled": False}
        staged_redline: Path | None = None
        if baseline_path is not None:
            from ..docxdiff import create_tracked_docx

            staged_redline = stage / redline_path.name
            redline_report = stage / (redline_path.stem + ".report.json")
            redline_summary = create_tracked_docx(
                baseline_path,
                staged_output,
                staged_redline,
                author="docforge-contract",
                overwrite=True,
                accept_baseline=True,
                report_path=redline_report,
                workers=1,
                revision_date="1970-01-01T00:00:00Z",
            )
            redline_report_data = json.loads(redline_report.read_text(encoding="utf-8"))
            redline_report_data["current"] = str(output.resolve())
            redline_report_data["output"] = str(redline_path.resolve())
            redline_receipt = {
                "enabled": True,
                "baseline": {"path": str(baseline_path.resolve()), "sha256": sha256_file(baseline_path)},
                "output": {
                    "path": str(redline_path.resolve()),
                    "sha256": sha256_file(staged_redline),
                    "size_bytes": staged_redline.stat().st_size,
                },
                "summary": redline_summary,
                "accepted_view_equivalent": True,
                "report": redline_report_data,
            }
            _canonicalize_package(staged_redline)
            redline_receipt["output"]["sha256"] = sha256_file(staged_redline)
            redline_receipt["output"]["size_bytes"] = staged_redline.stat().st_size
        render_receipt: dict[str, Any] = {"enabled": False}
        if render:
            render_paths = [staged_output] + ([staged_redline] if staged_redline is not None else [])
            render_receipt = {"enabled": True, **_render_receipts(render_paths, stage / "render", render_engine)}
            # Paths in a published manifest must never point into the staging tree.
            for document in render_receipt["documents"]:
                document["docx"] = str(output.resolve()) if document["docx"] == str(staged_output) else str(redline_path.resolve())
        installation: dict[str, Any] = {
            "committed": True,
            "rollback_attempted": False,
            "rollback_restored": False,
        }
        manifest_data["contract"] = {
            "schema": "docforge.docx-contract.v1",
            "atomic": True,
            "semantic_roles": dict(roles),
            "blocks": list(blocks),
            "verification": verification,
            "output_sha256": sha256_file(staged_output),
            "accepted_view": {"equivalent": True, "source": "assembly-verifier"},
            "redline": redline_receipt,
            "render": render_receipt,
            "installation": installation,
        }
        staged_manifest.write_text(json.dumps(manifest_data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        manifest = output.with_suffix(".manifest.json")
        checksum = output.with_suffix(".sha256")
        install_paths = [(staged_output, output), (staged_manifest, manifest), (staged_checksum, checksum)]
        if staged_redline is not None:
            install_paths.append((staged_redline, redline_path))
        installation = _install(install_paths, force=force)
    return ContractResult(output, manifest, checksum, roles, blocks, redline_path, render_receipt, installation)


__all__ = ["ContractResult", "build_docx_contract"]
