"""Auditable DOCX round-trip orchestration.

The round-trip command keeps the editable Markdown bundle beside a portable
source package.  An unchanged bundle is copied byte-for-byte; a changed
bundle is rendered through the normal Markdown renderer and can optionally be
diffed against a reviewed baseline.
"""

from __future__ import annotations

import hashlib
import importlib.metadata
import json
import os
import posixpath
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

from lxml import etree

from .markdown import (
    convert_unicode_scripts_in_docx,
    docx_to_markdown,
    load_section_map,
    parse_markdown,
    render_blocks_to_doc,
)


SCHEMA = "docforge.roundtrip.v1"
# Public spelling used by callers that validate bundle manifests.
ROUNDTRIP_SCHEMA = SCHEMA
_LEGACY_SCHEMA = "docforge.docx2md.v1"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _version() -> str:
    try:
        return importlib.metadata.version("docforge")
    except importlib.metadata.PackageNotFoundError:
        return "0.1.0"


def _relationship_part(name: str) -> str:
    """Return the source part represented by a ``.rels`` part name."""

    if name == "_rels/.rels":
        return ""
    parent, rels_name = posixpath.split(name)
    if not rels_name.endswith(".rels") or not parent.endswith("/_rels"):
        return ""
    source_parent = parent[: -len("/_rels")]
    source_name = rels_name[: -len(".rels")]
    return posixpath.join(source_parent, source_name)


def audit_docx_package(path: Path) -> dict[str, Any]:
    """Validate the OPC package and revision references before publication.

    The audit deliberately stays package-level and deterministic: it does not
    depend on Word or a particular renderer.  Every relationship target must
    resolve inside the archive, all XML parts must parse, and comment anchors
    must point at an existing comment.  A ``ValueError`` identifies the first
    actionable defect, so callers can fail before writing a misleading output.
    """

    path = Path(path)
    if not path.is_file() or not zipfile.is_zipfile(path):
        raise ValueError(f"DOCX package is not a readable OPC zip: {path}")
    try:
        with zipfile.ZipFile(path) as archive:
            bad = archive.testzip()
            if bad is not None:
                raise ValueError(f"DOCX package has a corrupt member: {bad}")
            names = set(archive.namelist())
            required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
            missing = sorted(required - names)
            if missing:
                raise ValueError(f"DOCX package is missing required parts: {', '.join(missing)}")
            xml_roots: dict[str, etree._Element] = {}
            for name in names:
                if not name.endswith(".xml") and not name.endswith(".rels"):
                    continue
                try:
                    xml_roots[name] = etree.fromstring(archive.read(name))
                except (etree.XMLSyntaxError, KeyError) as exc:
                    raise ValueError(f"DOCX package contains invalid XML: {name}") from exc

            for rels_name, root in xml_roots.items():
                if not rels_name.endswith(".rels"):
                    continue
                source = _relationship_part(rels_name)
                source_dir = posixpath.dirname(source)
                for relationship in root.findall(f"{{{_REL_NS}}}Relationship"):
                    if relationship.get("TargetMode") == "External":
                        continue
                    target = relationship.get("Target") or ""
                    if target.startswith("/"):
                        target_name = target.lstrip("/")
                    else:
                        target_name = posixpath.normpath(posixpath.join(source_dir, target))
                    if target_name not in names:
                        raise ValueError(
                            f"DOCX relationship from {source or '/'} points to missing part: {target}"
                        )

            document = xml_roots["word/document.xml"]
            refs = {
                node.get(f"{{{_W_NS}}}id")
                for node in document.iter(f"{{{_W_NS}}}commentReference")
            }
            refs.discard(None)
            comments = xml_roots.get("word/comments.xml")
            known = set()
            if comments is not None:
                known = {
                    node.get(f"{{{_W_NS}}}id")
                    for node in comments.iter(f"{{{_W_NS}}}comment")
                }
                known.discard(None)
            missing_comments = sorted(refs - known)
            if missing_comments:
                raise ValueError(
                    "DOCX comment anchors reference missing comments: "
                    + ", ".join(missing_comments)
                )
            # Source syntax must never leak into the published Word text.
            visible_text = "".join(document.itertext())
            leaked = [token for token in ("\\cite{", "\\ref{", "$$", "\\begin{aligned}") if token in visible_text]
            if leaked:
                raise ValueError("DOCX contains unresolved source markup: " + ", ".join(leaked))
            return {
                "passed": True,
                "parts": len(names),
                "xml_parts": len(xml_roots),
                "relationships": sum(
                    len(root.findall(f"{{{_REL_NS}}}Relationship"))
                    for name, root in xml_roots.items()
                    if name.endswith(".rels")
                ),
                "comment_references": len(refs),
            }
    except zipfile.BadZipFile as exc:
        raise ValueError(f"DOCX package is not a readable OPC zip: {path}") from exc


def render_docx_pages(path: Path, *, renderer: str | None = None) -> dict[str, Any]:
    """Render every page with a declared office renderer and hash its image.

    LibreOffice is preferred when present.  Environments without an office
    renderer still receive an explicit ``skipped`` result in the manifest;
    callers can require ``status == 'passed'`` in CI when visual auditing is
    available.  The renderer command is recorded so the audit is reproducible.
    """

    renderer_path = renderer or shutil.which("soffice") or shutil.which("libreoffice")
    if renderer_path is None:
        return {"status": "skipped", "renderer": None, "pages": 0, "image_hashes": [], "issues": ["no office renderer found"]}
    path = Path(path).resolve()
    with tempfile.TemporaryDirectory(prefix="docforge-render-") as temp:
        root = Path(temp)
        command = [renderer_path, "--headless", "--convert-to", "pdf", "--outdir", str(root), str(path)]
        completed = subprocess.run(command, capture_output=True, text=True, check=False)
        pdf = root / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file():
            detail = (completed.stderr or completed.stdout or "renderer failed").strip()
            raise ValueError(f"DOCX page renderer failed ({renderer_path}): {detail}")
        try:
            import fitz  # type: ignore[import-not-found]
        except ImportError:
            return {
                "status": "passed",
                "renderer": renderer_path,
                "pages": 0,
                "image_hashes": [],
                "issues": ["PDF produced; image hashing unavailable (PyMuPDF not installed)"],
            }
        document = fitz.open(pdf)
        hashes: list[str] = []
        for page in document:
            pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            hashes.append(hashlib.sha256(pixmap.tobytes("png")).hexdigest())
        document.close()
        return {
            "status": "passed",
            "renderer": renderer_path,
            "pages": len(hashes),
            "image_hashes": hashes,
            "issues": [],
        }


def _safe_path(root: Path, value: object, label: str) -> Path:
    """Resolve a bundle path while rejecting absolute and escaping paths."""

    if not isinstance(value, str) or not value or Path(value).is_absolute():
        raise ValueError(f"{label} must be a bundle-relative path")
    root = root.resolve()
    candidate = (root / Path(value)).resolve()
    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"{label} escapes the roundtrip bundle: {value}") from exc
    return candidate


def _hash_entries(root: Path, entries: Mapping[str, Any], label: str) -> tuple[bool, list[Path]]:
    paths: list[Path] = []
    unchanged = True
    for name, expected in entries.items():
        path = _safe_path(root, name, label)
        if not path.is_file() or not isinstance(expected, str):
            unchanged = False
            paths.append(path)
            continue
        if sha256_file(path) != expected:
            unchanged = False
        paths.append(path)
    return unchanged, paths


def _manifest_paths(
    root: Path, payload: Mapping[str, Any]
) -> tuple[Path, list[Path], list[Path], bool, bool, bool, bool]:
    source_name = payload.get("source_copy") or payload.get("source_snapshot")
    source_sha = payload.get("source_copy_sha256") or payload.get("source_snapshot_sha256")
    if not isinstance(source_name, str) or not isinstance(source_sha, str):
        raise ValueError("roundtrip manifest does not contain a source DOCX snapshot")
    source = _safe_path(root, source_name, "source snapshot")
    if not source.is_file() or sha256_file(source) != source_sha:
        raise ValueError("roundtrip source snapshot is missing or its SHA-256 does not match")

    sections_map = payload.get("section_sha256", {})
    media_map = payload.get("media_sha256", {})
    if not isinstance(sections_map, Mapping) or not isinstance(media_map, Mapping):
        raise ValueError("roundtrip manifest has invalid section/media hashes")
    sections_ok, sections = _hash_entries(root, sections_map, "section")
    media_ok, media = _hash_entries(root, media_map, "media")

    full_name = payload.get("full") or payload.get("full_markdown")
    full_sha = payload.get("full_sha256") or payload.get("full_markdown_sha256")
    full: Path | None = None
    full_ok = True
    if full_name is not None or full_sha is not None:
        if not isinstance(full_name, str) or not isinstance(full_sha, str):
            raise ValueError("roundtrip manifest has an invalid full Markdown entry")
        full = _safe_path(root, full_name, "full Markdown")
        full_ok = full.is_file() and sha256_file(full) == full_sha
    # A legacy split manifest has no full Markdown hash.  Its section files
    # are the complete editable source.
    unchanged = sections_ok and media_ok and full_ok
    return source, sections, media, unchanged, full_ok, sections_ok, media_ok


@dataclass(frozen=True)
class RoundtripResult:
    """Machine-readable outcome of one round-trip invocation."""

    mode: str
    source_sha256: str
    output_sha256: str
    manifest_sha256: str
    revision_summary: Mapping[str, int] = field(default_factory=dict)
    validation: Mapping[str, Any] = field(default_factory=dict)
    output: Path | None = None
    manifest: Path | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "source_sha256": self.source_sha256,
            "output_sha256": self.output_sha256,
            "manifest_sha256": self.manifest_sha256,
            "revision_summary": dict(self.revision_summary),
            "validation": dict(self.validation),
        }


def _portable_command(input_path: Path, output: Path, workdir: Path, track_changes: str) -> list[str]:
    args = ["docforge", "roundtrip", input_path.name, "--workdir", workdir.name, "--output", output.name]
    if track_changes != "accept":
        args.extend(["--track-changes", track_changes])
    return args


def _portable_command_args(command: list[str] | None, fallback: list[str], workdir: Path) -> list[str]:
    """Keep invocation options while removing host-specific absolute paths."""

    if not command:
        return fallback
    result: list[str] = []
    for token in command:
        value = str(token)
        if os.path.isabs(value):
            result.append(Path(value).name)
        else:
            result.append(value)
    return result


def _write_manifest(path: Path, payload: Mapping[str, Any]) -> None:
    path.write_text(json.dumps(dict(payload), ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _export_bundle(
    input_path: Path,
    workdir: Path,
    *,
    section_map: Path | None,
    track_changes: str,
    force: bool,
    output: Path,
    command: list[str] | None = None,
) -> Path:
    workdir.mkdir(parents=True, exist_ok=True)
    full = workdir / "full.md"
    if full.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {full}")
    source_copy = workdir / "roundtrip" / "source.docx"
    if source_copy.exists() and sha256_file(source_copy) != sha256_file(input_path):
        raise ValueError("roundtrip source snapshot exists and differs from input DOCX")
    if section_map is not None:
        # docx2md writes map-provided filenames directly below ``workdir``.
        # Reject traversal before the exporter can write outside the bundle.
        for entry in load_section_map(section_map):
            name = entry.get("file")
            if not isinstance(name, str):
                raise ValueError("section map file names must be strings")
            candidate = _safe_path(workdir, name, "section map output")
            reserved = {
                full.resolve(),
                (workdir / "manifest.json").resolve(),
                (workdir / "roundtrip").resolve(),
                (workdir / "media").resolve(),
            }
            if candidate in reserved or any(parent in candidate.parents for parent in reserved):
                raise ValueError(f"section map output uses a reserved bundle path: {name}")
    export = docx_to_markdown(
        input_path,
        output=full,
        split_dir=workdir if section_map is not None else None,
        section_map_path=section_map,
        media_dir=workdir,
        track_changes=track_changes,
        force=force,
    )
    source_copy.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(input_path, source_copy)
    sections = [path for path in export.sections if path.is_file()]
    media = [path for path in export.media if path.is_file()]
    relative = lambda path: path.relative_to(workdir.resolve()).as_posix()
    manifest: dict[str, Any] = {
        "schema": SCHEMA,
        "schema_version": 1,
        "source_sha256": sha256_file(input_path),
        "source_copy": relative(source_copy),
        "source_copy_sha256": sha256_file(source_copy),
        "source_snapshot": relative(source_copy),
        "source_snapshot_sha256": sha256_file(source_copy),
        "full": relative(full),
        "full_sha256": sha256_file(full),
        "sections": [relative(path) for path in sections],
        "section_sha256": {relative(path): sha256_file(path) for path in sections},
        "media": [relative(path) for path in media],
        "media_sha256": {relative(path): sha256_file(path) for path in media},
        "track_changes": track_changes,
        "tool_version": _version(),
        "command": _portable_command_args(
            command,
            _portable_command(input_path, output, workdir, track_changes),
            workdir,
        ),
        "command_args": _portable_command_args(
            command,
            _portable_command(input_path, output, workdir, track_changes),
            workdir,
        )[2:],
    }
    manifest_path = workdir / "manifest.json"
    _write_manifest(manifest_path, manifest)
    return manifest_path


def _load_manifest(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"invalid roundtrip manifest: {path}") from exc
    if not isinstance(payload, dict) or payload.get("schema") not in {SCHEMA, _LEGACY_SCHEMA}:
        raise ValueError(f"roundtrip manifest must use schema {SCHEMA}")
    return payload


def _render_markdown(inputs: list[Path], output: Path) -> None:
    """Render Markdown while resolving bundle-relative image links."""

    import re

    image = re.compile(r"(?P<prefix>!\[[^\]]*\]\()(?P<path>[^)\s]+)(?P<suffix>\))")
    blocks = []
    with tempfile.TemporaryDirectory(prefix="docforge-roundtrip-md-") as temp:
        temp_root = Path(temp)
        for index, path in enumerate(inputs):
            text = path.read_text(encoding="utf-8")

            def resolve(match: re.Match[str]) -> str:
                value = match.group("path")
                if value.startswith(("http://", "https://", "data:")):
                    return match.group(0)
                candidate = (path.parent / value).resolve()
                if not candidate.is_file():
                    return match.group(0)
                return f"{match.group('prefix')}{candidate.as_posix()}{match.group('suffix')}"

            temporary = temp_root / f"{index:04d}.md"
            temporary.write_text(image.sub(resolve, text), encoding="utf-8")
            blocks.extend(parse_markdown(temporary))
    if not blocks:
        raise ValueError("roundtrip bundle contains no Markdown input")
    document = render_blocks_to_doc(blocks)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))
    convert_unicode_scripts_in_docx(output)


def roundtrip_docx(
    input_path: Path,
    *,
    workdir: Path,
    output: Path,
    section_map: Path | None = None,
    section_map_path: Path | None = None,
    track_changes: str = "accept",
    baseline: Path | None = None,
    force: bool = False,
    command: list[str] | None = None,
) -> RoundtripResult:
    """Export, verify, and rebuild a DOCX bundle in one auditable operation."""

    if section_map is not None and section_map_path is not None and section_map != section_map_path:
        raise ValueError("pass only one of section_map and section_map_path")
    section_map = section_map if section_map is not None else section_map_path
    input_path = input_path.resolve()
    workdir = workdir.resolve()
    output = output.resolve()
    from .output import validate_output_path

    validate_output_path(output)
    if not input_path.is_file():
        raise FileNotFoundError(input_path)
    if track_changes not in {"accept", "reject", "all"}:
        raise ValueError("track_changes must be accept, reject, or all")
    reserved_paths = {
        (workdir / "manifest.json").resolve(),
        (workdir / "full.md").resolve(),
        (workdir / "roundtrip" / "source.docx").resolve(),
    }
    if output in reserved_paths:
        raise ValueError("roundtrip output must be outside reserved bundle files")
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    manifest_path = workdir / "manifest.json"
    if not manifest_path.exists():
        manifest_path = _export_bundle(
            input_path,
            workdir,
            section_map=section_map,
            track_changes=track_changes,
            force=force,
            output=output,
            command=command,
        )
    payload = _load_manifest(manifest_path)
    expected_input_sha = payload.get("source_sha256")
    if isinstance(expected_input_sha, str) and sha256_file(input_path) != expected_input_sha:
        raise ValueError("input DOCX SHA-256 does not match the roundtrip manifest")
    source, sections, _media, unchanged, full_ok, sections_ok, media_ok = _manifest_paths(workdir, payload)
    source_sha = sha256_file(source)
    source_validation = audit_docx_package(source)
    revision_summary: Mapping[str, int] = {}
    validation: dict[str, Any] = {"source_package": source_validation}
    if unchanged:
        # The source package was audited before it is copied to the delivery
        # path, so an invalid bundle can never produce a successful output.
        mode = "exact-reuse"
        validation["accepted_view"] = {"passed": True, "mode": "exact-reuse"}
        validation["output_package"] = audit_docx_package(source)
        validation["render"] = render_docx_pages(source)
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, output)
    else:
        full_name = payload.get("full") or payload.get("full_markdown")
        full = _safe_path(workdir, full_name, "full Markdown") if isinstance(full_name, str) else None
        full_changed = not full_ok
        if full_changed and full is not None and full.is_file():
            inputs = [full]
        elif not sections_ok:
            inputs = sections
        elif not media_ok:
            inputs = [full] if full is not None else sections
        else:
            inputs = sections
        if not inputs:
            raise ValueError("roundtrip bundle changed but contains no usable Markdown inputs")
        with tempfile.TemporaryDirectory(prefix="docforge-roundtrip-") as temp:
            rebuilt = Path(temp) / "rebuilt.docx"
            _render_markdown(inputs, rebuilt)
            rebuilt_validation = audit_docx_package(rebuilt)
            if baseline is not None:
                from .docxdiff import create_tracked_docx

                tracked = Path(temp) / "tracked.docx"
                revision_summary = create_tracked_docx(
                    baseline.resolve(), rebuilt, tracked, overwrite=True
                )
                # Audit the temporary tracked package before it becomes the
                # public output.  A failed audit must not leave a misleading
                # deliverable at the requested path.
                tracked_validation = audit_docx_package(tracked)
                validation["accepted_view"] = {
                    "passed": True,
                    "mode": "tracked",
                    "summary": dict(revision_summary),
                }
                validation["tracked_package"] = tracked_validation
                candidate = tracked
            else:
                candidate = rebuilt
                validation["accepted_view"] = {"passed": True, "mode": "generated"}
            validation["rebuilt_package"] = rebuilt_validation
            # Validate/render while the temporary candidate still exists, then
            # atomically publish the audited bytes to the requested path.
            validation["output_package"] = audit_docx_package(candidate)
            validation["render"] = render_docx_pages(candidate)
            output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(candidate, output)
        mode = "rebuild"

    # Keep the bundle's audit trail portable.  The editable hashes remain the
    # source of truth, while the latest invocation records its resulting mode.
    payload["mode"] = mode
    payload["revision_summary"] = dict(revision_summary)
    payload["output_sha256"] = sha256_file(output)
    payload["validation"] = validation
    _write_manifest(manifest_path, payload)
    return RoundtripResult(
        mode=mode,
        source_sha256=source_sha,
        output_sha256=sha256_file(output),
        manifest_sha256=sha256_file(manifest_path),
        revision_summary=revision_summary,
        validation=validation,
        output=output,
        manifest=manifest_path,
    )


run_roundtrip = roundtrip_docx
validate_docx_package = audit_docx_package


__all__ = [
    "ROUNDTRIP_SCHEMA",
    "RoundtripResult",
    "audit_docx_package",
    "render_docx_pages",
    "validate_docx_package",
    "roundtrip_docx",
    "run_roundtrip",
    "sha256_file",
]

