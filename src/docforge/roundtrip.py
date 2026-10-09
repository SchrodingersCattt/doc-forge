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
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

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
    output: Path | None = None
    manifest: Path | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "source_sha256": self.source_sha256,
            "output_sha256": self.output_sha256,
            "manifest_sha256": self.manifest_sha256,
            "revision_summary": dict(self.revision_summary),
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
    revision_summary: Mapping[str, int] = {}
    if unchanged:
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, output)
        mode = "exact-reuse"
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
            if baseline is not None:
                from .docxdiff import create_tracked_docx

                revision_summary = create_tracked_docx(
                    baseline.resolve(), rebuilt, output, overwrite=True
                )
            else:
                output.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(rebuilt, output)
        mode = "rebuild"

    # Keep the bundle's audit trail portable.  The editable hashes remain the
    # source of truth, while the latest invocation records its resulting mode.
    payload["mode"] = mode
    payload["revision_summary"] = dict(revision_summary)
    payload["output_sha256"] = sha256_file(output)
    _write_manifest(manifest_path, payload)
    return RoundtripResult(
        mode=mode,
        source_sha256=source_sha,
        output_sha256=sha256_file(output),
        manifest_sha256=sha256_file(manifest_path),
        revision_summary=revision_summary,
        output=output,
        manifest=manifest_path,
    )


run_roundtrip = roundtrip_docx


__all__ = ["ROUNDTRIP_SCHEMA", "RoundtripResult", "roundtrip_docx", "run_roundtrip", "sha256_file"]
