"""Destinations and verification for template assembly sidecars."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path, PurePosixPath, PureWindowsPath

import pytest
from docx import Document

from docforge.cli import build_parser, main
from docforge.markdown import assemble_markdown_template, verify_assembly_checksum, write_assembly_sidecars
from docforge.markdown.template import (
    AssemblyResult,
    checksum_recorded_name,
    refresh_assembly_sidecars,
    resolve_recorded_docx,
)


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _assemble(directory: Path, name: str = "output.docx") -> tuple[AssemblyResult, Path, Path, Path, Path]:
    directory.mkdir(parents=True, exist_ok=True)
    template = directory / "template.docx"
    document = Document()
    document.add_paragraph("Sample body")
    document.save(template)
    metadata = directory / "metadata.md"
    metadata.write_text("# TITLE\n\nSynthetic title\n", encoding="utf-8")
    source = directory / "source.md"
    source.write_text("Body text.\n", encoding="utf-8")
    output = directory / name
    result = assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata,
    )
    return result, output, source, template, metadata


def _write(directory: Path, name: str = "output.docx", **destinations: Path) -> tuple[Path, Path, Path]:
    result, output, source, template, metadata = _assemble(directory, name)
    manifest, checksum = write_assembly_sidecars(
        result,
        inputs=[source],
        template_path=template,
        metadata_path=metadata,
        bibliography_path=None,
        command=["docforge", "md2docx"],
        **destinations,
    )
    return output, manifest, checksum


def _cli(directory: Path, output: Path, *extra: str) -> int:
    template = directory / "template.docx"
    if not template.exists():
        document = Document()
        document.add_paragraph("Sample body")
        document.save(template)
    metadata = directory / "metadata.md"
    if not metadata.exists():
        metadata.write_text("# TITLE\n\nSynthetic title\n", encoding="utf-8")
    source = directory / "source.md"
    if not source.exists():
        source.write_text("Body text.\n", encoding="utf-8")
    return main([
        "md2docx", str(source), "-o", str(output),
        "--template", str(template), "--metadata", str(metadata),
        *extra,
    ])


def test_default_sidecars_stay_beside_docx_and_match_its_bytes(tmp_path: Path) -> None:
    output, manifest, checksum = _write(tmp_path)
    assert manifest == output.with_suffix(".manifest.json")
    assert checksum == output.with_suffix(".sha256")
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert Path(payload["output"]["path"]) == output.resolve()
    assert payload["output"]["sha256"] == _digest(output)
    assert payload["output"]["size_bytes"] == output.stat().st_size
    text = checksum.read_text(encoding="utf-8")
    assert text == f"{_digest(output)}  {output.name}\n"
    assert verify_assembly_checksum(checksum, output=output) == _digest(output)


def test_manifest_dir_leaves_checksum_beside_docx(tmp_path: Path) -> None:
    output = tmp_path / "output.docx"
    audit = tmp_path / "audit"
    assert _cli(tmp_path, output, "--manifest-dir", str(audit)) == 0
    assert output.is_file()
    assert (audit / "output.manifest.json").is_file()
    assert output.with_suffix(".sha256").is_file()
    assert not (audit / "output.sha256").exists()
    assert not output.with_suffix(".manifest.json").exists()
    assert verify_assembly_checksum(output.with_suffix(".sha256"), output=output) == _digest(output)
    recorded = json.loads((audit / "output.manifest.json").read_text(encoding="utf-8"))
    assert recorded["output"]["sha256"] == _digest(output)
    assert Path(recorded["output"]["path"]) == output.resolve()


def test_sidecar_dir_stores_manifest_and_checksum(tmp_path: Path) -> None:
    output = tmp_path / "output.docx"
    audit = tmp_path / "audit"
    assert _cli(tmp_path, output, "--sidecar-dir", str(audit)) == 0
    assert output.is_file()
    assert not output.with_suffix(".sha256").exists()
    assert not output.with_suffix(".manifest.json").exists()
    checksum = audit / "output.sha256"
    manifest = audit / "output.manifest.json"
    assert manifest.is_file() and checksum.is_file()
    assert checksum.read_text(encoding="utf-8") == f"{_digest(output)}  ../{output.name}\n"
    assert verify_assembly_checksum(checksum, output=output) == _digest(output)


def test_nested_sidecar_directory_records_a_relative_path(tmp_path: Path) -> None:
    output = tmp_path / "deliverables" / "output.docx"
    output.parent.mkdir()
    audit = tmp_path / "audit files" / "nested" / "run"
    assert _cli(tmp_path, output, "--sidecar-dir", str(audit)) == 0
    checksum = audit / "output.sha256"
    assert checksum.is_file()
    assert (audit / "output.manifest.json").is_file()
    assert checksum.read_text(encoding="utf-8").endswith("  ../../../deliverables/output.docx\n")
    assert verify_assembly_checksum(checksum, output=output) == _digest(output)
    assert resolve_recorded_docx(checksum, "../../../deliverables/output.docx") == output.resolve()


def test_explicit_checksum_path_overrides_sidecar_dir(tmp_path: Path) -> None:
    output, manifest, checksum = _write(
        tmp_path,
        manifest_dir=tmp_path / "manifests",
        sidecar_dir=tmp_path / "audit",
        checksum_path=tmp_path / "hashes" / "nested" / "output.sha256",
    )
    assert manifest == tmp_path / "manifests" / "output.manifest.json"
    assert checksum == tmp_path / "hashes" / "nested" / "output.sha256"
    assert not (tmp_path / "audit" / "output.sha256").exists()
    assert not (tmp_path / "audit" / "output.manifest.json").exists()
    assert not output.with_suffix(".sha256").exists()
    assert verify_assembly_checksum(checksum, output=output) == _digest(output)
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["output"]["sha256"] == _digest(output)


def test_existing_sidecar_targets_are_replaced_together(tmp_path: Path) -> None:
    result, output, source, template, metadata = _assemble(tmp_path)
    audit = tmp_path / "audit" / "nested"
    manifest, checksum = write_assembly_sidecars(
        result,
        inputs=[source],
        template_path=template,
        metadata_path=metadata,
        bibliography_path=None,
        command=["docforge", "md2docx"],
        sidecar_dir=audit,
    )
    previous = _digest(output)
    manifest.write_bytes(b"stale-manifest")
    checksum.write_bytes(b"stale-checksum")
    output.write_bytes(output.read_bytes() + b"external")
    replaced_manifest, replaced_checksum = write_assembly_sidecars(
        result,
        inputs=[source],
        template_path=template,
        metadata_path=metadata,
        bibliography_path=None,
        command=["docforge", "md2docx"],
        sidecar_dir=audit,
    )
    assert replaced_manifest == manifest and replaced_checksum == checksum
    assert _digest(output) != previous
    assert b"stale-manifest" not in manifest.read_bytes()
    assert b"stale-checksum" not in checksum.read_bytes()
    assert json.loads(manifest.read_text(encoding="utf-8"))["output"]["sha256"] == _digest(output)
    assert verify_assembly_checksum(checksum, output=output) == _digest(output)


def test_cli_force_replaces_existing_audit_sidecars(tmp_path: Path) -> None:
    output = tmp_path / "output.docx"
    audit = tmp_path / "audit"
    assert _cli(tmp_path, output, "--sidecar-dir", str(audit)) == 0
    first = _digest(output)
    (tmp_path / "source.md").write_text("Body text changed.\n", encoding="utf-8")
    (audit / "output.manifest.json").write_bytes(b"stale-manifest")
    (audit / "output.sha256").write_bytes(b"stale-checksum")
    assert _cli(tmp_path, output, "--sidecar-dir", str(audit), "--force") == 0
    assert _digest(output) != first
    assert b"stale-manifest" not in (audit / "output.manifest.json").read_bytes()
    assert b"stale-checksum" not in (audit / "output.sha256").read_bytes()
    assert verify_assembly_checksum(audit / "output.sha256", output=output) == _digest(output)


def test_failed_publish_leaves_existing_pair_unchanged(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    output, manifest, checksum = _write(tmp_path, sidecar_dir=tmp_path / "audit")
    stale = {manifest: b'{"stale": true}\n', checksum: b"stale checksum\n"}
    for path, payload in stale.items():
        path.write_bytes(payload)
    import docforge.markdown.template as template_module

    original = template_module.os.replace
    failed = {"done": False}

    def fail_checksum(source: Path | str, destination: Path | str) -> None:
        src_path, dst_path = Path(source), Path(destination)
        if not failed["done"] and dst_path == checksum and src_path.name.endswith(".tmp"):
            failed["done"] = True
            raise OSError("simulated sidecar failure")
        original(source, destination)

    monkeypatch.setattr(template_module.os, "replace", fail_checksum)
    result, _output, source, template, metadata = _assemble(tmp_path / "rebuild")
    with pytest.raises(RuntimeError, match="previous sidecar pair was left unchanged"):
        write_assembly_sidecars(
            result,
            inputs=[source],
            template_path=template,
            metadata_path=metadata,
            bibliography_path=None,
            command=["docforge", "md2docx"],
            sidecar_dir=tmp_path / "audit",
        )
    assert {path: path.read_bytes() for path in stale} == stale
    leftovers = [path for path in checksum.parent.glob(".*") if ".tmp" in path.name or ".bak" in path.name]
    assert leftovers == []


def test_failed_first_publish_removes_partial_sidecars(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    result, output, source, template, metadata = _assemble(tmp_path)
    audit = tmp_path / "audit"
    checksum = audit / output.name.replace(".docx", ".sha256")
    import docforge.markdown.template as template_module

    original = template_module.os.replace
    failed = {"done": False}

    def fail_checksum(source_path: Path | str, destination: Path | str) -> None:
        if not failed["done"] and Path(destination) == checksum and Path(source_path).name.endswith(".tmp"):
            failed["done"] = True
            raise OSError("simulated sidecar failure")
        original(source_path, destination)

    monkeypatch.setattr(template_module.os, "replace", fail_checksum)
    with pytest.raises(RuntimeError, match="previous sidecar pair was left unchanged"):
        write_assembly_sidecars(
            result,
            inputs=[source],
            template_path=template,
            metadata_path=metadata,
            bibliography_path=None,
            command=["docforge", "md2docx"],
            sidecar_dir=audit,
        )
    assert not (audit / "output.manifest.json").exists()
    assert not checksum.exists()
    assert failed["done"] is True


def test_invalid_destination_does_not_create_sidecars(tmp_path: Path) -> None:
    result, output, source, template, metadata = _assemble(tmp_path)
    blocker = tmp_path / "audit"
    blocker.write_text("not a directory", encoding="utf-8")
    with pytest.raises(ValueError, match="not a directory"):
        write_assembly_sidecars(
            result,
            inputs=[source],
            template_path=template,
            metadata_path=metadata,
            bibliography_path=None,
            command=["docforge", "md2docx"],
            sidecar_dir=blocker,
        )
    with pytest.raises(ValueError, match="must differ from the DOCX"):
        write_assembly_sidecars(
            result,
            inputs=[source],
            template_path=template,
            metadata_path=metadata,
            bibliography_path=None,
            command=["docforge", "md2docx"],
            checksum_path=output,
        )
    assert blocker.read_text(encoding="utf-8") == "not a directory"
    assert not output.with_suffix(".sha256").exists()
    assert not output.with_suffix(".manifest.json").exists()


def test_refresh_after_external_rewrite_updates_both_sidecars(tmp_path: Path) -> None:
    output, manifest, checksum = _write(tmp_path, checksum_path=tmp_path / "audit" / "output.sha256")
    assembled_at = json.loads(manifest.read_text(encoding="utf-8"))["generated_at_utc"]
    title = json.loads(manifest.read_text(encoding="utf-8"))["title"]
    output.write_bytes(output.read_bytes() + b"external-post-process")
    with pytest.raises(ValueError, match="does not match"):
        verify_assembly_checksum(checksum, output=output)
    refresh_assembly_sidecars(output, manifest_path=manifest, checksum_path=checksum)
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["generated_at_utc"] == assembled_at
    assert payload["title"] == title
    assert payload["output"]["sha256"] == _digest(output)
    assert payload["output"]["size_bytes"] == output.stat().st_size
    assert Path(payload["output"]["path"]) == output.resolve()
    assert verify_assembly_checksum(checksum, output=output) == payload["output"]["sha256"]
    digest, _recorded = checksum.read_text(encoding="utf-8").split("  ", 1)
    assert digest == payload["output"]["sha256"]


def test_failed_refresh_keeps_the_previous_pair(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    output, manifest, checksum = _write(tmp_path)
    previous = {manifest: manifest.read_bytes(), checksum: checksum.read_bytes()}
    output.write_bytes(output.read_bytes() + b"changed")
    import docforge.markdown.template as template_module

    original = template_module.os.replace

    def fail_checksum(source: Path | str, destination: Path | str) -> None:
        if Path(destination) == checksum and Path(source).name.endswith(".tmp"):
            raise OSError("simulated refresh failure")
        original(source, destination)

    monkeypatch.setattr(template_module.os, "replace", fail_checksum)
    with pytest.raises(RuntimeError, match="previous sidecar pair was left unchanged"):
        refresh_assembly_sidecars(output, manifest_path=manifest, checksum_path=checksum)
    assert manifest.read_bytes() == previous[manifest]
    assert checksum.read_bytes() == previous[checksum]
    with pytest.raises(ValueError, match="does not match"):
        verify_assembly_checksum(checksum, output=output)


def test_checksum_verification_rejects_a_different_docx(tmp_path: Path) -> None:
    output, _manifest, checksum = _write(tmp_path)
    other = tmp_path / "other.docx"
    other.write_bytes(output.read_bytes() + b"other")
    with pytest.raises(ValueError, match="resolves to"):
        verify_assembly_checksum(checksum, output=other)


@pytest.mark.parametrize(
    ("flavor", "docx", "checksum", "expected"),
    [
        (PurePosixPath, "/work/output.docx", "/work/output.sha256", "output.docx"),
        (PurePosixPath, "/work/output.docx", "/work/audit/output.sha256", "../output.docx"),
        (
            PurePosixPath,
            "/work/deliverables/output.docx",
            "/work/audit/nested/output.sha256",
            "../../deliverables/output.docx",
        ),
        (
            PurePosixPath,
            "/work/my output.docx",
            "/work/audit files/nested/output.sha256",
            "../../my output.docx",
        ),
        (PureWindowsPath, r"C:\work\output.docx", r"C:\work\output.sha256", "output.docx"),
        (PureWindowsPath, r"C:\work\output.docx", r"C:\work\audit\output.sha256", "../output.docx"),
        (
            PureWindowsPath,
            r"C:\work\deliverables\output.docx",
            r"C:\work\audit\nested\output.sha256",
            "../../deliverables/output.docx",
        ),
        (PureWindowsPath, r"C:\work\output.docx", r"D:\audit\output.sha256", "C:/work/output.docx"),
        (PureWindowsPath, r"C:\work\output.docx", r"c:\work\audit\output.sha256", "../output.docx"),
        (
            PureWindowsPath,
            r"\\server\share\output.docx",
            r"\\server\share\audit\output.sha256",
            "../output.docx",
        ),
        (
            PureWindowsPath,
            r"\\server\share\output.docx",
            r"\\other\share\output.sha256",
            "//server/share/output.docx",
        ),
    ],
)
def test_posix_and_windows_recorded_names(flavor: type[PurePosixPath] | type[PureWindowsPath], docx: str, checksum: str, expected: str) -> None:
    assert checksum_recorded_name(flavor(docx), flavor(checksum)) == expected


def test_recorded_windows_absolute_path_is_rejected_on_posix(tmp_path: Path) -> None:
    checksum = tmp_path / "output.sha256"
    checksum.write_text("a" * 64 + "  C:/work/output.docx\n", encoding="utf-8")
    if os.name == "nt":
        resolved = resolve_recorded_docx(checksum, "C:/work/output.docx")
        assert resolved == Path("C:/work/output.docx")
    else:
        with pytest.raises(ValueError, match="Windows absolute path"):
            resolve_recorded_docx(checksum, "C:/work/output.docx")
    with pytest.raises(ValueError, match="not usable"):
        resolve_recorded_docx(checksum, r"..\output.docx")


def test_same_checksum_and_docx_path_is_rejected() -> None:
    with pytest.raises(ValueError, match="must differ"):
        checksum_recorded_name(PurePosixPath("/work/output.docx"), PurePosixPath("/work/output.docx"))
    with pytest.raises(ValueError, match="must differ"):
        checksum_recorded_name(PureWindowsPath(r"C:\work\output.docx"), PureWindowsPath(r"c:\work\output.docx"))


def test_destination_flags_require_template(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("Body text.\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    for flag, value in (
        ("--manifest-dir", "audit"),
        ("--sidecar-dir", "audit"),
        ("--checksum-path", "audit/output.sha256"),
    ):
        assert main(["md2docx", str(source), "-o", str(output), flag, value]) == 2
    assert not output.exists()
    assert not (tmp_path / "audit").exists()


def test_parser_accepts_sidecar_destinations() -> None:
    args = build_parser().parse_args([
        "md2docx", "source.md", "-o", "out.docx",
        "--manifest-dir", "manifests",
        "--sidecar-dir", "audit",
        "--checksum-path", "hashes/output.sha256",
    ])
    assert args.manifest_dir == Path("manifests")
    assert args.sidecar_dir == Path("audit")
    assert args.checksum_path == Path("hashes/output.sha256")
    default = build_parser().parse_args(["md2docx", "source.md", "-o", "out.docx"])
    assert default.manifest_dir is None
    assert default.sidecar_dir is None
    assert default.checksum_path is None


def test_normalized_relative_checksum_path_keeps_the_sibling_filename(tmp_path: Path) -> None:
    output, _manifest, checksum = _write(
        tmp_path,
        checksum_path=tmp_path / "audit" / ".." / "output.sha256",
    )
    assert checksum.resolve() == output.with_suffix(".sha256").resolve()
    assert checksum.resolve().read_text(encoding="utf-8") == f"{_digest(output)}  {output.name}\n"
