from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from docforge.docxmd import roundtrip


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _bundle(root: Path, source: Path, text: str = "Original\n") -> Path:
    work = root / "bundle"
    work.mkdir()
    snapshot = work / "roundtrip" / "source.docx"
    snapshot.parent.mkdir()
    snapshot.write_bytes(source.read_bytes())
    section = work / "main.md"
    section.write_text(text, encoding="utf-8")
    manifest = {
        "schema": "docforge.docx2md.v1",
        "source": source.name,
        "source_sha256": _sha(source),
        "source_copy": "roundtrip/source.docx",
        "source_copy_sha256": _sha(snapshot),
        "sections": ["main.md"],
        "section_sha256": {"main.md": _sha(section)},
        "media_sha256": {},
    }
    (work / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    return work


def test_roundtrip_reuses_unchanged_source_bytes(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package bytes")
    work = _bundle(tmp_path, source)
    output = tmp_path / "out.docx"
    report = roundtrip(source, work, section_map=tmp_path / "map.json", output=output)
    assert report["mode"] == "exact-reuse"
    assert report["exact_source_reused"] is True
    assert output.read_bytes() == source.read_bytes()
    audit = json.loads(output.with_suffix(".manifest.json").read_text())
    assert audit["mode"] == "exact-reuse"
    assert report["output_sha256"] == _sha(output)
    assert report["source_sha256"] == _sha(source)


def test_roundtrip_audit_names_package_patch_mode(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package bytes")
    work = _bundle(tmp_path, source)
    (work / "main.md").write_text("Edited\n", encoding="utf-8")

    def fake_export(_source, *, split_dir, **_kwargs):
        (split_dir / "main.md").write_text("Original\n", encoding="utf-8")

    def fake_delta(source_path, _baseline, _edited, output, **_kwargs):
        output.write_bytes(source_path.read_bytes() + b" patched")
        return {"replaced": 1, "inserted": 0}

    monkeypatch.setattr("docforge.docxmd.verify.docx_to_markdown", fake_export, raising=False)
    monkeypatch.setattr("docforge.markdown.docx_export.docx_to_markdown", fake_export)
    monkeypatch.setattr("docforge.markdown.source_edit.apply_markdown_delta", fake_delta)
    # The command imports helpers lazily; its module-level monkeypatch target
    # remains intentionally explicit in this test for API documentation.
    report = roundtrip(source, work, section_map=tmp_path / "map.json", output=tmp_path / "out.docx")
    assert report["mode"] == "rebuild"
    assert report["output_sha256"] == _sha(tmp_path / "out.docx")


def test_roundtrip_checks_output_before_exact_source_reuse(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package bytes")
    work = _bundle(tmp_path, source)
    output = tmp_path / "out.docx"
    output.write_bytes(b"KEEP")

    with pytest.raises(FileExistsError, match="out.docx"):
        roundtrip(source, work, section_map=tmp_path / "map.json", output=output)
    assert output.read_bytes() == b"KEEP"


def test_roundtrip_validates_manifest_paths_before_publishing(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package bytes")
    work = _bundle(tmp_path, source)
    outside = tmp_path / "outside.md"
    outside.write_text("outside", encoding="utf-8")
    payload = json.loads((work / "manifest.json").read_text())
    payload["section_sha256"] = {"../outside.md": _sha(outside)}
    (work / "manifest.json").write_text(json.dumps(payload), encoding="utf-8")
    output = tmp_path / "out.docx"

    with pytest.raises(ValueError, match="escapes the workdir"):
        roundtrip(source, work, section_map=tmp_path / "map.json", output=output)
    assert not output.exists()


def test_roundtrip_rejects_track_policy_mismatch(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package bytes")
    work = _bundle(tmp_path, source)
    payload = json.loads((work / "manifest.json").read_text())
    payload["track_changes"] = "all"
    (work / "manifest.json").write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(ValueError, match="track_changes"):
        roundtrip(source, work, section_map=tmp_path / "map.json", output=tmp_path / "out.docx")


def test_roundtrip_redline_failure_does_not_publish_partial_output(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package bytes")
    work = _bundle(tmp_path, source)
    (work / "main.md").write_text("Edited\n", encoding="utf-8")

    def fake_export(_source, *, split_dir, **_kwargs):
        (split_dir / "main.md").write_text("Original\n", encoding="utf-8")

    def fake_delta(source_path, _baseline, _edited, output, **_kwargs):
        output.write_bytes(source_path.read_bytes() + b" patched")
        return {"replaced": 1, "inserted": 0}

    def fail_redline(*_args, **_kwargs):
        raise RuntimeError("redline generation failed")

    monkeypatch.setattr("docforge.docxmd.verify.docx_to_markdown", fake_export, raising=False)
    monkeypatch.setattr("docforge.markdown.docx_export.docx_to_markdown", fake_export)
    monkeypatch.setattr("docforge.markdown.source_edit.apply_markdown_delta", fake_delta)
    monkeypatch.setattr("docforge.docxdiff.redline.create_tracked_docx", fail_redline)

    output = tmp_path / "out.docx"
    audit = output.with_suffix(".manifest.json")
    redline = tmp_path / "redline.docx"
    with pytest.raises(RuntimeError, match="redline generation failed"):
        roundtrip(
            source,
            work,
            section_map=tmp_path / "map.json",
            output=output,
            redline=redline,
        )
    assert not output.exists()
    assert not audit.exists()
    assert not redline.exists()
