from __future__ import annotations

import hashlib
import json
from pathlib import Path

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
