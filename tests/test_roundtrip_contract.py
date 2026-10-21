from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

import pytest

from docforge.roundtrip import (
    PathCollisionError,
    RoundTripMode,
    SnapshotMismatchError,
    SourceSnapshot,
    compare_projection,
    exact_reuse,
    redact_report,
)
from docforge.output import InvalidFormatError, UnsupportedFormatError, detect_format


def test_exact_reuse_verifies_snapshot_and_bytes(tmp_path: Path) -> None:
    source = tmp_path / "fixture.bin"
    source.write_bytes(b"opaque fixture")
    snapshot = tmp_path / "snapshot.bin"
    snapshot.write_bytes(source.read_bytes())
    result = exact_reuse(source, tmp_path / "copy.bin", snapshot=SourceSnapshot.from_file(snapshot))
    assert result.code == "RT-EXACT"
    assert result.status == "pass"
    assert (tmp_path / "copy.bin").read_bytes() == source.read_bytes()


def test_corrupt_snapshot_fails_before_publish(tmp_path: Path) -> None:
    source = tmp_path / "fixture.bin"
    source.write_bytes(b"source")
    snapshot = tmp_path / "snapshot.bin"
    snapshot.write_bytes(b"snapshot")
    expected = SourceSnapshot(snapshot, hashlib.sha256(b"different").hexdigest())
    with pytest.raises(SnapshotMismatchError) as error:
        exact_reuse(source, tmp_path / "copy.bin", snapshot=expected)
    assert error.value.code == "SNAPSHOT_SHA_MISMATCH"
    assert not (tmp_path / "copy.bin").exists()


def test_changed_source_does_not_take_exact_copy_path(tmp_path: Path) -> None:
    source = tmp_path / "fixture.bin"
    source.write_bytes(b"source")
    snapshot = SourceSnapshot.from_file(source)
    source.write_bytes(b"changed")
    with pytest.raises(SnapshotMismatchError):
        exact_reuse(source, tmp_path / "copy.bin", snapshot=snapshot)
    assert not (tmp_path / "copy.bin").exists()


def test_exact_reuse_rejects_source_collision_even_with_force(tmp_path: Path) -> None:
    source = tmp_path / "fixture.bin"
    source.write_bytes(b"source")
    with pytest.raises(PathCollisionError):
        exact_reuse(source, source, force=True)


def test_projection_comparison_is_semantic_and_redacted() -> None:
    before = {"content": ("alpha",), "structure": {"blocks": 1}, "style": {"font": "sans"}}
    after = {"content": ("alpha",), "structure": {"blocks": 1}, "style": {"font": "sans"}}
    results = compare_projection(before, after)
    assert all(item.status == "pass" for item in results.values())
    report = redact_report({
        "artifact_id": "F-01",
        "format": "docx",
        "mode": RoundTripMode.EXACT_REUSE.value,
        "invariants": {key: value.to_dict() for key, value in results.items()},
        "counts": {"blocks": 12, "secret_count": "not-an-int"},
        "diagnostics": [
            {"category": "content", "reason": "RT-CONTENT", "counts": {"changed": 1}, "secret": "text"}
        ],
        "source": "/private/user.docx",
        "raw_xml": "<w:p>secret</w:p>",
        "sha256": "0" * 64,
    })
    serialized = json.dumps(report, ensure_ascii=False)
    assert "private" not in serialized and "secret" not in serialized and "0" * 64 not in serialized
    assert report["artifact_id"] == "F-01"
    assert report["invariants"]["RT-CONTENT"]["status"] == "pass"
    assert report["invariants"]["RT-CONTENT"]["counts"] == {"equal": 1}
    assert report["counts"] == {"blocks": 12}
    assert report["diagnostics"][0]["counts"] == {"changed": 1}

    assert redact_report({"artifact_id": "/private/Thesis draft.docx"})["artifact_id"] == "REDACTED"
    assert redact_report({"artifact_id": "F-01"})["artifact_id"] == "F-01"


def test_docx_projection_canonicalizes_property_and_whitespace_transport() -> None:
    from lxml import etree
    from docforge.docxmd.ooxml import W, flatten
    from docforge.docxmd.reader import _sample_text
    from docforge.docxmd.verify import _canonical_tokens

    first = etree.fromstring(
        f'<w:pPr xmlns:w="{W}"><w:sectPr><w:headerReference w:id="1"/>'
        '<w:footerReference w:id="2"/></w:sectPr><w:spacing w:after="1"/></w:pPr>'
    )
    second = etree.fromstring(
        f'<w:pPr xmlns:w="{W}"><w:spacing w:after="1"/><w:sectPr>'
        '<w:footerReference w:id="2"/><w:headerReference w:id="1"/></w:sectPr></w:pPr>'
    )
    assert _canonical_tokens(flatten(first)) == _canonical_tokens(flatten(second))
    assert "\u00a0" in _sample_text("A\u00a0中")


def _package(path: Path, main: str) -> None:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(main, "<root/>")


def test_format_probe_rejects_presentation_before_staging(tmp_path: Path) -> None:
    presentation = tmp_path / "fixture.docx"
    _package(presentation, "ppt/presentation.xml")
    with pytest.raises(UnsupportedFormatError) as error:
        detect_format(presentation)
    assert error.value.code == "UNSUPPORTED_FORMAT"


def test_format_probe_requires_matching_word_package_extension(tmp_path: Path) -> None:
    renamed = tmp_path / "fixture.bin"
    _package(renamed, "word/document.xml")
    with pytest.raises(InvalidFormatError) as error:
        detect_format(renamed)
    assert error.value.code == "INVALID_FORMAT"
