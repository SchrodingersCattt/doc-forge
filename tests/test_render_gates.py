"""Layout gates. A skip is not a pass, and Word visual parity requires Word."""

from __future__ import annotations

from pathlib import Path

from docx import Document

from docforge.docxmd.verify import _soffice, _word_available
from docforge.gates.render import counts_as_pass, layout_gate


def test_skip_is_not_a_pass() -> None:
    assert counts_as_pass({"status": "skipped", "visual_parity": "not-passed"}) is False
    assert counts_as_pass({"status": "failed", "visual_parity": "not-passed"}) is False
    assert counts_as_pass({"status": "passed", "visual_parity": "not-passed"}) is True


def test_word_visual_parity_is_not_passed_without_word(tmp_path: Path) -> None:
    source = _docx(tmp_path / "page.docx")
    report = layout_gate(source, tmp_path / "word", "word")
    if not _word_available():
        assert report["status"] == "skipped"
        assert report["resolved_engine"] is None
        assert report["visual_parity"] == "not-passed"
        assert counts_as_pass(report) is False
        assert "not a pass" in str(report["reason"])
        return
    assert report["status"] == "passed"
    assert report["resolved_engine"] == "word"
    assert report["visual_parity"] == "passed"
    assert counts_as_pass(report) is True
    assert int(report["pdf_bytes"] or 0) > 0


def test_libreoffice_gate_skips_or_renders(tmp_path: Path) -> None:
    source = _docx(tmp_path / "page.docx")
    report = layout_gate(source, tmp_path / "libreoffice", "libreoffice")
    if _soffice() is None:
        assert report["status"] == "skipped"
        assert report["resolved_engine"] is None
        assert report["visual_parity"] == "not-passed"
        assert counts_as_pass(report) is False
        assert "not a pass" in str(report["reason"])
        return
    assert report["status"] == "passed"
    assert report["resolved_engine"] == "libreoffice"
    assert report["visual_parity"] == "not-passed"
    assert counts_as_pass(report) is True
    assert int(report["pdf_bytes"] or 0) > 0


def test_auto_does_not_claim_word_parity_without_word(tmp_path: Path) -> None:
    source = _docx(tmp_path / "page.docx")
    report = layout_gate(source, tmp_path / "auto", "auto")
    if not _word_available():
        assert report["visual_parity"] == "not-passed"
    if _soffice() is None and not _word_available():
        assert report["status"] == "skipped"
        assert counts_as_pass(report) is False


def _docx(path: Path) -> Path:
    document = Document()
    document.add_paragraph("Render gate sentence.")
    document.save(path)
    return path
