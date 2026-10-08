"""Citation styles stay off chemical subscripts and superscripts."""

from __future__ import annotations

from pathlib import Path

from docx.oxml.ns import qn

from docforge.markdown.citations import CitationStyle, materialize_citations, replace_block_citations, replace_citations
from docforge.markdown.launcher import parse_markdown, render_blocks_to_doc


def _aligned_runs(document) -> list[tuple[str, str | None]]:
    found = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            properties = run._r.find(qn("w:rPr"))
            align = None
            if properties is not None:
                marker = properties.find(qn("w:vertAlign"))
                if marker is not None:
                    align = marker.get(qn("w:val"))
            found.append((run.text, align))
    return found


def _render(tmp_path: Path, style: CitationStyle):
    source = tmp_path / "source.md"
    source.write_text("离子$\\mathrm{NH_4^+}$见\\citep{a}。\n", encoding="utf-8")
    blocks = replace_block_citations(parse_markdown(source), {"a": 1}, style)
    document = render_blocks_to_doc(blocks)
    materialize_citations(document, style)
    return _aligned_runs(document)


def test_adjacent_citations_collapse_to_an_en_dash_range() -> None:
    text = "See \\citep{a} \\citep{b} \\citep{c} and \\citep{e}."
    rendered = replace_citations(text, {"a": 4, "b": 5, "c": 6, "e": 8}, CitationStyle.SUPERSCRIPT)
    assert rendered == "See\ue0004–6\ue001 and\ue0008\ue001."


def test_unknown_citation_style_is_rejected() -> None:
    try:
        CitationStyle.by_name("footnote")
    except ValueError as exc:
        assert "superscript-bracketed" in str(exc)
    else:
        raise AssertionError("unknown citation style was accepted")


def test_superscript_brackets_leave_formula_scripts(tmp_path: Path) -> None:
    runs = _render(tmp_path, CitationStyle.SUPERSCRIPT_BRACKETED)
    assert ("4", "subscript") in runs
    assert ("+", "superscript") in runs
    assert ("[1]", "superscript") in runs
    assert all(text != "[4]" for text, _ in runs)


def test_bare_superscript_citations_stay_unbracketed(tmp_path: Path) -> None:
    runs = _render(tmp_path, CitationStyle.SUPERSCRIPT)
    assert ("4", "subscript") in runs
    assert ("1", "superscript") in runs
    assert all(text != "[1]" for text, _ in runs)


def test_bracketed_citations_stay_on_the_baseline(tmp_path: Path) -> None:
    runs = _render(tmp_path, CitationStyle.BRACKETED)
    assert ("4", "subscript") in runs
    assert any("[1]" in text and align != "superscript" for text, align in runs)
    assert all(text != "[1]" or align != "superscript" for text, align in runs)
