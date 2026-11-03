"""Acceptance coverage for scoped table-note syntax (#63)."""
from __future__ import annotations

import warnings
from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from docforge.markdown.launcher import parse_markdown, render_blocks_to_doc
from docforge.markdown.table_notes import (
    diagnose_table_notes,
    option_value,
)
from docforge.markdown.template import assemble_markdown_template


EXPLICIT_EXAMPLE = """# Results

Table: Measurements.

| Item | Mean<sup>[a]</sup> | Error<sup>[b]</sup> |
| --- | --- | --- |
| Sample 1 | 2.4 | 0.3 |

Table note: [a] Mean value.
Table note: [b] Estimated uncertainty.

This is an ordinary paragraph with [a] and [b].
"""


def _write(path: Path, text: str) -> Path:
    path.write_text(text, encoding="utf-8")
    return path


@pytest.fixture
def template_fixture(tmp_path: Path):
    template = tmp_path / "template.docx"
    doc = Document()
    for name in (
        "Title1", "Authors", "Address", "Abstract", "P1", "H1", "H2",
        "FigureCaption", "References", "TableCaption", "TableHead", "TableBody",
        "TableFoot", "Keywords", "BackMatterHeading", "BackMatter", "TableOfContentText",
    ):
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Georgia"
        style.font.size = Pt(8.5)
        style.font.color.rgb = RGBColor.from_string("123456")
        doc.add_paragraph("Example " + name, style=name)
    doc.save(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\nA title\n# AUTHORS\nA. Writer\n", encoding="utf-8")

    def build(text: str, **options):
        source = tmp_path / "source.md"
        source.write_text(text, encoding="utf-8")
        output = tmp_path / "output.docx"
        result = assemble_markdown_template(
            [source],
            template_path=template,
            metadata_path=metadata,
            output=output,
            force=True,
            **options,
        )
        return Document(output), result

    return tmp_path, template, build


def test_explicit_example_blank_lines_and_ordinary_prose(tmp_path: Path):
    source = _write(tmp_path / "notes.md", EXPLICIT_EXAMPLE)
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        blocks = parse_markdown(source, table_notes="explicit")
    assert not any("table_note" in str(item.message) for item in caught)
    kinds = [block.kind for block in blocks]
    assert kinds.count("table") == 1
    notes = [block for block in blocks if block.kind == "table_foot"]
    assert [block.text for block in notes] == ["[a] Mean value.", "[b] Estimated uncertainty."]
    assert all(option_value(block, "table_id") == "1" for block in notes)
    assert [option_value(block, "label") for block in notes] == ["a", "b"]
    prose = next(block for block in blocks if block.kind == "paragraph" and "[a] and [b]" in block.text)
    assert prose.kind == "paragraph"


def test_explicit_mode_keeps_bracket_lines_as_paragraphs(tmp_path: Path):
    source = _write(
        tmp_path / "bracket.md",
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\n[a] Not a table note in explicit mode.\n",
    )
    blocks = parse_markdown(source, table_notes="explicit")
    assert [block.kind for block in blocks] == ["table", "paragraph"]
    assert blocks[1].text.startswith("[a] Not a table note")


def test_legacy_mode_preserves_bracket_heuristic(tmp_path: Path):
    source = _write(
        tmp_path / "legacy.md",
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\n[a] Legacy note.\n\nTable note: Another note.\n",
    )
    blocks = parse_markdown(source, table_notes="legacy")
    notes = [block for block in blocks if block.kind == "table_foot"]
    assert [block.text for block in notes] == ["[a] Legacy note.", "Another note."]
    assert all(option_value(block, "table_id") == "1" for block in notes)


def test_backslash_escape_keeps_directive_literal(tmp_path: Path):
    source = _write(
        tmp_path / "escape.md",
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\n\\Table note: literal text.\n",
    )
    blocks = parse_markdown(source, table_notes="explicit")
    assert [block.kind for block in blocks] == ["table", "paragraph"]
    assert blocks[1].text == "Table note: literal text."


def test_orphan_declaration_does_not_attach_to_earlier_table(tmp_path: Path):
    source = _write(
        tmp_path / "orphan.md",
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\nBody paragraph.\n\nTable note: [a] Orphan.\n",
    )
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        blocks = parse_markdown(source, table_notes="explicit")
    assert not any(block.kind == "table_foot" for block in blocks)
    orphan = next(block for block in blocks if block.text.startswith("Table note:"))
    assert orphan.kind == "paragraph"
    assert option_value(orphan, "table_note_orphan") == "true"
    assert any("table_note.orphan" in str(item.message) for item in caught)


def test_orphan_errors_when_policy_is_error(tmp_path: Path):
    # Legacy still emits table_foot for orphan directives; error policy fails.
    source = _write(tmp_path / "orphan-error.md", "Table note: [a] Alone.\n")
    with pytest.raises(ValueError, match="table note diagnostics failed"):
        parse_markdown(source, table_notes="legacy", table_note_diagnostics="error")


def test_multiple_tables_reuse_labels(tmp_path: Path):
    source = _write(
        tmp_path / "multi.md",
        "| X |\n| --- |\n| 1<sup>[a]</sup> |\n\nTable note: [a] First.\n\n"
        "| Y |\n| --- |\n| 2<sup>[a]</sup> |\n\nTable note: [a] Second.\n",
    )
    blocks = parse_markdown(source, table_notes="explicit")
    notes = [block for block in blocks if block.kind == "table_foot"]
    assert [option_value(block, "table_id") for block in notes] == ["1", "2"]
    assert [block.text for block in notes] == ["[a] First.", "[a] Second."]
    assert diagnose_table_notes(blocks) == []


def test_duplicate_and_missing_labels(tmp_path: Path):
    source = _write(
        tmp_path / "labels.md",
        "| Mean<sup>[a]</sup> | Err<sup>[b]</sup> |\n| --- | --- |\n| 1 | 2 |\n\n"
        "Table note: [a] One.\nTable note: [a] Dup.\n",
    )
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        blocks = parse_markdown(source, table_notes="explicit")
    messages = " ".join(str(item.message) for item in caught)
    assert "table_note.duplicate_label" in messages
    assert "table_note.missing_label" in messages
    assert option_value(next(block for block in blocks if block.kind == "table_foot"), "table_id") == "1"


def test_preserves_links_citations_code_and_math_after_table(tmp_path: Path):
    source = _write(
        tmp_path / "inline.md",
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\n"
        "See [docs](https://example.test) and \\cite{key} plus `code` and $x^2$.\n",
    )
    blocks = parse_markdown(source, table_notes="explicit")
    assert [block.kind for block in blocks] == ["table", "paragraph"]
    assert "https://example.test" in blocks[1].text
    assert "$x^2$" in blocks[1].text


def test_standalone_render_uses_table_foot_style_and_keep_next(tmp_path: Path):
    source = _write(
        tmp_path / "standalone.md",
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\nTable note: [a] Note.\nTable note: [b] More.\n",
    )
    blocks = parse_markdown(source, table_notes="explicit")
    doc = render_blocks_to_doc(blocks)
    notes = [p for p in doc.paragraphs if p.style.name == "TableFoot"]
    assert [p.text for p in notes] == ["[a] Note.", "[b] More."]
    assert notes[0].paragraph_format.keep_with_next is True
    assert doc.tables
    assert doc.tables[0].rows[-1].cells[0].paragraphs[0].paragraph_format.keep_with_next is True


def test_template_table_foot_style_and_body_fallback(template_fixture):
    root, template, build = template_fixture
    doc, result = build(EXPLICIT_EXAMPLE, table_notes="explicit")
    assert result.table_notes == "explicit"
    assert [p.text for p in doc.paragraphs if p.style.name == "TableFoot"] == [
        "[a] Mean value.",
        "[b] Estimated uncertainty.",
    ]
    prose = next(p for p in doc.paragraphs if "[a] and [b]" in p.text)
    assert prose.style.name == "P1"
    # Template without TableFoot falls back to body.
    bare = Document(template)
    bare.styles.element.remove(bare.styles["TableFoot"].element)
    bare.save(template)
    source = root / "fallback.md"
    source.write_text("| A |\n| --- |\n| 1 |\n\nTable note: [a] Fallback.\n", encoding="utf-8")
    output = root / "fallback.docx"
    assemble_markdown_template(
        [source],
        template_path=template,
        metadata_path=root / "metadata.md",
        output=output,
        force=True,
        table_notes="explicit",
    )
    fallback = Document(output)
    note = next(p for p in fallback.paragraphs if p.text == "[a] Fallback.")
    assert note.style.name == "P1"


def test_unlabelled_note_and_blank_line_association(tmp_path: Path):
    source = _write(
        tmp_path / "blank.md",
        "| A |\n| --- |\n| 1 |\n\n\nTable note: Unlabelled note.\n\nTable note: [a] Labelled.\n",
    )
    blocks = parse_markdown(source, table_notes="explicit", table_note_diagnostics="ignore")
    notes = [block for block in blocks if block.kind == "table_foot"]
    assert [block.text for block in notes] == ["Unlabelled note.", "[a] Labelled."]
    assert all(option_value(block, "table_id") == "1" for block in notes)


@pytest.mark.parametrize("engine_available", [False])
def test_visual_page_checks_reported_as_skipped(engine_available):
    """Word visual inspection for multi-page / two-column notes is skipped here."""
    if not engine_available:
        pytest.skip("Word visual checks for table-note pagination are skipped on this runner")
