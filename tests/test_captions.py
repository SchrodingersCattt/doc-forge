"""Authored table numbers and image captions for standalone and template assembly."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from PIL import Image

from docforge.adapters.markdown import MarkdownReader, render_latex, render_markdown
from docforge.adapters.protocol import ReadOptions
from docforge.cli import main
from docforge.markdown import assemble_markdown_template, parse_markdown, render_blocks_to_doc, write_assembly_sidecars
from docforge.markdown.contract import _block_audit


def _png(path: Path) -> None:
    Image.new("RGB", (80, 40), "white").save(path)


def _template(path: Path) -> None:
    document = Document()
    document.add_paragraph("Placeholder")
    document.save(path)


def _visible(document: Document) -> list[str]:
    return [
        paragraph.text
        for paragraph in document.paragraphs
        if paragraph.text.strip() and paragraph.text.strip() != "Fixture"
    ]


def _descriptions(document: Document) -> list[str]:
    return [node.get("descr") or "" for node in document.element.iter(qn("wp:docPr"))]


def _cnv_descriptions(document: Document) -> list[str]:
    return [node.get("descr") or "" for node in document.element.iter(qn("pic:cNvPr"))]


def _render_pair(tmp_path: Path, markdown: str) -> tuple[Document, Document]:
    _png(tmp_path / "figure.png")
    source = tmp_path / "input.md"
    source.write_text(markdown, encoding="utf-8")
    standalone = render_blocks_to_doc(parse_markdown(source))
    template = tmp_path / "template.docx"
    _template(template)
    output = tmp_path / "output.docx"
    assemble_markdown_template(
        [source], template_path=template, output=output, title="Fixture",
    )
    return standalone, Document(output)


def test_table_directive_with_authored_number_is_rejected(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    source = tmp_path / "input.md"
    source.write_text(
        "Table: Table 5. Measurements.\n\n| Item | Value |\n| --- | --- |\n| A | 1 |\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="table.caption_ambiguous"):
        parse_markdown(source)
    output = tmp_path / "output.docx"
    assert main(["md2docx", str(source), "-o", str(output)]) == 2
    captured = capsys.readouterr()
    assert "table.caption_ambiguous" in captured.err
    assert "Table 1. Table 5" not in captured.err
    assert not output.exists()

    template = tmp_path / "template.docx"
    _template(template)
    templated = tmp_path / "templated.docx"
    with pytest.raises(ValueError, match="table.caption_ambiguous"):
        assemble_markdown_template([source], template_path=template, output=templated, title="Fixture")
    assert not templated.exists()


def test_emphasized_authored_number_inside_table_directive_is_rejected(tmp_path: Path) -> None:
    source = tmp_path / "input.md"
    source.write_text("Table: **Table 5.** Measurements.\n\n| A |\n| --- |\n| 1 |\n", encoding="utf-8")
    with pytest.raises(ValueError, match="table.caption_ambiguous"):
        parse_markdown(source)


def test_table_mention_inside_automatic_caption_stays_single_numbered(tmp_path: Path) -> None:
    source = tmp_path / "input.md"
    source.write_text(
        "Table: See Table 5. Later measurements.\n\n| A |\n| --- |\n| 1 |\n",
        encoding="utf-8",
    )
    document = render_blocks_to_doc(parse_markdown(source))
    assert _visible(document) == ["Table 1. See Table 5. Later measurements."]


def test_authored_table_number_is_kept_and_recorded(tmp_path: Path) -> None:
    source = tmp_path / "input.md"
    source.write_text(
        "Table P4. Measurements.\n\n| Item | Value |\n| --- | --- |\n| A | 1 |\n\n"
        "Table: Next values.\n\n| Item |\n| --- |\n| B |\n",
        encoding="utf-8",
    )
    standalone = render_blocks_to_doc(parse_markdown(source), number_prefix="P")
    assert _visible(standalone) == ["Table P4. Measurements.", "Table P1. Next values."]

    template = tmp_path / "template.docx"
    _template(template)
    output = tmp_path / "output.docx"
    result = assemble_markdown_template(
        [source], template_path=template, output=output, title="Fixture", numbering_prefix="P",
    )
    assert result.tables == (
        {
            "number": 4,
            "label": "Table P4",
            "caption": "Measurements.",
            "role": "table_caption",
            "numbering": "authored",
            "span": "column",
        },
        {
            "number": 1,
            "label": "Table P1",
            "caption": "Next values.",
            "role": "table_caption",
            "numbering": "generated",
            "span": "column",
        },
    )
    text = "\n".join(_visible(Document(output)))
    assert "Table P4. Measurements." in text
    assert "Table P1. Next values." in text
    assert "Supplementary" not in text
    manifest, _checksum = write_assembly_sidecars(
        result,
        inputs=[source],
        template_path=template,
        metadata_path=None,
        bibliography_path=None,
        command=["docforge", "md2docx"],
    )
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["tables"] == [dict(table) for table in result.tables]
    assert payload["numbering_prefix"] == "P"
    audit = _block_audit([source], keep_comments=False)
    captions = [record for record in audit if record["kind"] == "table_caption"]
    assert captions[0]["role"] == "table_caption"
    assert captions[0]["numbering"] == "authored"
    assert captions[0]["label"] == "Table P4"
    assert captions[1]["numbering"] == "generated"
    assert "label" not in captions[1]


def test_authored_and_generated_prefixes_report_collisions(tmp_path: Path) -> None:
    source = tmp_path / "input.md"
    source.write_text(
        "Table: First.\n\n| A |\n| --- |\n| a |\n\n"
        "Table P1. Authored.\n\n| A |\n| --- |\n| b |\n\n"
        "Table: Next.\n\n| A |\n| --- |\n| c |\n",
        encoding="utf-8",
    )
    with pytest.warns(UserWarning, match=r"table\.number_collision: automatic table caption skipped Table P1"):
        document = render_blocks_to_doc(parse_markdown(source), number_prefix="P")
    assert _visible(document) == ["Table P2. First.", "Table P1. Authored.", "Table P3. Next."]

    template = tmp_path / "template.docx"
    _template(template)
    with pytest.warns(UserWarning, match="table.number_collision"):
        result = assemble_markdown_template(
            [source],
            template_path=template,
            output=tmp_path / "output.docx",
            title="Fixture",
            numbering_prefix="P",
        )
    assert [table["label"] for table in result.tables] == ["Table P2", "Table P1", "Table P3"]
    assert [table["numbering"] for table in result.tables] == ["generated", "authored", "generated"]


def test_table_caption_round_trip_keeps_role_and_number(tmp_path: Path) -> None:
    source = tmp_path / "input.md"
    source.write_text(
        "Table P4. Measurements.\n\n"
        "| First | Second |\n| :--- | ---: |\n| a | b |\n\n"
        "Table: Next values.\n\n| Only |\n| --- |\n| c |\n",
        encoding="utf-8",
    )
    document = MarkdownReader().read(source, options=ReadOptions()).document
    captions = [block for block in document.blocks if block.kind == "caption" and block.attr("kind") == "table"]
    assert [block.attr("numbering") for block in captions] == ["authored", "generated"]
    assert captions[0].attr("role") == "table_caption"
    assert captions[0].attr("label") == "Table P4"
    markdown = render_markdown(document)
    assert "Table P4. Measurements." in markdown
    assert "Table: Next values." in markdown
    assert r"\begin{tabular}{lr}" in render_latex(document)
    rewritten = tmp_path / "rewritten.md"
    rewritten.write_text(markdown, encoding="utf-8")
    blocks = parse_markdown(rewritten)
    assert [block.kind for block in blocks] == ["table_caption", "table", "table_caption", "table"]
    assert dict(blocks[0].options)["caption_label"] == "Table P4"
    assert "caption_label" not in dict(blocks[2].options)
    assert blocks[2].text == "Next values."


@pytest.mark.parametrize(
    ("markdown", "standalone_text", "template_text", "description"),
    [
        (
            "![Figure 1](figure.png)\n\nFigure 1. Example caption.\n",
            ["Figure 1. Example caption."],
            ["Figure 1. Example caption."],
            "Figure 1",
        ),
        (
            "![A white rectangle.](figure.png)\n",
            ["A white rectangle."],
            ["Figure 1. A white rectangle."],
            "A white rectangle.",
        ),
        (
            "![](figure.png)\n",
            [],
            [],
            "",
        ),
        (
            "![](figure.png)\n\nFigure 1. Empty alternative.\n",
            ["Figure 1. Empty alternative."],
            ["Figure 1. Empty alternative."],
            "Figure 1. Empty alternative.",
        ),
        (
            "![figure.png](figure.png)\n",
            [],
            [],
            "figure.png",
        ),
        (
            "![figure.png](figure.png)\n\nFigure 1. Named file.\n",
            ["Figure 1. Named file."],
            ["Figure 1. Named file."],
            "figure.png",
        ),
        (
            "![A white rectangle.](figure.png)\n\nOrdinary nearby prose.\n",
            ["A white rectangle.", "Ordinary nearby prose."],
            ["Figure 1. A white rectangle.", "Ordinary nearby prose."],
            "A white rectangle.",
        ),
    ],
)
def test_image_caption_follows_alt_rules(
    tmp_path: Path,
    markdown: str,
    standalone_text: list[str],
    template_text: list[str],
    description: str,
) -> None:
    standalone, templated = _render_pair(tmp_path, markdown)
    assert _visible(standalone) == standalone_text
    assert _visible(templated) == template_text
    assert _descriptions(standalone) == [description]
    assert _cnv_descriptions(standalone) == [description]
    assert _descriptions(templated) == [description]
    assert _cnv_descriptions(templated) == [description]


def test_following_caption_is_recorded_separately_from_alt(tmp_path: Path) -> None:
    _png(tmp_path / "figure.png")
    source = tmp_path / "input.md"
    source.write_text(
        "![Figure 1](figure.png)\n\nFigure 1. Example caption.\n",
        encoding="utf-8",
    )
    output = tmp_path / "output.docx"
    assert main(["md2docx", str(source), "-o", str(output)]) == 0
    document = Document(output)
    assert _visible(document) == ["Figure 1. Example caption."]
    assert _descriptions(document) == ["Figure 1"]

    template = tmp_path / "template.docx"
    _template(template)
    templated = tmp_path / "templated.docx"
    result = assemble_markdown_template(
        [source], template_path=template, output=templated, title="Fixture",
    )
    assert result.figures[0]["alt"] == "Figure 1"
    assert result.figures[0]["caption"] == "Figure 1. Example caption."
    assert _visible(Document(templated)) == ["Figure 1. Example caption."]
    manifest, _checksum = write_assembly_sidecars(
        result,
        inputs=[source],
        template_path=template,
        metadata_path=None,
        bibliography_path=None,
        command=["docforge", "md2docx"],
    )
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["figures"][0]["alt"] == "Figure 1"
    assert payload["figures"][0]["caption"] == "Figure 1. Example caption."
