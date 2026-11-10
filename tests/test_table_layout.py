"""Column alignment, explicit page span, and optional vertical grouping."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

from docforge.adapters.docx import DocxReader
from docforge.adapters.markdown import MarkdownReader, render_markdown
from docforge.adapters.protocol import ReadOptions
from docforge.cli import build_parser, main
from docforge.markdown.launcher import parse_markdown, render_blocks_to_doc
from docforge.markdown.template import assemble_markdown_template


def _metadata(path: Path) -> None:
    path.write_text("# TITLE\nSample study\n\n# AUTHOR\nA. Researcher\n", encoding="utf-8")


def _columns(section, count: int, space: str = "720") -> None:
    cols = section._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        section._sectPr.append(cols)
    if count <= 1:
        cols.attrib.pop(qn("w:num"), None)
        cols.attrib.pop(qn("w:space"), None)
        return
    cols.set(qn("w:num"), str(count))
    cols.set(qn("w:space"), space)


def _letter_section(section) -> None:
    section.page_width = Twips(12240)
    section.page_height = Twips(15840)
    section.left_margin = Twips(1800)
    section.right_margin = Twips(1800)
    section.top_margin = Twips(1440)
    section.bottom_margin = Twips(1440)


def _template(path: Path, *, columns: int = 2, page_height: int = 15840) -> None:
    document = Document()
    section = document.sections[0]
    _letter_section(section)
    section.page_height = Twips(page_height)
    _columns(section, columns)
    section.header.is_linked_to_previous = False
    section.header.paragraphs[0].text = "Running head"
    section.footer.is_linked_to_previous = False
    section.footer.paragraphs[0].text = "Footer mark"
    document.styles["Normal"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    document.add_paragraph("A neutral sample paragraph for body typography.")
    document.save(path)


def _assemble(tmp_path: Path, text: str, **options):
    tmp_path.mkdir(parents=True, exist_ok=True)
    template = tmp_path / "template.docx"
    _template(template, columns=options.pop("columns_in_template", 2), page_height=options.pop("page_height", 15840))
    _metadata(tmp_path / "metadata.md")
    source = tmp_path / "input.md"
    source.write_text(text, encoding="utf-8")
    output = tmp_path / "output.docx"
    result = assemble_markdown_template(
        [source],
        template_path=template,
        output=output,
        metadata_path=tmp_path / "metadata.md",
        line_numbers="on",
        **options,
    )
    return Document(output), result, output


def _grid_widths(table) -> list[int]:
    return [int(node.get(qn("w:w"))) for node in table._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))]


def _printable(section) -> int:
    return int((section.page_width - section.left_margin - section.right_margin) / 635)


def _column_width(section) -> int:
    cols = section._sectPr.find(qn("w:cols"))
    count = int(cols.get(qn("w:num"), "1")) if cols is not None else 1
    gap = int(cols.get(qn("w:space"), "0")) if cols is not None and count > 1 else 0
    return (_printable(section) - gap * max(count - 1, 0)) // max(count, 1)


def _column_counts(document: Document) -> list[int]:
    counts = []
    for section in document.sections:
        cols = section._sectPr.find(qn("w:cols"))
        counts.append(int(cols.get(qn("w:num"), "1")) if cols is not None else 1)
    return counts


def _body_chunks(document: Document):
    body = document.element.body
    chunks = []
    current = []
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            chunks.append((current, child))
            current = []
            continue
        current.append(child)
        sect = child.find(qn("w:pPr") + "/" + qn("w:sectPr"))
        if sect is not None:
            chunks.append((current, sect))
            current = []
    if current:
        chunks.append((current, None))
    return chunks


def _chunk_text(nodes) -> str:
    return "\n".join("".join(node.itertext()) for node in nodes)


def _chunk_columns(sect) -> int:
    if sect is None:
        return 1
    cols = sect.find(qn("w:cols"))
    return int(cols.get(qn("w:num"), "1")) if cols is not None else 1


def _direct_alignments(table) -> list[list[str | None]]:
    rows = []
    for tr in table._tbl.findall(qn("w:tr")):
        row = []
        for tc in tr.findall(qn("w:tc")):
            paragraph = tc.find(qn("w:p"))
            jc = None if paragraph is None else paragraph.find(qn("w:pPr") + "/" + qn("w:jc"))
            row.append(None if jc is None else jc.get(qn("w:val")))
        rows.append(row)
    return rows


def _vmerges(table) -> list[list[str | None]]:
    rows = []
    for tr in table._tbl.findall(qn("w:tr")):
        row = []
        for tc in tr.findall(qn("w:tc")):
            marker = tc.find(qn("w:tcPr") + "/" + qn("w:vMerge"))
            row.append(None if marker is None else (marker.get(qn("w:val")) or "continue"))
        rows.append(row)
    return rows


def _raw_rows(table) -> list[list[str]]:
    rows = []
    for tr in table._tbl.findall(qn("w:tr")):
        row = []
        for tc in tr.findall(qn("w:tc")):
            row.append("".join(node.text or "" for node in tc.findall(".//" + qn("w:t"))))
        rows.append(row)
    return rows


MIXED_TABLE = """\
| Left | Center | Right | Inherit |
| :----- | :---: | ------: | --- |
| **alpha** | *beta* | gamma<sup>1</sup> | delta |
| short | a much longer centered value | x | tail |
"""


def test_cli_defaults_and_table_span_requires_template(tmp_path: Path) -> None:
    parser = build_parser()
    default = parser.parse_args(["md2docx", "source.md", "-o", "out.docx"])
    override = parser.parse_args([
        "md2docx", "source.md", "-o", "out.docx", "--table-span", "page", "--table-grouping", "off",
    ])
    assert default.table_span == "column"
    assert default.table_grouping == "auto"
    assert override.table_span == "page"
    assert override.table_grouping == "off"

    source = tmp_path / "input.md"
    source.write_text("| A |\n| --- |\n| 1 |\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    assert main(["md2docx", str(source), "-o", str(output), "--table-span", "page"]) == 2
    assert not output.exists()
    assert main(["md2docx", str(source), "-o", str(output), "--table-grouping", "off"]) == 0


def test_layout_directives_are_explicit_and_rejected_when_dangling(tmp_path: Path) -> None:
    path = tmp_path / "input.md"
    path.write_text(
        "Table span: page\nTable grouping: off\n"
        "Table: Measurements.|widths=proportional\n\n"
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n",
        encoding="utf-8",
    )
    blocks = parse_markdown(path)
    assert [block.kind for block in blocks] == ["table_caption", "table"]
    assert dict(blocks[0].options)["span"] == "page"
    assert dict(blocks[0].options)["grouping"] == "off"
    assert dict(blocks[0].options)["widths"] == "proportional"
    assert dict(blocks[1].options)["span"] == "page"
    assert dict(blocks[1].options)["grouping"] == "off"
    assert dict(blocks[1].options)["widths"] == "proportional"

    path.write_text("\\Table span: page\n\nParagraph.\n", encoding="utf-8")
    assert [block.text for block in parse_markdown(path)] == ["Table span: page", "Paragraph."]

    path.write_text("Table span: wide\n\n| A |\n| --- |\n| 1 |\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported table span"):
        parse_markdown(path)
    path.write_text("Table span: page\n\nNo table follows.\n", encoding="utf-8")
    with pytest.raises(ValueError, match="not followed by a table"):
        parse_markdown(path)
    path.write_text(
        "Table span: page\n\nTable: Caption.|span=column\n\n| A |\n| --- |\n| 1 |\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="Conflicting table span"):
        parse_markdown(path)


def test_mixed_alignment_markup_and_nonuniform_rows_standalone(tmp_path: Path) -> None:
    path = tmp_path / "input.md"
    path.write_text(MIXED_TABLE, encoding="utf-8")
    blocks = parse_markdown(path)
    assert blocks[0].alignments == ("left", "center", "right", "")
    assert blocks[0].rows[1][1] == "*beta*"
    assert len(blocks[0].rows[1][0]) < len(blocks[0].rows[2][1])
    document = render_blocks_to_doc(blocks)
    table = document.tables[0]
    expected = [
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.LEFT,
    ]
    for row in table.rows:
        assert [cell.paragraphs[0].alignment for cell in row.cells] == expected
    header = table.cell(0, 0).paragraphs[0]
    assert any(run.bold and run.text == "Left" for run in header.runs) or header.runs[0].bold
    body = table.cell(1, 0).paragraphs[0]
    assert any(run.bold and run.text == "alpha" for run in body.runs)
    center = table.cell(1, 1).paragraphs[0]
    assert any(run.italic and run.text == "beta" for run in center.runs)
    right = table.cell(1, 2).paragraphs[0]
    assert any(run.font.superscript and run.text == "1" for run in right.runs)
    raw = _raw_rows(table)
    assert all(len(row) == 4 for row in raw)
    assert raw[2][1] == "a much longer centered value"
    assert raw[2][0] == "short"


def test_template_alignment_does_not_override_unmarked_columns_and_round_trips(tmp_path: Path) -> None:
    document, _result, output = _assemble(tmp_path, MIXED_TABLE, columns_in_template=1)
    table = document.tables[0]
    assert _direct_alignments(table) == [
        ["left", "center", "right", None],
        ["left", "center", "right", None],
        ["left", "center", "right", None],
    ]
    for row in table.rows:
        assert [cell.paragraphs[0].alignment for cell in row.cells] == [
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            None,
        ]
        assert row.cells[3].paragraphs[0].style.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
    body = table.cell(1, 0).paragraphs[0]
    assert any(run.bold and run.text == "alpha" for run in body.runs)
    center = table.cell(1, 1).paragraphs[0]
    assert any(run.italic and run.text == "beta" for run in center.runs)
    right = table.cell(1, 2).paragraphs[0]
    assert any(run.font.superscript and run.text == "1" for run in right.runs)
    raw = _raw_rows(table)
    assert [len(row) for row in raw] == [4, 4, 4]
    assert len(raw[1][0]) != len(raw[2][1])

    semantic = DocxReader().read(output, options=ReadOptions()).document
    markdown = render_markdown(semantic)
    assert "| :--- | :---: | ---: | --- |" in markdown
    source = MarkdownReader().read(tmp_path / "input.md", options=ReadOptions()).document
    assert "| :--- | :---: | ---: | --- |" in render_markdown(source)


def test_column_count_does_not_select_page_span(tmp_path: Path) -> None:
    text = (
        "Body before.\n\nTable: Measurements.\n\n"
        "| A | B | C | D |\n| --- | --- | --- | --- |\n| 1 | 2 | 3 | 4 |\n\n"
        "Table note: Units are arbitrary.\n\nBody after.\n"
    )
    document, result, _output = _assemble(tmp_path, text)
    assert result.table_span == "column"
    assert result.tables[0]["span"] == "column"
    assert _column_counts(document) == [2]
    table = document.tables[0]
    assert sum(_grid_widths(table)) == 3960
    assert sum(_grid_widths(table)) == _column_width(document.sections[0])
    assert _printable(document.sections[0]) == 8640
    assert sum(_grid_widths(table)) != _printable(document.sections[0])
    assert table._tbl.find(qn("w:tblPr")).find(qn("w:tblLayout")).get(qn("w:type")) == "fixed"
    assert table._tbl.find(qn("w:tblPr")).find(qn("w:tblW")).get(qn("w:type")) == "dxa"


def test_page_span_keeps_caption_table_and_notes_then_restores_columns(tmp_path: Path) -> None:
    text = (
        "Body before.\n\n"
        "Table span: page\n"
        "Table: Measurements.|span=page\n\n"
        "| A | B | C | D |\n| --- | --- | --- | --- |\n| 1 | 2 | 3 | 4 |\n\n"
        "Table note: Units are arbitrary.\n\n"
        "Body after.\n"
    )
    document, result, output = _assemble(tmp_path, text)
    assert result.table_span == "column"
    assert result.tables[0]["span"] == "page"
    assert _column_counts(document) == [2, 1, 2]
    assert all(
        (section._sectPr.find(qn("w:type")) is None)
        or section._sectPr.find(qn("w:type")).get(qn("w:val")) == "continuous"
        for section in document.sections
    )
    assert '<w:br w:type="page"' not in document.element.xml
    table = document.tables[0]
    assert sum(_grid_widths(table)) == 8640
    assert sum(_grid_widths(table)) == _printable(document.sections[1])
    header = table._tbl.find(qn("w:tr")).find(qn("w:trPr")).find(qn("w:tblHeader"))
    assert header is not None
    for section in document.sections:
        assert section.header.paragraphs[0].text == "Running head"
        assert section.footer.paragraphs[0].text == "Footer mark"
        assert section._sectPr.find(qn("w:lnNumType")) is not None
    chunks = _body_chunks(document)
    table_index = next(index for index, (nodes, _sect) in enumerate(chunks) if any(node.tag == qn("w:tbl") for node in nodes))
    table_text = _chunk_text(chunks[table_index][0])
    assert "Measurements." in table_text
    assert "Units are arbitrary." in table_text
    assert _chunk_columns(chunks[table_index][1]) == 1
    assert "Body before." in _chunk_text(chunks[table_index - 1][0])
    assert _chunk_columns(chunks[table_index - 1][1]) == 2
    assert "Body after." in _chunk_text(chunks[table_index + 1][0])
    assert _chunk_columns(chunks[table_index + 1][1]) == 2
    assert output.is_file()
    assert result.figure_span == "column"


def test_page_span_default_override_and_single_column_has_no_extra_breaks(tmp_path: Path) -> None:
    narrow = (
        "Before.\n\nTable: Narrow.|span=column\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n"
        "Table span: page\nTable: Wide.\n\n| A | B | C | D |\n| --- | --- | --- | --- |\n| 1 | 2 | 3 | 4 |\n\n"
        "After.\n"
    )
    document, result, _output = _assemble(tmp_path, narrow, table_span="page")
    assert [table["span"] for table in result.tables] == ["column", "page"]
    assert _column_counts(document) == [2, 1, 2]
    assert sum(_grid_widths(document.tables[0])) == _column_width(document.sections[0])
    assert sum(_grid_widths(document.tables[1])) == _printable(document.sections[1])

    single, single_result, _output = _assemble(
        tmp_path / "one",
        "Before.\n\nTable span: page\n\n| A | B | C | D |\n| --- | --- | --- | --- |\n| 1 | 2 | 3 | 4 |\n\nAfter.\n",
        columns_in_template=1,
        table_span="column",
    )
    assert _column_counts(single) == [1]
    assert len(single.sections) == 1
    assert single_result.tables == ()
    assert sum(_grid_widths(single.tables[0])) == _printable(single.sections[0])


def test_fixed_and_proportional_widths_fill_the_page_span(tmp_path: Path) -> None:
    text = (
        "Table span: page\n"
        "Table: Widths.|widths={mode}\n\n"
        "| A | Broad-heading |\n| --- | --- |\n| z | alphabetical |\n"
    )
    fixed, _fixed_result, _output = _assemble(tmp_path / "fixed", text.format(mode="fixed"))
    proportional, _proportional_result, _output = _assemble(tmp_path / "proportional", text.format(mode="proportional"))
    fixed_widths = _grid_widths(fixed.tables[0])
    proportional_widths = _grid_widths(proportional.tables[0])
    assert sum(fixed_widths) == 8640
    assert sum(proportional_widths) == 8640
    assert abs(fixed_widths[0] - fixed_widths[1]) <= 1
    assert proportional_widths[1] > proportional_widths[0] * 2
    assert fixed.tables[0]._tbl.find(qn("w:tblPr")).find(qn("w:tblW")).get(qn("w:type")) == "dxa"
    assert proportional.tables[0]._tbl.find(qn("w:tblPr")).find(qn("w:tblLayout")).get(qn("w:type")) == "fixed"


def test_page_span_table_repeats_its_header_across_pages(tmp_path: Path) -> None:
    rows = "\n".join(f"| g{index} | {index} | value | tail |" for index in range(1, 28))
    text = (
        "Before the long table.\n\nTable span: page\nTable: Long.|grouping=off\n\n"
        "| Group | Index | Value | Tail |\n| --- | --- | --- | --- |\n"
        f"{rows}\n\nTable note: Continued.\n\nAfter the long table.\n"
    )
    document, _result, output = _assemble(tmp_path, text, page_height=5000)
    table = document.tables[0]
    assert sum(_grid_widths(table)) == _printable(document.sections[1])
    assert table._tbl.find(qn("w:tr")).find(qn("w:trPr")).find(qn("w:tblHeader")) is not None
    merges = _vmerges(table)
    assert all(marker is None for row in merges for marker in row)
    raw = _raw_rows(table)
    assert raw[1][0] == "g1"
    assert raw[-1][0] == "g27"
    middle = table._tbl.findall(qn("w:tr"))[10]
    assert middle.find(qn("w:trPr")).find(qn("w:tblHeader")) is None
    assert _column_counts(document) == [2, 1, 2]
    chunks = _body_chunks(document)
    table_index = next(index for index, (nodes, _sect) in enumerate(chunks) if any(node.tag == qn("w:tbl") for node in nodes))
    assert "Continued." in _chunk_text(chunks[table_index][0])
    assert "After the long table." in _chunk_text(chunks[table_index + 1][0])

    from docforge.docxmd.verify import _soffice
    from docforge.gates.render import layout_gate

    if _soffice() is None:
        pytest.skip("LibreOffice is unavailable. This skip is not a rendered-page pass.")
    fitz = pytest.importorskip("fitz")
    report = layout_gate(output, tmp_path / "render", "libreoffice")
    assert report["status"] == "passed"
    assert report["page_count"] is not None and report["page_count"] >= 2
    with fitz.open(tmp_path / "render" / "output.pdf") as pdf:
        text = "\n".join(page.get_text() for page in pdf)
    assert "g1" in text and "g27" in text
    assert "After the long table." in text


def test_rendered_page_span_is_wider_than_the_column(tmp_path: Path) -> None:
    from docforge.docxmd.verify import _soffice
    from docforge.gates.render import layout_gate

    if _soffice() is None:
        pytest.skip("LibreOffice is unavailable. This skip is not a rendered-page pass.")
    fitz = pytest.importorskip("fitz")
    text = (
        "Body before.\n\n{marker}"
        "Table: Measurements.\n\n"
        "| Left | Mid | EDGE |\n| --- | --- | --- |\n| a | b | z |\n\n"
        "Body after.\n"
    )

    def right_edge(directory: Path, marker: str) -> float:
        document, _result, output = _assemble(directory, text.format(marker=marker))
        assert document.tables
        report = layout_gate(output, directory / "render", "libreoffice")
        assert report["status"] == "passed"
        with fitz.open(directory / "render" / "output.pdf") as pdf:
            hits = []
            for page in pdf:
                for word in page.get_text("words"):
                    if word[4] == "EDGE":
                        hits.append(word[2])
            assert hits
            return max(hits)

    column_edge = right_edge(tmp_path / "column", "")
    page_edge = right_edge(tmp_path / "page", "Table span: page\n")
    assert page_edge > column_edge + 40


GROUPED = """\
| Value | Type | Group |
| --- | --- | --- |
| 1 | Alpha | A |
| 2 | Alpha | A |
| 3 | Beta | B |
| 3 | Beta | B |
"""


def test_grouping_off_keeps_every_cell_in_standalone_and_template(tmp_path: Path) -> None:
    path = tmp_path / "input.md"
    path.write_text(GROUPED, encoding="utf-8")
    blocks = parse_markdown(path)
    grouped = render_blocks_to_doc(blocks)
    separate = render_blocks_to_doc(blocks, table_grouping="off")
    grouped_merges = _vmerges(grouped.tables[0])
    assert grouped_merges[1][1] == "restart"
    assert grouped_merges[2][1] == "continue"
    assert grouped_merges[1][2] == "restart"
    assert grouped_merges[3][0] is None
    assert _raw_rows(grouped.tables[0])[2][1] == ""
    assert _raw_rows(grouped.tables[0])[1][0] == "1"
    assert _raw_rows(grouped.tables[0])[2][0] == "2"

    assert all(marker is None for row in _vmerges(separate.tables[0]) for marker in row)
    assert _raw_rows(separate.tables[0]) == [
        ["Value", "Type", "Group"],
        ["1", "Alpha", "A"],
        ["2", "Alpha", "A"],
        ["3", "Beta", "B"],
        ["3", "Beta", "B"],
    ]
    separate.save(tmp_path / "separate.docx")
    semantic = DocxReader().read(tmp_path / "separate.docx", options=ReadOptions()).document
    markdown = render_markdown(semantic)
    assert "| 1 | Alpha | A |" in markdown
    assert "| 2 | Alpha | A |" in markdown
    assert markdown.count("| 3 | Beta | B |") == 2

    template_doc, _result, template_output = _assemble(
        tmp_path / "template",
        "Table: Groups.|grouping=off\n\n" + GROUPED,
        columns_in_template=1,
    )
    assert all(marker is None for row in _vmerges(template_doc.tables[0]) for marker in row)
    assert _raw_rows(template_doc.tables[0])[2] == ["2", "Alpha", "A"]
    template_semantic = DocxReader().read(template_output, options=ReadOptions()).document
    template_markdown = render_markdown(template_semantic)
    assert "| 2 | Alpha | A |" in template_markdown


def test_grouping_modes_cover_blanks_repeats_and_per_table_overrides(tmp_path: Path) -> None:
    path = tmp_path / "input.md"
    path.write_text(
        "| Group | Value |\n| --- | --- |\n| A | 1 |\n|  | 2 |\n| A | 3 |\n\n"
        "| Sample | Y |\n| --- | --- |\n| s1 | 5 |\n| s2 | 5 |\n",
        encoding="utf-8",
    )
    blocks = parse_markdown(path)
    document = render_blocks_to_doc(blocks)
    assert all(marker is None for row in _vmerges(document.tables[0]) for marker in row)
    assert _raw_rows(document.tables[0])[1][0] == "A"
    assert _raw_rows(document.tables[0])[2][0] == ""
    assert _raw_rows(document.tables[0])[3][0] == "A"
    assert _raw_rows(document.tables[1]) == [["Sample", "Y"], ["s1", "5"], ["s2", "5"]]
    assert all(marker is None for row in _vmerges(document.tables[1]) for marker in row)

    path.write_text(
        "Table: First.|grouping=off\n\n| Group | Value |\n| --- | --- |\n| A | 1 |\n| A | 2 |\n\n"
        "Table grouping: auto\nTable: Second.\n\n| Group | Value |\n| --- | --- |\n| A | 1 |\n| A | 2 |\n",
        encoding="utf-8",
    )
    mixed = render_blocks_to_doc(parse_markdown(path), table_grouping="off")
    assert all(marker is None for row in _vmerges(mixed.tables[0]) for marker in row)
    assert _raw_rows(mixed.tables[0])[2] == ["A", "2"]
    assert _vmerges(mixed.tables[1])[1][0] == "restart"
    assert _vmerges(mixed.tables[1])[2][0] == "continue"
    assert _raw_rows(mixed.tables[1])[2][0] == ""
    assert _raw_rows(mixed.tables[1])[1][1] == "1"
    assert _raw_rows(mixed.tables[1])[2][1] == "2"


def test_multipage_table_keeps_authored_text_when_grouping_is_off(tmp_path: Path) -> None:
    text = "| Group | Value |\n| --- | --- |\n" + "\n".join(f"| A | {index} |" for index in range(1, 25)) + "\n"
    path = tmp_path / "input.md"
    path.write_text(text, encoding="utf-8")
    document = render_blocks_to_doc(parse_markdown(path), table_grouping="off")
    table = document.tables[0]
    assert table._tbl.find(qn("w:tr")).find(qn("w:trPr")).find(qn("w:tblHeader")) is not None
    assert all(marker is None for row in _vmerges(table) for marker in row)
    raw = _raw_rows(table)
    assert len(raw) == 25
    assert [row[0] for row in raw[1:]] == ["A"] * 24
    assert [row[1] for row in raw[1:]] == [str(index) for index in range(1, 25)]

    template_doc, _result, _output = _assemble(
        tmp_path / "template",
        "Table grouping: off\n\n" + text,
        columns_in_template=1,
        page_height=5000,
    )
    template_table = template_doc.tables[0]
    assert template_table._tbl.find(qn("w:tr")).find(qn("w:trPr")).find(qn("w:tblHeader")) is not None
    assert all(marker is None for row in _vmerges(template_table) for marker in row)
    template_raw = _raw_rows(template_table)
    assert [row[0] for row in template_raw[1:]] == ["A"] * 24
    assert template_raw[10][1] == "10"
