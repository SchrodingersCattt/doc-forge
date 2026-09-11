from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

from docforge.markdown import (
    assemble_markdown_template,
    discover_template_styles,
    parse_metadata,
    verify_template_output,
    write_assembly_sidecars,
)


def _template(path: Path) -> None:
    document = Document()
    for name in (
        "BB_Author_Name",
        "FA_Corresponding_Author_Footnote",
        "BD_Abstract_Title",
        "BD_Abstract",
        "TA_Main_Text",
        "TF_References_Section",
        "EndNote Bibliography",
    ):
        document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    document.add_paragraph("[TITLE] [placeholder text]").style = document.styles["BB_Author_Name"]
    document.add_paragraph("[ABSTRACT] [placeholder text]").style = document.styles["BD_Abstract"]
    document.add_section(WD_SECTION_START.CONTINUOUS)
    document.add_paragraph("[SECTION] [placeholder text]")
    document.save(path)


def _one_section_template(path: Path) -> None:
    document = Document()
    document.add_paragraph("[TITLE] [placeholder text]")
    document.save(path)


def _figure_template(path: Path, image: Path) -> None:
    document = Document()
    for name in ("BB_Author_Name", "BD_Abstract", "TA_Main_Text", "Caption", "EndNote Bibliography"):
        if name not in [style.name for style in document.styles]:
            document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    document.add_paragraph("[TITLE]").style = document.styles["BB_Author_Name"]
    document.add_paragraph("[ABSTRACT]").style = document.styles["BD_Abstract"]
    document.add_section(WD_SECTION_START.CONTINUOUS)
    document.add_paragraph("Body").style = document.styles["TA_Main_Text"]
    document.add_section(WD_SECTION_START.CONTINUOUS)
    document.add_paragraph().add_run().add_picture(str(image), width=1000000)
    document.add_paragraph("[FIGURE CAPTION] [placeholder text]").style = document.styles["Caption"]
    document.save(path)


def test_metadata_and_style_discovery(tmp_path: Path) -> None:
    metadata = tmp_path / "metadata.md"
    metadata.write_text(
        "# TITLE\n\nA title\n# AUTHOR\n\nA. Author\n# AFFILIATION:\n\nA Lab\n# EMAILS\n\n*one@example.com\n**two@example.com\n",
        encoding="utf-8",
    )
    parsed = parse_metadata(metadata)
    assert parsed.title == "A title"
    assert parsed.authors == "A. Author"
    assert parsed.affiliations == "A Lab"
    assert parsed.contacts == ("*one@example.com", "**two@example.com")

    template = tmp_path / "template.docx"
    _template(template)
    styles = discover_template_styles(Document(template))
    assert styles["title"] == "BB_Author_Name"
    assert styles["body"] == "TA_Main_Text"
    assert styles["reference"] == "EndNoteBibliography"


def test_template_assembly_replaces_placeholders_and_numbers_citations(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    abstract = tmp_path / "abstract.md"
    abstract.write_text("# Abstract\n\nAbstract cites \\citep{ref_b}.\n", encoding="utf-8")
    body = tmp_path / "body.md"
    body.write_text("# Introduction\n\nBody cites \\cite{ref_a}.\n", encoding="utf-8")
    bibliography = tmp_path / "refs.json"
    bibliography.write_text(
        json.dumps({"ref_a": "Author A. *Journal*, **2020**.", "ref_b": "Author B. *Journal*, **2021**."}),
        encoding="utf-8",
    )
    output = tmp_path / "output.docx"

    result = assemble_markdown_template(
        [abstract, body],
        template_path=template,
        output=output,
        metadata_path=metadata,
        bibliography_path=bibliography,
    )
    document = Document(output)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert result.used_citations == ("ref_a", "ref_b")
    assert "Abstract cites [2]." in text
    assert "Body cites [1]." in text
    assert "1.\tAuthor A." in text
    assert "2.\tAuthor B." in text
    assert "placeholder" not in text.lower()
    assert len(document.sections) == 2
    verify_template_output(
        output,
        expected_sections=2,
        expected_geometry=result.section_geometry,
    )
    manifest, checksum = write_assembly_sidecars(
        result,
        inputs=[abstract, body],
        template_path=template,
        metadata_path=metadata,
        bibliography_path=bibliography,
        command=["docforge", "md2docx"],
    )
    assert manifest.exists()
    assert checksum.exists()


def test_unknown_citation_fails_before_writing(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("# Intro\n\n\\citep{missing}\n", encoding="utf-8")
    bibliography = tmp_path / "refs.json"
    bibliography.write_text("{}", encoding="utf-8")
    output = tmp_path / "output.docx"
    with pytest.raises(ValueError, match="Unknown citation key"):
        assemble_markdown_template(
            [source],
            template_path=template,
            output=output,
            metadata_path=metadata,
            bibliography_path=bibliography,
        )
    assert not output.exists()


def test_one_section_template_uses_terminal_section_properties(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _one_section_template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("# Introduction\n\nBody.\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    result = assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata
    )
    assert result.template_sections_before == 1
    assert result.template_sections_after == 1
    assert "placeholder" not in "\n".join(p.text for p in Document(output).paragraphs).lower()


def test_template_assembly_keeps_inline_figure_and_caption_pair(tmp_path: Path) -> None:
    source_image = tmp_path / "source.png"
    Image.new("RGB", (80, 40), "white").save(source_image)
    template = tmp_path / "template.docx"
    _figure_template(template, source_image)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text(
        "# Introduction\n\nBody.\n\n![Figure 1|columns=double](source.png)\n\nFigure 1. Example inline figure.\n",
        encoding="utf-8",
    )
    output = tmp_path / "output.docx"
    result = assemble_markdown_template([source], template_path=template, output=output, metadata_path=metadata)
    document = Document(output)
    assert result.figures and result.figures[0]["caption"].startswith("Figure 1.")
    assert result.figures[0]["columns"] == 2
    assert len(document.inline_shapes) == 1
    assert any(p.text.startswith("Figure 1.") for p in document.paragraphs)
    verify_template_output(output, expected_sections=4)


def test_table_fits_explicit_one_column_section(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    document = Document(template)
    columns = document.sections[1]._sectPr.find(qn("w:cols"))
    if columns is None:
        columns = OxmlElement("w:cols")
        document.sections[1]._sectPr.append(columns)
    columns.set(qn("w:num"), "2")
    document.save(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("# Intro\n\n| One | Two |\n|---|---|\n| Alpha | Beta |\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    assemble_markdown_template([source], template_path=template, output=output, metadata_path=metadata, columns="one")
    result = Document(output)
    table = result.tables[0]._tbl
    width = sum(int(node.get(qn("w:w"))) for node in table.find(qn("w:tblGrid")))
    section = result.sections[1]
    expected = int((section.page_width - section.left_margin - section.right_margin) / 635)
    assert width == expected


def test_template_inherits_superscript_citations_without_private_markers(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    document = Document(template)
    document.add_paragraph().add_run("1").font.superscript = True
    document.save(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("# Intro\n\nFirst \\citep{ref}.\n\nSecond \\cite{ref}.\n\nThird \\citep{ref}.\n", encoding="utf-8")
    bibliography = tmp_path / "references.json"
    bibliography.write_text('{"ref": "Author. Journal. 2026."}', encoding="utf-8")
    output = tmp_path / "output.docx"
    result = assemble_markdown_template([source], template_path=template, output=output, metadata_path=metadata, bibliography_path=bibliography)
    document = Document(output)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert "\ue000" not in text and "\ue001" not in text
    assert result.verification["citation_count"] == 3
    assert sum(run.font.superscript is True for paragraph in document.paragraphs for run in paragraph.runs) == 3
    assert all(
        (section._sectPr.find(qn("w:type")) is None
         or section._sectPr.find(qn("w:type")).get(qn("w:val")) == "continuous")
        for section in document.sections
    )


def test_si_reuses_main_citation_numbers_and_prefixes_si_only_references(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    bibliography = tmp_path / "references.json"
    bibliography.write_text(
        '{"shared": "Shared. Journal. 2020.", "si_only": "SI only. Journal. 2021."}',
        encoding="utf-8",
    )
    main_source = tmp_path / "main.md"
    main_source.write_text("# Introduction\n\nMain \\citep{shared}.\n", encoding="utf-8")
    main_output = tmp_path / "main.docx"
    main_result = assemble_markdown_template(
        [main_source], template_path=template, output=main_output,
        metadata_path=metadata, bibliography_path=bibliography,
    )
    main_manifest, _ = write_assembly_sidecars(
        main_result, inputs=[main_source], template_path=template,
        metadata_path=metadata, bibliography_path=bibliography,
        command=["docforge", "md2docx"],
    )
    si_source = tmp_path / "si.md"
    si_source.write_text("# Supporting Information\n\nSI \\citep{shared,si_only}.\n", encoding="utf-8")
    si_output = tmp_path / "si.docx"
    si_result = assemble_markdown_template(
        [si_source], template_path=template, output=si_output,
        metadata_path=metadata, bibliography_path=bibliography,
        citation_base_path=main_manifest,
    )
    text = "\n".join(paragraph.text for paragraph in Document(si_output).paragraphs)
    assert si_result.citation_map == {"shared": 1, "si_only": "S1"}
    assert "\ue000" not in text and "[1,S1]" in text
    assert "S1.\tSI only." in text


def test_template_assembly_preserves_inline_runs_and_numbers_display_equation(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text(
        "# Introduction\n\nInline $t_{\\mathrm{chem}}$ uses runs.\n\n"
        "$$\nE = mc^2\n$$\n",
        encoding="utf-8",
    )
    output = tmp_path / "output.docx"
    assemble_markdown_template([source], template_path=template, output=output, metadata_path=metadata)
    document = Document(output)
    inline = next(p for p in document.paragraphs if "Inline" in p.text)
    equation = next(p for p in document.paragraphs if "(1)" in p.text)
    assert "oMath" not in inline._p.xml
    assert any(run.text == "chem" and run.font.subscript for run in inline.runs)
    assert "oMathPara" in equation._p.xml
    assert "(1)" in equation.text
    tab = equation._p.find(".//" + qn("w:tab"))
    assert tab is not None and int(tab.get(qn("w:pos"))) > 0


def test_no_title_strips_level_one_and_generated_reference_heading(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text(
        "# Introduction\n\nBody \\citep{ref}.\n\n## Detail\n\nMore.\n",
        encoding="utf-8",
    )
    bibliography = tmp_path / "references.json"
    bibliography.write_text('{"ref": "Author. Journal. 2026."}', encoding="utf-8")
    output = tmp_path / "output.docx"
    result = assemble_markdown_template(
        [source],
        template_path=template,
        output=output,
        metadata_path=metadata,
        bibliography_path=bibliography,
        strip_level_one_headings=True,
    )
    text = "\n".join(paragraph.text for paragraph in Document(output).paragraphs)
    assert result.strip_level_one_headings is True
    assert "A title" in text
    assert "Body [1]." in text
    assert "Detail" in text
    assert "Introduction" not in text
    assert "References" not in text
    assert "1.\tAuthor." in text


def test_unfilled_optional_metadata_is_omitted(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n\n# Acknowledgement\n\nTODO\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("# Introduction\n\nBody.\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata
    )
    text = "\n".join(paragraph.text for paragraph in Document(output).paragraphs)
    assert "TODO" not in text
    assert "ACKNOWLEDGMENT" not in text

def test_template_rendering_args_override_font_style_and_line_numbers(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _template(template)
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    style = tmp_path / "styles.json"
    style.write_text('{"body": "TA_Main_Text"}', encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("# Introduction\n\nBody.\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    result = assemble_markdown_template(
        [source],
        template_path=template,
        output=output,
        metadata_path=metadata,
        font_family="Arial",
        east_asia_font="SimSun",
        style_profile=str(style),
        line_numbers="on",
        strip_level_one_headings=True,
        heading_before=6,
    )
    document = Document(output)
    assert result.style_profile == str(style)
    assert result.font_family == "Arial"
    assert result.east_asia_font == "SimSun"
    assert result.line_numbers == "on"
    assert result.include_title is True
    assert result.strip_level_one_headings is True
    assert result.heading_before == 6
    assert all(section._sectPr.find(qn("w:lnNumType")) is not None for section in document.sections)
    text = "\n".join(p.text for p in document.paragraphs)
    assert "A title" in text
    assert "Introduction" not in text
    text_runs = [run for paragraph in document.paragraphs for run in paragraph.runs if run.text]
    assert text_runs
    for run in text_runs:
        fonts = run._r.rPr.rFonts
        assert fonts.get(qn("w:ascii")) == "Arial"
        assert fonts.get(qn("w:hAnsi")) == "Arial"
        assert fonts.get(qn("w:eastAsia")) == "SimSun"
