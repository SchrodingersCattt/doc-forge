"""Cover and caption typography acceptance tests.

Fixtures are synthetic. Cover prototype preservation and caption prototype
preservation stay independent, and named styles remain the default authority.
"""

from __future__ import annotations

import json
import warnings
from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from PIL import Image

from docforge.cli import main
from docforge.markdown import Block, render_blocks_to_doc
from docforge.markdown.template import assemble_markdown_template, write_assembly_sidecars


def _cover_template(path: Path, *, named_direct: bool = False) -> None:
    """Build the cover fixture.

    The default shape matches a Normal cover whose visible formatting is
    direct. ``named_direct`` puts that formatting on paragraphs that already
    use explicit role styles, while the style definitions stay plain.
    """
    document = Document()
    names = ["P1", "FigureCaption", "References", "Authors", "Address"]
    if named_direct:
        names.append("Title1")
    for name in names:
        if name not in document.styles:
            document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if named_direct:
        document.styles["Title1"].font.size = Pt(11)
        document.styles["Title1"].font.color.rgb = RGBColor(0x17, 0x36, 0x5D)
        document.styles["Authors"].font.size = Pt(11)
        document.styles["Address"].font.size = Pt(11)
        document.styles["P1"].font.size = Pt(12)
    samples = (
        ("Supporting Information", "P1" if named_direct else None, 12, False, True),
        ("Article Title", "Title1" if named_direct else None, 16, True, False),
        ("Author One", "Authors" if named_direct else None, 12, False, False),
        ("Affiliation", "Address" if named_direct else None, 12, False, True),
        ("author@example.org", "Address" if named_direct else None, 12, False, False),
    )
    for text, style_name, size, bold, italic in samples:
        paragraph = document.add_paragraph(style=style_name) if style_name else document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
        if named_direct and text == "Article Title":
            run.font.color.rgb = RGBColor(0, 0, 0)
    document.add_paragraph("Methods", style="Heading 1")
    document.add_paragraph("Example text.", style="P1")
    noise = document.add_paragraph("Unrelated sample formatting.")
    noise.alignment = WD_ALIGN_PARAGRAPH.CENTER
    noise_run = noise.runs[0]
    noise_run.bold = True
    noise_run.font.size = Pt(28)
    noise_run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    document.add_paragraph("Figure 1. Example.", style="FigureCaption")
    document.add_paragraph("References", style="Heading 1")
    document.add_paragraph("Example citation.", style="References")
    document.save(path)


def _named_role_template(path: Path) -> None:
    """Role styles themselves carry the cover formatting."""
    document = Document()
    specs = (
        ("Title1", 16, True, False, RGBColor(0, 0, 0)),
        ("Authors", 12, False, False, RGBColor(0, 0, 0)),
        ("Address", 12, False, True, RGBColor(0, 0, 0)),
        ("P1", 12, False, False, RGBColor(0, 0, 0)),
    )
    for name, size, bold, italic, color in specs:
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.italic = italic
        style.font.color.rgb = color
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if name != "P1" else WD_ALIGN_PARAGRAPH.LEFT
    document.add_paragraph("Article Title", style="Title1")
    document.add_paragraph("Author One", style="Authors")
    document.add_paragraph("Affiliation", style="Address")
    document.add_paragraph("Methods", style="Heading 1")
    document.add_paragraph("Example text.", style="P1")
    document.save(path)


def _write_cover_sources(root: Path) -> tuple[Path, Path]:
    metadata = root / "metadata.md"
    metadata.write_text(
        "# TITLE\nSample study\n\n# AUTHOR\nA. Researcher\n\n"
        "# AFFILIATION\nExample Institute\n\n# EMAILS\nauthor@example.org\n",
        encoding="utf-8",
    )
    source = root / "input.md"
    source.write_text("# Methods\n\nA sample paragraph.\n", encoding="utf-8")
    return source, metadata


def _cover_paragraphs(path: Path) -> list:
    return [paragraph for paragraph in Document(path).paragraphs if paragraph.text][:6]


def _assert_no_cover_warning(callback):
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        result = callback()
    assert not [item for item in caught if "Cover sample" in str(item.message)]
    return result


def test_legacy_level_one_heading_stays_heading_2() -> None:
    document = render_blocks_to_doc([Block("heading", "Methods", level=1), Block("paragraph", "Body.")])
    assert document.paragraphs[0].style.name == "Heading 2"
    assert document.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.JUSTIFY


def test_unused_title_is_reported_until_cover_prototypes_are_requested(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _cover_template(template)
    source, metadata = _write_cover_sources(tmp_path)
    output = tmp_path / "default.docx"
    with pytest.warns(UserWarning, match="Cover sample formatting conflicts"):
        result = assemble_markdown_template(
            [source], template_path=template, output=output, metadata_path=metadata,
        )
    title = next(paragraph for paragraph in Document(output).paragraphs if paragraph.text == "Sample study")
    assert title.style.name == "Title"
    assert title.alignment != WD_ALIGN_PARAGRAPH.CENTER
    roles = [item["role"] for item in result.metadata_style_conflicts]
    assert roles == ["supporting_information", "title", "authors", "affiliations", "contacts"]
    title_conflict = next(item for item in result.metadata_style_conflicts if item["role"] == "title")
    assert title_conflict["selected_style"] == "Title"
    assert title_conflict["prototype_style"] == "Normal"
    assert title_conflict["selected_style_unused"] is True
    assert title_conflict["override"] == "--preserve-metadata-prototypes"
    assert result.verification["metadata_style_conflicts"] == [dict(item) for item in result.metadata_style_conflicts]

    cli_output = tmp_path / "cli-default.docx"
    with pytest.warns(UserWarning, match="title -> Title"):
        assert main([
            "md2docx", str(source), "-o", str(cli_output), "--template", str(template),
            "--metadata", str(metadata), "--style", "template",
        ]) == 0
    manifest = json.loads(cli_output.with_suffix(".manifest.json").read_text(encoding="utf-8"))
    assert manifest["rendering"]["preserve_metadata_prototypes"] is False
    assert [item["role"] for item in manifest["rendering"]["metadata_style_conflicts"]] == roles
    assert manifest["verification"]["metadata_style_conflicts"][1]["selected_style"] == "Title"

    preserved = tmp_path / "preserved.docx"
    preserved_result = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=preserved, metadata_path=metadata,
        preserve_metadata_prototypes=True,
    ))
    cover = _cover_paragraphs(preserved)[:5]
    assert [paragraph.text for paragraph in cover] == [
        "Supporting Information", "Sample study", "A. Researcher",
        "Example Institute", "author@example.org",
    ]
    assert all(paragraph.style.name == "Normal" for paragraph in cover)
    assert all(paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER for paragraph in cover)
    assert [paragraph.runs[0].font.size.pt for paragraph in cover] == [12, 16, 12, 12, 12]
    assert [paragraph.runs[0].italic for paragraph in cover] == [True, False, False, True, False]
    assert cover[1].runs[0].bold is True
    assert cover[2].runs[0].bold is not True
    assert cover[1].runs[0].font.color.rgb is None or str(cover[1].runs[0].font.color.rgb) == "000000"
    body = next(paragraph for paragraph in Document(preserved).paragraphs if paragraph.text == "A sample paragraph.")
    assert body.alignment != WD_ALIGN_PARAGRAPH.CENTER
    assert all(run.font.size is None or run.font.size.pt != 28 for run in body.runs)
    assert preserved_result.metadata_style_conflicts == ()
    assert preserved_result.preserve_metadata_prototypes is True
    assert preserved_result.preserve_caption_prototypes is False


def test_explicit_style_profile_is_the_cover_override(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _cover_template(template)
    source, metadata = _write_cover_sources(tmp_path)
    profile = tmp_path / "styles.json"
    profile.write_text(json.dumps({
        "title": "Title", "authors": "Authors", "affiliations": "Address", "body": "P1",
    }), encoding="utf-8")
    output = tmp_path / "mapped.docx"
    result = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata,
        style_profile=str(profile),
    ))
    title = next(paragraph for paragraph in Document(output).paragraphs if paragraph.text == "Sample study")
    assert title.style.name == "Title"
    assert result.metadata_style_conflicts == ()

    partial = tmp_path / "title-only.json"
    partial.write_text(json.dumps({"title": "Title"}), encoding="utf-8")
    partial_output = tmp_path / "partial.docx"
    with pytest.warns(UserWarning, match="authors -> Authors"):
        partial_result = assemble_markdown_template(
            [source], template_path=template, output=partial_output, metadata_path=metadata,
            style_profile=str(partial),
        )
    assert "title" not in [item["role"] for item in partial_result.metadata_style_conflicts]
    assert "authors" in [item["role"] for item in partial_result.metadata_style_conflicts]


def test_used_named_cover_style_is_not_a_conflict(tmp_path: Path) -> None:
    document = Document()
    document.styles.add_style("BBAuthorName", WD_STYLE_TYPE.PARAGRAPH)
    document.styles.add_style("Address", WD_STYLE_TYPE.PARAGRAPH)
    document.styles.add_style("P1", WD_STYLE_TYPE.PARAGRAPH)
    document.add_paragraph("Template title", style="BBAuthorName")
    document.add_paragraph("Body sample.", style="P1")
    template = tmp_path / "template.docx"
    document.save(template)
    source, metadata = _write_cover_sources(tmp_path)
    output = tmp_path / "output.docx"
    result = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata,
    ))
    assert result.metadata_style_conflicts == ()
    title = next(paragraph for paragraph in Document(output).paragraphs if paragraph.text == "Sample study")
    assert title.style.name == "BBAuthorName"


def test_named_role_styles_keep_their_own_formatting(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _named_role_template(template)
    source, metadata = _write_cover_sources(tmp_path)
    output = tmp_path / "output.docx"
    result = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata,
    ))
    assert result.metadata_style_conflicts == ()
    rendered = Document(output)
    title = next(paragraph for paragraph in rendered.paragraphs if paragraph.text == "Sample study")
    author = next(paragraph for paragraph in rendered.paragraphs if paragraph.text == "A. Researcher")
    affiliation = next(paragraph for paragraph in rendered.paragraphs if paragraph.text == "Example Institute")
    assert title.style.name == "Title1"
    assert author.style.name == "Authors"
    assert affiliation.style.name == "Address"
    assert rendered.styles["Title1"].font.size.pt == 16
    assert rendered.styles["Title1"].font.bold is True
    assert str(rendered.styles["Title1"].font.color.rgb) == "000000"
    assert rendered.styles["Title1"].paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert rendered.styles["Authors"].font.size.pt == 12
    assert rendered.styles["Address"].font.italic is True
    assert rendered.styles["Address"].paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    body = next(paragraph for paragraph in rendered.paragraphs if paragraph.text == "A sample paragraph.")
    assert body.style.name == "P1"
    assert body.alignment != WD_ALIGN_PARAGRAPH.CENTER


def test_named_role_direct_formatting_is_opt_in(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _cover_template(template, named_direct=True)
    source, metadata = _write_cover_sources(tmp_path)
    default_output = tmp_path / "default.docx"
    preserved_output = tmp_path / "preserved.docx"
    default = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=default_output, metadata_path=metadata,
    ))
    preserved = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=preserved_output, metadata_path=metadata,
        preserve_metadata_prototypes=True,
    ))
    assert default.metadata_style_conflicts == ()
    assert preserved.metadata_style_conflicts == ()
    default_title = next(paragraph for paragraph in Document(default_output).paragraphs if paragraph.text == "Sample study")
    preserved_title = next(paragraph for paragraph in Document(preserved_output).paragraphs if paragraph.text == "Sample study")
    assert default_title.style.name == "Title1"
    assert default_title.alignment != WD_ALIGN_PARAGRAPH.CENTER
    assert default_title.runs[0].font.size is None
    assert default_title.runs[0].font.color.rgb is None
    assert preserved_title.style.name == "Title1"
    assert preserved_title.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert preserved_title.runs[0].font.size.pt == 16
    assert preserved_title.runs[0].bold is True
    assert str(preserved_title.runs[0].font.color.rgb) == "000000"
    preserved_doc = Document(preserved_output)
    author = next(paragraph for paragraph in preserved_doc.paragraphs if paragraph.text == "A. Researcher")
    affiliation = next(paragraph for paragraph in preserved_doc.paragraphs if paragraph.text == "Example Institute")
    email = next(paragraph for paragraph in preserved_doc.paragraphs if paragraph.text == "author@example.org")
    banner = next(paragraph for paragraph in preserved_doc.paragraphs if paragraph.text == "Supporting Information")
    assert author.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert author.runs[0].font.size.pt == 12
    assert affiliation.runs[0].italic is True
    assert affiliation.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert email.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert email.runs[0].italic is not True
    assert email.runs[0].font.size.pt == 12
    assert banner.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert banner.runs[0].italic is True
    body = next(paragraph for paragraph in preserved_doc.paragraphs if paragraph.text == "A sample paragraph.")
    assert body.alignment != WD_ALIGN_PARAGRAPH.CENTER
    assert all(run.font.size is None or run.font.size.pt != 28 for run in body.runs)
    assert all(run.font.color.rgb is None or str(run.font.color.rgb) != "C00000" for run in body.runs)


def _caption_template(path: Path, image: Path, *, figure: bool = True, table: bool = True) -> None:
    document = Document()
    document.styles["Caption"].font.color.rgb = RGBColor(0x00, 0x44, 0xBB)
    document.add_paragraph("A neutral sample paragraph for body typography.")
    document.add_paragraph().add_run().add_picture(str(image), width=1000000)
    if figure:
        paragraph = document.add_paragraph(style="Normal")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.paragraph_format.line_spacing = 1.08
        paragraph.add_run("Figure 1.").bold = True
        run = paragraph.add_run(" A representative caption sample.")
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.name = "Calibri"
    if table:
        paragraph = document.add_paragraph(style="Normal")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(3)
        paragraph.paragraph_format.space_after = Pt(8)
        paragraph.add_run("Table 1.").bold = True
        run = paragraph.add_run(" A representative table caption.")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 0, 0)
    document.save(path)


def _caption_source(root: Path, image: Path) -> Path:
    source = root / "input.md"
    source.write_text(
        "Table: A *measured* H<sub>2</sub>O value.\n\n"
        "| A | B |\n|---|---|\n| 1 | 2 |\n\n"
        f"![Figure 1]({image.name})\n\nFigure 1. A *measured* x<sup>2</sup> response.\n",
        encoding="utf-8",
    )
    return source


def _paragraph(path: Path, prefix: str):
    return next(paragraph for paragraph in Document(path).paragraphs if paragraph.text.startswith(prefix))


def test_caption_prototypes_keep_paragraph_metrics_and_authored_emphasis(tmp_path: Path) -> None:
    image = tmp_path / "figure.png"
    Image.new("RGB", (80, 40), "white").save(image)
    template = tmp_path / "template.docx"
    _caption_template(template, image)
    source = _caption_source(tmp_path, image)
    default_output = tmp_path / "default.docx"
    preserved_output = tmp_path / "preserved.docx"
    assemble_markdown_template([source], template_path=template, output=default_output, title="Sample study")
    result = assemble_markdown_template(
        [source], template_path=template, output=preserved_output, title="Sample study",
        preserve_caption_prototypes=True,
    )
    default_figure = _paragraph(default_output, "Figure 1.")
    figure = _paragraph(preserved_output, "Figure 1.")
    table = _paragraph(preserved_output, "Table 1.")
    assert default_figure.style.name == "Caption"
    assert figure.style.name == "Normal"
    assert figure.alignment == WD_ALIGN_PARAGRAPH.LEFT
    assert figure.paragraph_format.first_line_indent.pt == 0
    assert figure.paragraph_format.space_before.pt == 6
    assert figure.paragraph_format.space_after.pt == 2
    assert figure.paragraph_format.line_spacing == pytest.approx(1.08, abs=0.02)
    assert figure.runs[0].bold is True
    assert next(run for run in figure.runs if run.text == "measured").italic is True
    assert next(run for run in figure.runs if run.text == "2").font.superscript is True
    ordinary = next(run for run in figure.runs if "response" in run.text)
    assert ordinary.font.size.pt == 10
    assert str(ordinary.font.color.rgb) == "000000"
    assert ordinary.font.name == "Calibri"
    assert table.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert table.paragraph_format.space_before.pt == 3
    assert table.paragraph_format.space_after.pt == 8
    assert next(run for run in table.runs if "value" in run.text).font.size.pt == 9
    assert next(run for run in table.runs if run.text == "measured").italic is True
    assert next(run for run in table.runs if run.text == "2").font.subscript is True
    assert result.caption_prototype_roles == ("caption", "table_caption")
    assert result.preserve_metadata_prototypes is False
    manifest, _ = write_assembly_sidecars(
        result, inputs=[source], template_path=template, metadata_path=None,
        bibliography_path=None, command=["docforge", "md2docx"],
    )
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    assert payload["rendering"]["preserve_caption_prototypes"] is True
    assert payload["rendering"]["caption_prototype_roles"] == ["caption", "table_caption"]
    assert payload["verification"]["caption_prototype_roles"] == ["caption", "table_caption"]

    sized = tmp_path / "sized.docx"
    assemble_markdown_template(
        [source], template_path=template, output=sized, title="Sample study",
        preserve_caption_prototypes=True, caption_font_size=8, font_family="Georgia",
    )
    sized_figure = _paragraph(sized, "Figure 1.")
    assert all(run.font.size.pt == 8 for run in sized_figure.runs if run.text)
    assert all(run.font.name == "Georgia" for run in sized_figure.runs if run.text)
    assert next(run for run in sized_figure.runs if run.text == "2").font.superscript is True


def test_caption_samples_fall_back_independently(tmp_path: Path) -> None:
    image = tmp_path / "figure.png"
    Image.new("RGB", (80, 40), "white").save(image)
    source = _caption_source(tmp_path, image)
    figure_only = tmp_path / "figure-only.docx"
    _caption_template(figure_only, image, table=False)
    figure_output = tmp_path / "figure-output.docx"
    figure_result = assemble_markdown_template(
        [source], template_path=figure_only, output=figure_output, title="Sample study",
        preserve_caption_prototypes=True,
    )
    assert _paragraph(figure_output, "Figure 1.").style.name == "Normal"
    assert _paragraph(figure_output, "Table 1.").style.name == "Caption"
    assert figure_result.caption_prototype_roles == ("caption",)

    table_only = tmp_path / "table-only.docx"
    _caption_template(table_only, image, figure=False)
    table_output = tmp_path / "table-output.docx"
    table_result = assemble_markdown_template(
        [source], template_path=table_only, output=table_output, title="Sample study",
        preserve_caption_prototypes=True,
    )
    assert _paragraph(table_output, "Table 1.").style.name == "Normal"
    assert _paragraph(table_output, "Figure 1.").style.name == "Caption"
    assert table_result.caption_prototype_roles == ("table_caption",)

    profile = tmp_path / "styles.json"
    profile.write_text(json.dumps({"caption": "Caption", "table_caption": "Caption"}), encoding="utf-8")
    both = tmp_path / "both.docx"
    _caption_template(both, image)
    mapped = tmp_path / "mapped.docx"
    mapped_result = assemble_markdown_template(
        [source], template_path=both, output=mapped, title="Sample study",
        style_profile=str(profile), preserve_caption_prototypes=True,
    )
    assert _paragraph(mapped, "Figure 1.").style.name == "Caption"
    assert _paragraph(mapped, "Table 1.").style.name == "Caption"
    assert mapped_result.caption_prototype_roles == ()


def test_cover_and_caption_prototype_flags_do_not_change_each_other(tmp_path: Path) -> None:
    image = tmp_path / "figure.png"
    Image.new("RGB", (80, 40), "white").save(image)
    document = Document()
    for name in ("P1", "References", "Authors", "Address"):
        document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    document.styles["Caption"].font.color.rgb = RGBColor(0x00, 0x44, 0xBB)
    for text, size, bold, italic in (
        ("Supporting Information", 12, False, True),
        ("Article Title", 16, True, False),
        ("Author One", 12, False, False),
        ("Affiliation", 12, False, True),
        ("author@example.org", 12, False, False),
    ):
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
    document.add_paragraph("Methods", style="Heading 1")
    document.add_paragraph("Example text.", style="P1")
    document.add_paragraph().add_run().add_picture(str(image), width=1000000)
    sample = document.add_paragraph(style="Normal")
    sample.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sample.add_run("Figure 1.").bold = True
    caption_run = sample.add_run(" A representative caption sample.")
    caption_run.font.size = Pt(10)
    caption_run.font.color.rgb = RGBColor(0, 0, 0)
    template = tmp_path / "template.docx"
    document.save(template)
    source, metadata = _write_cover_sources(tmp_path)
    source.write_text(
        "# Methods\n\nA sample paragraph.\n\n"
        f"![Figure 1]({image.name})\n\nFigure 1. Example caption.\n",
        encoding="utf-8",
    )
    cover_output = tmp_path / "cover.docx"
    cover = _assert_no_cover_warning(lambda: assemble_markdown_template(
        [source], template_path=template, output=cover_output, metadata_path=metadata,
        preserve_metadata_prototypes=True,
    ))
    assert cover.caption_prototype_roles == ()
    assert _paragraph(cover_output, "Sample study").alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert _paragraph(cover_output, "Figure 1.").style.name == "Caption"

    caption_output = tmp_path / "caption.docx"
    with pytest.warns(UserWarning, match="title -> Title"):
        caption = assemble_markdown_template(
            [source], template_path=template, output=caption_output, metadata_path=metadata,
            preserve_caption_prototypes=True,
        )
    assert _paragraph(caption_output, "Sample study").style.name == "Title"
    assert _paragraph(caption_output, "Figure 1.").style.name == "Normal"
    assert caption.preserve_metadata_prototypes is False
    assert caption.caption_prototype_roles == ("caption",)


def _pdf_line(pdf: Path, needle: str) -> dict[str, object]:
    pymupdf = pytest.importorskip("pymupdf")
    with pymupdf.open(pdf) as document:
        page = document[0]
        width = float(page.rect.width)
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                spans = [span for span in line.get("spans", []) if span.get("text", "").strip()]
                text = "".join(span.get("text", "") for span in spans)
                if needle not in text:
                    continue
                box_left = min(span["bbox"][0] for span in spans)
                box_right = max(span["bbox"][2] for span in spans)
                return {
                    "text": text,
                    "size": spans[0]["size"],
                    "color": spans[0]["color"],
                    "flags": spans[0].get("flags", 0),
                    "center": (box_left + box_right) / 2,
                    "width": width,
                }
    raise AssertionError(f"{needle!r} was not rendered in {pdf}")


def _render(docx: Path, directory: Path) -> Path:
    from docforge.docxmd.verify import _soffice, render_pdf

    if _soffice() is None:
        pytest.skip("LibreOffice is unavailable. This skip is not a visual pass.")
    return render_pdf(docx, directory, "libreoffice")


def _assert_centered(metrics: dict[str, object], label: str) -> None:
    center = float(metrics["center"])
    width = float(metrics["width"])
    assert abs(center - width / 2) < 72, (label, metrics)


def test_rendered_normal_cover_preserves_visible_typography(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _cover_template(template)
    source, metadata = _write_cover_sources(tmp_path)
    preserved = tmp_path / "preserved.docx"
    default = tmp_path / "default.docx"
    with pytest.warns(UserWarning, match="title -> Title"):
        assemble_markdown_template(
            [source], template_path=template, output=default, metadata_path=metadata,
        )
    assemble_markdown_template(
        [source], template_path=template, output=preserved, metadata_path=metadata,
        preserve_metadata_prototypes=True,
    )
    preserved_pdf = _render(preserved, tmp_path / "preserved-pdf")
    default_pdf = _render(default, tmp_path / "default-pdf")
    title = _pdf_line(preserved_pdf, "Sample study")
    banner = _pdf_line(preserved_pdf, "Supporting Information")
    author = _pdf_line(preserved_pdf, "A. Researcher")
    affiliation = _pdf_line(preserved_pdf, "Example Institute")
    email = _pdf_line(preserved_pdf, "author@example.org")
    assert 14 <= float(title["size"]) <= 18
    assert title["color"] == 0
    _assert_centered(title, "title")
    assert 10 <= float(banner["size"]) <= 13
    assert banner["color"] == 0
    _assert_centered(banner, "banner")
    assert int(banner["flags"]) & 2
    _assert_centered(author, "author")
    assert not int(author["flags"]) & 2
    _assert_centered(affiliation, "affiliation")
    assert int(affiliation["flags"]) & 2
    _assert_centered(email, "email")
    default_title = _pdf_line(default_pdf, "Sample study")
    assert float(default_title["size"]) >= 22
    assert default_title["color"] != 0
    assert abs(float(default_title["center"]) - float(default_title["width"]) / 2) > 72


def test_rendered_named_role_cover_uses_style_metrics(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _named_role_template(template)
    source, metadata = _write_cover_sources(tmp_path)
    output = tmp_path / "named.docx"
    assemble_markdown_template(
        [source], template_path=template, output=output, metadata_path=metadata,
    )
    pdf = _render(output, tmp_path / "pdf")
    title = _pdf_line(pdf, "Sample study")
    author = _pdf_line(pdf, "A. Researcher")
    affiliation = _pdf_line(pdf, "Example Institute")
    assert 14 <= float(title["size"]) <= 18
    assert title["color"] == 0
    _assert_centered(title, "named title")
    assert 10 <= float(author["size"]) <= 13
    _assert_centered(author, "named author")
    _assert_centered(affiliation, "named affiliation")
    assert int(affiliation["flags"]) & 2


def test_rendered_caption_sample_stays_black(tmp_path: Path) -> None:
    image = tmp_path / "figure.png"
    Image.new("RGB", (80, 40), "white").save(image)
    template = tmp_path / "template.docx"
    _caption_template(template, image, table=False)
    source = tmp_path / "input.md"
    source.write_text(f"![Figure 1]({image.name})\n\nFigure 1. Example caption.\n", encoding="utf-8")
    preserved = tmp_path / "preserved.docx"
    default = tmp_path / "default.docx"
    assemble_markdown_template([source], template_path=template, output=default, title="Sample study")
    assemble_markdown_template(
        [source], template_path=template, output=preserved, title="Sample study",
        preserve_caption_prototypes=True,
    )
    preserved_line = _pdf_line(_render(preserved, tmp_path / "preserved-pdf"), "Example caption")
    default_line = _pdf_line(_render(default, tmp_path / "default-pdf"), "Example caption")
    assert 8 <= float(preserved_line["size"]) <= 12
    assert preserved_line["color"] == 0
    assert default_line["color"] != 0
