from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

from docforge import _pandoc
from docforge.docxdiff import create_tracked_docx
from docforge.docxdiff.redline import _accepted_revision_view, _final_blocks, visible_text
from docforge.markdown import (
    Block,
    docx_to_markdown,
    normalize_html_tables,
    normalize_inline_html,
    normalize_script_boundaries,
    render_blocks_to_doc,
    reuse_unchanged_roundtrip_source,
    split_markdown_sections,
)
from docforge.cli import main as cli_main
from docforge.tex import docx_to_tex


PANDOC_AVAILABLE = shutil.which("pandoc") is not None


def _sample_docx(path: Path, image: Path) -> None:
    document = Document()
    document.add_heading("Title", level=1)
    document.add_paragraph("Author")
    document.add_heading("Abstract", level=1)
    paragraph = document.add_paragraph("H")
    superscript = paragraph.add_run("2")
    superscript.font.superscript = True
    document.add_heading("Main", level=1)
    document.add_paragraph("Body")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    document.add_picture(str(image))
    document.add_heading("References", level=1)
    document.add_paragraph("1. Example reference")
    document.add_heading("Methods", level=1)
    document.add_paragraph("Method")
    document.add_heading("References", level=1)
    document.add_paragraph("51. Method reference")
    document.add_heading("Data availability", level=1)
    document.add_paragraph("Data")
    document.save(path)


@pytest.mark.skipif(not PANDOC_AVAILABLE, reason="Pandoc is required for integration tests")
def test_docx_to_markdown_and_split(tmp_path: Path) -> None:
    image = tmp_path / "source.png"
    Image.new("RGB", (16, 8), "white").save(image)
    source = tmp_path / "source.docx"
    _sample_docx(source, image)
    split_dir = tmp_path / "sections"
    section_map = tmp_path / "section-map.json"
    section_map.write_text(
        json.dumps(
            {
                "sections": [
                    {"file": "00_metadata.md", "kind": "metadata", "end": {"heading": "Abstract"}},
                    {"file": "01_abstract.md", "start": "Abstract", "end": "Main"},
                    {"file": "02_main.md", "start": "Main", "end": {"heading": "References", "occurrence": 1}},
                    {"file": "03_main_references.md", "start": {"heading": "References", "occurrence": 1}, "end": "Methods"},
                    {"file": "04_methods.md", "start": "Methods", "end": {"heading": "References", "occurrence": 2}},
                    {"file": "05_methods_references.md", "start": {"heading": "References", "occurrence": 2}, "end": "Data availability"},
                    {"file": "06_endmatter.md", "start": "Data availability"},
                ]
            }
        ),
        encoding="utf-8",
    )
    result = docx_to_markdown(source, split_dir=split_dir, section_map_path=section_map, force=True)
    assert result.output is None
    assert (split_dir / "02_main.md").read_text(encoding="utf-8").startswith("# Main")
    assert "| A" in (split_dir / "02_main.md").read_text(encoding="utf-8")
    assert "![" in (split_dir / "02_main.md").read_text(encoding="utf-8")
    assert (split_dir / "media").exists()
    metadata = (split_dir / "00_metadata.md").read_text(encoding="utf-8")
    assert "# TITLE" in metadata and "# AUTHOR" in metadata
    assert "C:\\" not in "\n".join(path.read_text(encoding="utf-8") for path in split_dir.glob("*.md"))


@pytest.mark.skipif(not PANDOC_AVAILABLE, reason="Pandoc is required for integration tests")
def test_unchanged_roundtrip_reuses_exact_source_package(tmp_path: Path) -> None:
    image = tmp_path / "source.png"
    Image.new("RGB", (16, 8), "white").save(image)
    source = tmp_path / "source.docx"
    _sample_docx(source, image)
    split_dir = tmp_path / "sections"
    section_map = tmp_path / "section-map.json"
    section_map.write_text(
        json.dumps(
            {
                "sections": [
                    {"file": "00_metadata.md", "kind": "metadata", "end": "Abstract"},
                    {"file": "01_abstract.md", "start": "Abstract", "end": "Main"},
                    {"file": "02_main.md", "start": "Main", "end": {"heading": "References", "occurrence": 1}},
                    {"file": "03_refs.md", "start": {"heading": "References", "occurrence": 1}, "end": "Methods"},
                    {"file": "04_methods.md", "start": "Methods", "end": {"heading": "References", "occurrence": 2}},
                    {"file": "05_method_refs.md", "start": {"heading": "References", "occurrence": 2}, "end": "Data availability"},
                    {"file": "06_endmatter.md", "start": "Data availability"},
                ]
            }
        ),
        encoding="utf-8",
    )
    docx_to_markdown(source, split_dir=split_dir, section_map_path=section_map, force=True)
    manifest = split_dir / "manifest.json"
    assert (split_dir / "roundtrip" / "source.docx").read_bytes() == source.read_bytes()

    output = tmp_path / "exact.docx"
    assert cli_main(["md2docx", "--roundtrip-manifest", str(manifest), "-o", str(output), "--force"]) == 0
    assert output.read_bytes() == source.read_bytes()

    section = split_dir / "02_main.md"
    section.write_text(section.read_text(encoding="utf-8") + "\nEdited.", encoding="utf-8")
    changed = tmp_path / "changed.docx"
    assert not reuse_unchanged_roundtrip_source(manifest, changed, force=True)


@pytest.mark.skipif(not PANDOC_AVAILABLE, reason="Pandoc is required for integration tests")
def test_docx_to_tex_is_standalone_and_uses_png_media(tmp_path: Path) -> None:
    image = tmp_path / "source.png"
    Image.new("RGB", (16, 8), "white").save(image)
    source = tmp_path / "source.docx"
    _sample_docx(source, image)
    output = tmp_path / "output.tex"
    result = docx_to_tex(source, output=output, force=True)
    text = output.read_text(encoding="utf-8")
    assert result.output == output
    assert "\\documentclass" in text
    assert "\\includegraphics" in text
    assert "media/" in text


def test_sup_sub_markup_is_rendered_as_true_scripts() -> None:
    document = render_blocks_to_doc([Block("paragraph", "H<sub>2</sub>O<sup>+−</sup>")])
    runs = document.paragraphs[0].runs
    assert any(run.text == "2" and run.font.subscript for run in runs)
    assert any(run.text == "+−" and run.font.superscript for run in runs)


def test_pandoc_subscript_run_does_not_capture_formula_delimiters() -> None:
    pandoc_markdown = r"(Et<sub>4</sub>N)<sub>2</sub>\[Cu<sub>8</sub>(N<sub>3)18\]</sub>"
    normalized = normalize_script_boundaries(pandoc_markdown)
    assert normalized == r"(Et<sub>4</sub>N)<sub>2</sub>\[Cu<sub>8</sub>(N<sub>3</sub>)<sub>18</sub>\]"

    document = render_blocks_to_doc([Block("paragraph", normalized)])
    run_modes = [
        (
            run.text,
            "subscript" if run.font.subscript else "superscript" if run.font.superscript else "normal",
        )
        for run in document.paragraphs[0].runs
    ]
    assert ("3", "subscript") in run_modes
    assert (")", "normal") in run_modes
    assert ("18", "subscript") in run_modes
    assert any(text.endswith("]") and mode == "normal" for text, mode in run_modes)


def test_raw_html_table_is_normalized_to_pipe_table_with_rowspans() -> None:
    source = """<table><thead><tr><th>Pressure</th><th><em>χ</em><sup>(2)</sup></th></tr></thead>
    <tbody><tr><td rowspan="2">3.7 GPa</td><td><em>χ</em><sub>xxy</sub></td></tr>
    <tr><td><em>χ</em><sub>xyz</sub></td></tr></tbody></table>"""
    normalized = normalize_inline_html(normalize_html_tables(source))
    assert "| Pressure | *χ*<sup>(2)</sup> |" in normalized
    assert normalized.count("| 3.7 GPa |") == 2
    assert "*χ*<sub>xxy</sub>" in normalized
    assert "<table" not in normalized


def test_split_markdown_sections_handles_duplicate_headings() -> None:
    text = "**Title**\n\n**Abstract**:\n\nA\n\n**Main**\n\nB\n\n**References**\n\n1. R\n\n**Methods**\n\nC\n\n**References**\n\n51. S\n\n**Data availability**\n\nD\n"
    sections = split_markdown_sections(
        text,
        [
            {"file": "00_metadata.md", "kind": "metadata", "end": "Abstract"},
            {"file": "01_abstract.md", "start": "Abstract", "end": "Main"},
            {"file": "02_main.md", "start": "Main", "end": {"heading": "References", "occurrence": 1}},
            {"file": "03_refs.md", "start": {"heading": "References", "occurrence": 1}, "end": "Methods"},
            {"file": "04_methods.md", "start": "Methods", "end": {"heading": "References", "occurrence": 2}},
            {"file": "05_refs.md", "start": {"heading": "References", "occurrence": 2}, "end": "Data availability"},
            {"file": "06_end.md", "start": "Data availability"},
        ],
    )
    assert [name for name, _ in sections] == [
        "00_metadata.md", "01_abstract.md", "02_main.md", "03_refs.md", "04_methods.md", "05_refs.md", "06_end.md"
    ]
    assert sections[2][1].startswith("# Main")
    assert sections[5][1].startswith("# References")


def test_run_pandoc_reports_missing_external_dependency(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    def missing(*args, **kwargs):
        raise FileNotFoundError("pandoc")

    monkeypatch.setattr(_pandoc.subprocess, "run", missing)
    with pytest.raises(RuntimeError, match="pandoc is required"):
        _pandoc.run_pandoc(tmp_path / "input.docx", tmp_path / "out.md", output_format="gfm")


def _tracked_current(path: Path) -> None:
    source = path.with_suffix(".source.docx")
    document = Document()
    document.add_paragraph("after")
    document.save(source)
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = __import__("lxml.etree", fromlist=["etree"]).fromstring(parts["word/document.xml"])
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraph = root.find(".//w:body/w:p", ns)
    assert paragraph is not None
    run = paragraph.find("./w:r", ns)
    assert run is not None
    paragraph.remove(run)
    deleted = OxmlElement("w:del")
    deleted.set(qn("w:id"), "1")
    deleted_run = OxmlElement("w:r")
    deleted_text = OxmlElement("w:delText")
    deleted_text.text = "before"
    deleted_run.append(deleted_text)
    deleted.append(deleted_run)
    inserted = OxmlElement("w:ins")
    inserted.set(qn("w:id"), "2")
    inserted_run = OxmlElement("w:r")
    inserted_text = OxmlElement("w:t")
    inserted_text.text = "after"
    inserted_run.append(inserted_text)
    inserted.append(inserted_run)
    paragraph.append(deleted)
    paragraph.append(inserted)
    parts["word/document.xml"] = __import__("lxml.etree", fromlist=["etree"]).tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    source.unlink()


def test_redline_accepts_current_revisions_before_diff(tmp_path: Path) -> None:
    base = tmp_path / "base.docx"
    base_document = Document()
    base_document.add_paragraph("before")
    base_document.save(base)
    current = tmp_path / "current.docx"
    _tracked_current(current)
    output = tmp_path / "tracked.docx"
    summary = create_tracked_docx(base, current, output, overwrite=True)
    assert summary["changed"] == 1
    with zipfile.ZipFile(output) as archive:
        from lxml import etree

        root = etree.fromstring(archive.read("word/document.xml"))
    assert visible_text(root, "final").strip() == "after"
    assert visible_text(root, "original").strip() == "before"
    assert not root.xpath(".//w:rPrChange", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})


def test_redline_tracks_inserted_table_rows(tmp_path: Path) -> None:
    base = tmp_path / "base-table.docx"
    current = tmp_path / "current-table.docx"
    for path, values in (
        (base, (("Pressure", "Value"), ("0 GPa", "1"), ("10 GPa", "3"))),
        (current, (("Pressure", "Value"), ("0 GPa", "1"), ("3.7 GPa", "2"), ("10 GPa", "3"))),
    ):
        document = Document()
        table = document.add_table(rows=len(values), cols=2)
        for row, data in zip(table.rows, values):
            for cell, value in zip(row.cells, data):
                cell.text = value
        document.save(path)
    output = tmp_path / "tracked-table.docx"
    create_tracked_docx(base, current, output, overwrite=True)
    with zipfile.ZipFile(output) as archive:
        from lxml import etree

        root = etree.fromstring(archive.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    assert len(root.xpath(".//w:tr[w:trPr/w:ins]", namespaces=ns)) == 1
    assert "3.7 GPa" in visible_text(root, "final")
    assert "3.7 GPa" not in visible_text(root, "original")


def test_redline_tracks_moved_table_rows(tmp_path: Path) -> None:
    base = tmp_path / "base.docx"
    old = Document()
    old.add_paragraph("Before table.")
    old.add_table(rows=1, cols=1).cell(0, 0).text = "Moved data"
    old.add_paragraph("After table.")
    old.save(base)

    current = tmp_path / "current.docx"
    new = Document()
    new.add_paragraph("Before table.")
    new.add_paragraph("After table.")
    new.add_table(rows=1, cols=1).cell(0, 0).text = "Moved data"
    new.save(current)

    tracked = tmp_path / "tracked.docx"
    create_tracked_docx(base, current, tracked)
    with zipfile.ZipFile(tracked) as archive:
        from lxml import etree

        root = etree.fromstring(archive.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    assert root.xpath(".//w:trPr/w:ins", namespaces=ns)
    assert root.xpath(".//w:trPr/w:del", namespaces=ns)
    with zipfile.ZipFile(current) as archive:
        current_root = etree.fromstring(archive.read("word/document.xml"))
    assert _final_blocks(_accepted_revision_view(root)) == _final_blocks(current_root)


def test_redline_keeps_small_citation_edits_inside_one_paragraph(tmp_path: Path) -> None:
    from lxml import etree

    def cited(path: Path, citation: str, anchor: str) -> None:
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Detonation velocity is estimated by")
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), anchor)
        run = OxmlElement("w:r")
        props = OxmlElement("w:rPr")
        vert = OxmlElement("w:vertAlign")
        vert.set(qn("w:val"), "superscript")
        props.append(vert)
        run.append(props)
        node = OxmlElement("w:t")
        node.text = citation
        run.append(node)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
        paragraph.add_run(".")
        document.save(path)

    base = tmp_path / "base.docx"
    current = tmp_path / "current.docx"
    cited(base, "37", "ref_37")
    cited(current, "33", "ref_33")
    tracked = tmp_path / "tracked.docx"
    create_tracked_docx(base, current, tracked)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(tracked) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    paragraphs = root.xpath(".//w:body/w:p", namespaces=ns)
    assert len(paragraphs) == 1
    assert paragraphs[0].find("w:pPr/w:rPr/w:del", ns) is None
    assert [node.text for node in paragraphs[0].findall(".//w:delText", ns)] == ["37"]
    assert [node.text for node in paragraphs[0].findall(".//w:hyperlink//w:t", ns)] == ["33"]
    with zipfile.ZipFile(current) as archive:
        current_root = etree.fromstring(archive.read("word/document.xml"))
    assert _final_blocks(_accepted_revision_view(root)) == _final_blocks(current_root)


def test_redline_preserves_original_hyperlink_paragraph_and_picture(tmp_path: Path) -> None:
    from lxml import etree

    base = tmp_path / "base.docx"
    current = tmp_path / "current.docx"
    for path, text, color in ((base, "July wording", "red"), (current, "Current wording", "blue")):
        image = tmp_path / f"{color}.png"
        Image.new("RGB", (12, 12), color).save(image)
        document = Document()
        paragraph = document.add_paragraph()
        hyperlink = OxmlElement("w:hyperlink")
        run = OxmlElement("w:r")
        node = OxmlElement("w:t")
        node.text = text
        run.append(node)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
        document.add_paragraph().add_run().add_picture(str(image))
        document.save(path)

    tracked = tmp_path / "tracked.docx"
    create_tracked_docx(base, current, tracked)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(tracked) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        assert archive.testzip() is None
    with zipfile.ZipFile(base) as archive:
        base_root = etree.fromstring(archive.read("word/document.xml"))
    with zipfile.ZipFile(current) as archive:
        current_root = etree.fromstring(archive.read("word/document.xml"))
    original = [
        visible_text(p, "original")
        for p in root.xpath(".//w:body/w:p[not(w:pPr/w:rPr/w:ins)]", namespaces=ns)
    ]
    expected = [visible_text(p) for p in base_root.xpath(".//w:body/w:p", namespaces=ns)]
    assert original == expected
    assert _final_blocks(_accepted_revision_view(root)) == _final_blocks(current_root)
    # The picture still moves as a whole paragraph. The hyperlink sentence is
    # an in-paragraph revision, so its old wording is struck out rather than
    # left behind as a second live paragraph.
    assert len(root.xpath(".//w:p[w:pPr/w:rPr/w:del]", namespaces=ns)) == 1
    assert len(root.xpath(".//w:p[w:pPr/w:rPr/w:ins]", namespaces=ns)) == 1
    deleted_drawings = root.xpath(".//w:p[w:pPr/w:rPr/w:del]//w:drawing", namespaces=ns)
    assert deleted_drawings == []
    assert len(root.xpath(".//w:p[w:pPr/w:rPr/w:ins]//w:drawing", namespaces=ns)) == 1
    wording = next(p for p in root.xpath(".//w:body/w:p", namespaces=ns) if "wording" in visible_text(p, "original"))
    assert visible_text(wording, "original") == "July wording"
    assert visible_text(wording, "final") == "Current wording"
    assert wording.find("w:pPr/w:rPr/w:del", ns) is None
    assert wording.xpath(".//w:hyperlink", namespaces=ns)


def _run_with_align(text: str, align: str | None):
    run = OxmlElement("w:r")
    if align is not None:
        properties = OxmlElement("w:rPr")
        marker = OxmlElement("w:vertAlign")
        marker.set(qn("w:val"), align)
        properties.append(marker)
        run.append(properties)
    node = OxmlElement("w:t")
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    run.append(node)
    return run


def test_redline_does_not_subscript_formula_delimiters(tmp_path: Path) -> None:
    base = tmp_path / "base.docx"
    current = tmp_path / "current.docx"
    for path, tail in ((base, "here."), (current, "now.")):
        document = Document()
        paragraph = document.add_paragraph()
        paragraph._p.append(_run_with_align("[K(ClO", "baseline"))
        paragraph._p.append(_run_with_align("4", "subscript"))
        paragraph._p.append(_run_with_align(")", "baseline"))
        paragraph._p.append(_run_with_align("6", "subscript"))
        paragraph._p.append(_run_with_align("] units ", "baseline"))
        paragraph._p.append(_run_with_align(tail, "baseline"))
        document.save(path)
    output = tmp_path / "tracked.docx"
    create_tracked_docx(base, current, output, overwrite=True)
    from lxml import etree

    root = etree.fromstring(zipfile.ZipFile(output).read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    leaked = []
    for run in root.xpath(".//w:r[not(ancestor::w:del)]", namespaces=ns):
        text = "".join(run.xpath(".//w:t/text()", namespaces=ns))
        align = run.xpath("string(w:rPr/w:vertAlign/@w:val)", namespaces=ns)
        if any(char in text for char in "()[]") and align == "subscript":
            leaked.append(text)
    assert leaked == []


def test_visible_text_includes_math_and_bookmarks_survive_merge() -> None:
    from lxml import etree

    from docforge.docxdiff.redline import M, W, Context, _merge_paragraph

    nsmap = {"w": W, "m": M}
    base = etree.Element(f"{{{W}}}p", nsmap=nsmap)
    start = etree.SubElement(base, f"{{{W}}}bookmarkStart")
    start.set(f"{{{W}}}id", "4")
    start.set(f"{{{W}}}name", "eq-anchor")
    math = etree.SubElement(base, f"{{{M}}}oMath")
    token = etree.SubElement(math, f"{{{M}}}t")
    token.text = "Vdet"
    end = etree.SubElement(base, f"{{{W}}}bookmarkEnd")
    end.set(f"{{{W}}}id", "4")
    assert visible_text(base) == "Vdet"

    current = etree.Element(f"{{{W}}}p", nsmap=nsmap)
    replacement = etree.SubElement(current, f"{{{M}}}oMath")
    replacement_text = etree.SubElement(replacement, f"{{{M}}}t")
    replacement_text.text = "V"
    merged = _merge_paragraph(base, current, Context("tester", 1))
    assert merged.find(f"{{{W}}}bookmarkStart").get(f"{{{W}}}name") == "eq-anchor"
    assert merged.find(f"{{{W}}}bookmarkEnd") is not None

    edited = etree.Element(f"{{{W}}}p", nsmap=nsmap)
    edited_start = etree.SubElement(edited, f"{{{W}}}bookmarkStart")
    edited_start.set(f"{{{W}}}id", "9")
    edited_start.set(f"{{{W}}}name", "cite-anchor")
    run = etree.SubElement(edited, f"{{{W}}}r")
    text = etree.SubElement(run, f"{{{W}}}t")
    text.text = "Alpha"
    edited_end = etree.SubElement(edited, f"{{{W}}}bookmarkEnd")
    edited_end.set(f"{{{W}}}id", "9")
    changed = etree.Element(f"{{{W}}}p", nsmap=nsmap)
    changed_run = etree.SubElement(changed, f"{{{W}}}r")
    changed_text = etree.SubElement(changed_run, f"{{{W}}}t")
    changed_text.text = "Alpha beta"
    merged_text = _merge_paragraph(edited, changed, Context("tester", 1))
    assert merged_text.find(f"{{{W}}}bookmarkStart").get(f"{{{W}}}name") == "cite-anchor"
    assert merged_text.find(f"{{{W}}}bookmarkEnd") is not None


def test_page_break_survives_merge_against_empty_base_paragraph() -> None:
    from lxml import etree

    from docforge.docxdiff.redline import W, Context, _merge_paragraph

    base = etree.Element(f"{{{W}}}p", nsmap={"w": W})
    current = etree.Element(f"{{{W}}}p", nsmap={"w": W})
    run = etree.SubElement(current, f"{{{W}}}r")
    etree.SubElement(run, f"{{{W}}}br").set(f"{{{W}}}type", "page")
    merged = _merge_paragraph(base, current, Context("tester", 1))
    breaks = merged.findall(f".//{{{W}}}br")
    assert [item.get(f"{{{W}}}type") for item in breaks] == ["page"]


def test_multiline_display_math_is_one_numbered_equation(tmp_path: Path) -> None:
    from lxml import etree

    from docforge.tex.converter import latex_to_docx

    source = r"""
\documentclass{article}
\begin{document}
Before the table.
\docxpagebreak
\begin{equation}
\begin{aligned}
a &= 1 \\
b &= 2
\end{aligned}
\end{equation}
\end{document}
"""
    out = tmp_path / "math.docx"
    latex_to_docx(source, output=out)
    root = etree.fromstring(zipfile.ZipFile(out).read("word/document.xml"))
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    }
    assert len(root.xpath(".//w:br[@w:type='page']", namespaces=ns)) == 1
    arrays = root.xpath(".//m:eqArr", namespaces=ns)
    assert len(arrays) == 1
    assert len(arrays[0].xpath("./m:e", namespaces=ns)) == 2
    equation = arrays[0].getparent().getparent()
    assert "(1)" in visible_text(equation)
    assert len(root.xpath(".//m:eqArr/ancestor::w:p", namespaces=ns)) == 1


def test_ratio_cache_keeps_tracked_output_and_is_reused(tmp_path: Path) -> None:
    base, current = Document(), Document()
    for index in range(12):
        base.add_paragraph(f"Paragraph {index} describes cluster growth from seed ions.")
        current.add_paragraph(f"Paragraph {index} describes cluster growth from {index} seed ions.")
    current.add_paragraph("An inserted closing paragraph.")
    base.save(tmp_path / "base.docx")
    current.save(tmp_path / "current.docx")
    cache = tmp_path / "cache" / "ratios.json"

    plain = create_tracked_docx(tmp_path / "base.docx", tmp_path / "current.docx", tmp_path / "plain.docx", workers=1)
    cold = create_tracked_docx(tmp_path / "base.docx", tmp_path / "current.docx", tmp_path / "cold.docx", ratio_cache=cache, workers=1)
    stored = json.loads(cache.read_text(encoding="utf-8"))
    warm = create_tracked_docx(tmp_path / "base.docx", tmp_path / "current.docx", tmp_path / "warm.docx", ratio_cache=cache, workers=1)

    assert plain == cold == warm
    assert stored and json.loads(cache.read_text(encoding="utf-8")) == stored
    views = {
        name: [visible_text(p._p) for p in Document(tmp_path / f"{name}.docx").paragraphs]
        for name in ("plain", "cold", "warm")
    }
    assert views["plain"] == views["cold"] == views["warm"]


def test_deleted_hyperlink_field_does_not_keep_field_codes(tmp_path: Path) -> None:
    from lxml import etree

    def add_hyperlink_field(paragraph, text: str) -> None:
        def append_run(*nodes) -> None:
            run = OxmlElement("w:r")
            for node in nodes:
                run.append(node)
            paragraph._p.append(run)

        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instruction = OxmlElement("w:instrText")
        instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instruction.text = ' HYPERLINK "https://example.com" '
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        append_run(begin)
        append_run(instruction)
        append_run(separate)
        paragraph.add_run(text)
        append_run(end)

    base = tmp_path / "base.docx"
    current = tmp_path / "current.docx"
    document = Document()
    add_hyperlink_field(document.add_paragraph(), "Andreas Marek")
    document.save(base)
    document = Document()
    document.add_paragraph("Replacement author")
    document.save(current)
    tracked = tmp_path / "tracked.docx"
    create_tracked_docx(base, current, tracked, overwrite=True)
    root = etree.fromstring(zipfile.ZipFile(tracked).read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    assert root.xpath(".//w:del//w:fldChar", namespaces=ns) == []
    assert "Andreas Marek" in visible_text(root, "original")
