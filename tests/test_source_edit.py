from __future__ import annotations

from docx import Document

from docforge.docxdiff import create_tracked_docx
from docforge.docxdiff.redline import visible_text
from docforge.markdown.source_edit import apply_markdown_delta, export_block_map


def test_block_map_and_noop_preserve_document_xml_and_media(tmp_path):
    from zipfile import ZipFile

    source = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("A source paragraph with enough text to map.")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "opaque table"
    document.save(source)
    baseline = tmp_path / "baseline.md"
    edited = tmp_path / "edited.md"
    body = "A source paragraph with enough text to map.\n\n| value |\n| --- |\n| opaque table |"
    baseline.write_text(body, encoding="utf-8")
    edited.write_text(body, encoding="utf-8")
    block_map = tmp_path / "blocks.json"
    payload = export_block_map(source, block_map)
    assert payload["text"]["0"] == 0
    assert any(entry["kind"] == "table" and entry["opaque_id"] for entry in payload["blocks"])
    output = tmp_path / "output.docx"
    apply_markdown_delta(source, baseline, edited, output, block_map=block_map, overwrite=True)
    with ZipFile(source) as left, ZipFile(output) as right:
        assert left.read("word/document.xml") == right.read("word/document.xml")
        assert {
            name: left.read(name)
            for name in left.namelist()
            if name.startswith("word/media/")
        } == {
            name: right.read(name)
            for name in right.namelist()
            if name.startswith("word/media/")
        }


def test_opaque_map_change_fails_before_output(tmp_path):
    import json

    source = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("A source paragraph with enough text to map.")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "opaque table"
    document.save(source)
    baseline = tmp_path / "baseline.md"
    edited = tmp_path / "edited.md"
    baseline.write_text("A source paragraph with enough text to map.", encoding="utf-8")
    edited.write_text("A changed source paragraph with enough text to map.", encoding="utf-8")
    block_map = tmp_path / "blocks.json"
    payload = export_block_map(source, block_map)
    table = next(entry for entry in payload["blocks"] if entry["kind"] == "table")
    table["signature"] = "tampered"
    block_map.write_text(json.dumps(payload), encoding="utf-8")
    output = tmp_path / "output.docx"
    import pytest

    with pytest.raises(ValueError, match="opaque block changed"):
        apply_markdown_delta(source, baseline, edited, output, block_map=block_map, overwrite=True)
    assert not output.exists()


def test_delete_requires_explicit_opt_in(tmp_path):
    source = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("A source paragraph with enough text to map.")
    document.add_paragraph("Another source paragraph with enough text to map.")
    document.save(source)
    baseline = tmp_path / "baseline.md"
    edited = tmp_path / "edited.md"
    baseline.write_text(
        "A source paragraph with enough text to map.\n\nAnother source paragraph with enough text to map.",
        encoding="utf-8",
    )
    edited.write_text("A source paragraph with enough text to map.", encoding="utf-8")
    import pytest

    with pytest.raises(ValueError, match="allow_delete"):
        apply_markdown_delta(source, baseline, edited, tmp_path / "refused.docx", overwrite=True)
    output = tmp_path / "deleted.docx"
    summary = apply_markdown_delta(source, baseline, edited, output, allow_delete=True, overwrite=True)
    assert summary["deleted"] == 1
    assert [paragraph.text for paragraph in Document(output).paragraphs] == [
        "A source paragraph with enough text to map."
    ]


def test_replacement_extras_stay_after_intervening_table(tmp_path):
    from zipfile import ZipFile
    from lxml import etree

    source = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("First paragraph has enough words for matching.")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "opaque table content"
    document.add_paragraph("Second paragraph has enough words for matching.")
    document.save(source)
    baseline = tmp_path / "baseline.md"
    edited = tmp_path / "edited.md"
    baseline.write_text(
        "First paragraph has enough words for matching.\n\n"
        "| value |\n| --- |\n| opaque table content |\n\n"
        "Second paragraph has enough words for matching.",
        encoding="utf-8",
    )
    edited.write_text(
        "First paragraph has changed words for matching.\n\n"
        "Inserted after the opaque table and before the next match.\n\n"
        "| value |\n| --- |\n| opaque table content |\n\n"
        "Second paragraph has enough words for matching.",
        encoding="utf-8",
    )
    output = tmp_path / "output.docx"
    apply_markdown_delta(source, baseline, edited, output, overwrite=True)
    with ZipFile(output) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    body = root.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
    assert [etree.QName(child).localname for child in body] == ["p", "tbl", "p", "p", "sectPr"]


def test_markdown_delta_keeps_unedited_paragraphs(tmp_path):
    source = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("Electronic Supplementary Information")
    document.add_paragraph("Keep this sentence exactly, including its place in the document.")
    document.add_paragraph(
        "EAP-4 and EAP-8 show nearly identical onset temperatures but very different decomposition behaviors."
    )
    document.add_paragraph("Sensitivity Tests")
    document.add_paragraph("Friction sensitivity measurements were performed in accordance with STANAG 4487.")
    document.save(source)

    baseline = tmp_path / "baseline.md"
    edited = tmp_path / "edited.md"
    baseline.write_text(
        "\n\n".join(
            [
                "Electronic Supplementary Information",
                "Keep this sentence exactly, including its place in the document.",
                "EAP-4 and EAP-8 show nearly identical onset temperatures but very different decomposition behaviors.",
                "4.  **Sensitivity Tests**",
                "Friction sensitivity measurements were performed in accordance with STANAG 4487.",
            ]
        ),
        encoding="utf-8",
    )
    edited.write_text(
        "\n\n".join(
            [
                "Electronic Supplementary Information",
                "Keep this sentence exactly, including its place in the document.",
                "EAP-4 and **EAP-8** show nearly identical onset temperatures but markedly different decomposition kinetics.",
                "The 1666 K trajectories follow nitrogen provenance.",
                "## Reactive molecular dynamics",
                "The trajectories used a four-member committee.",
                "4.  **Sensitivity Tests**",
                "Friction sensitivity measurements were performed in accordance with STANAG 4487.",
            ]
        ),
        encoding="utf-8",
    )
    current = tmp_path / "current.docx"
    summary = apply_markdown_delta(source, baseline, edited, current, overwrite=True)
    assert summary["replaced"] == 1
    assert summary["inserted"] == 0

    current_doc = Document(current)
    texts = [paragraph.text for paragraph in current_doc.paragraphs]
    assert texts[0] == "Electronic Supplementary Information"
    assert texts[1] == "Keep this sentence exactly, including its place in the document."
    assert "markedly different" in texts[2]
    assert "1666 K" in texts[3]
    assert texts[4] == "Reactive molecular dynamics"
    assert "four-member" in texts[5]
    assert texts[6] == "Sensitivity Tests"

    tracked = tmp_path / "tracked.docx"
    redline = create_tracked_docx(source, current, tracked, author="gmy", overwrite=True)
    assert redline["deleted"] == 0
    accepted = []
    from lxml import etree
    from zipfile import ZipFile

    root = etree.fromstring(ZipFile(tracked).read("word/document.xml"))
    for paragraph in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
        text = visible_text(paragraph).strip()
        if text:
            accepted.append(text)
    assert "Keep this sentence exactly, including its place in the document." in accepted
    assert "Electronic Supplementary Information" in accepted


def test_insert_keeps_anchor_size_and_unescapes_markdown(tmp_path):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    source = tmp_path / "source.docx"
    document = Document()
    title = document.add_paragraph("Electronic Supplementary Information")
    title.runs[0].font.size = Pt(18)
    title.runs[0].bold = True
    body = document.add_paragraph("Keep this sentence exactly, including its place in the document.")
    body.runs[0].font.size = Pt(12)
    heading = document.add_paragraph("Sensitivity Tests")
    heading.runs[0].bold = True
    heading.runs[0].font.size = Pt(14)
    p_pr = heading._p.get_or_add_pPr()
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "1")
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), "2")
    num_pr.append(ilvl)
    num_pr.append(num_id)
    p_pr.append(num_pr)
    document.save(source)

    baseline = tmp_path / "baseline.md"
    edited = tmp_path / "edited.md"
    baseline.write_text(
        "\n\n".join(
            [
                "Electronic Supplementary Information",
                "Keep this sentence exactly, including its place in the document.",
                "4.  **Sensitivity Tests**",
            ]
        ),
        encoding="utf-8",
    )
    edited.write_text(
        "\n\n".join(
            [
                "Electronic Supplementary Information",
                "Keep this sentence exactly, including its place in the document.",
                "## Reactive molecular dynamics",
                r"\[11\] Y. Zhang, H. Wang, example citation for the delta.",
                "4.  **Sensitivity Tests**",
            ]
        ),
        encoding="utf-8",
    )
    current = tmp_path / "current.docx"
    apply_markdown_delta(source, baseline, edited, current, overwrite=True)
    current_doc = Document(current)
    texts = [paragraph.text for paragraph in current_doc.paragraphs]
    assert texts[2] == "Reactive molecular dynamics"
    assert texts[3].startswith("[11] Y. Zhang")
    assert "\\[" not in texts[3]
    heading_run = current_doc.paragraphs[2].runs[0]
    citation_run = current_doc.paragraphs[3].runs[0]
    assert heading_run.bold is True
    assert heading_run.font.size == Pt(14)
    assert citation_run.font.size == Pt(12)
    assert current_doc.paragraphs[2]._p.find(qn("w:pPr")).find(qn("w:numPr")) is None
