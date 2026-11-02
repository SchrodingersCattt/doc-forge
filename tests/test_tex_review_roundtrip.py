"""Synthetic tests for the explicitly limited TeX review bundle."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

from docforge.cli import main
from docforge.docxdiff.package import Package
from docforge.tex.review_roundtrip import ReviewRoundtripError, build_reviewed_docx, export_review_bundle


@pytest.fixture
def inputs(tmp_path: Path):
    picture = tmp_path / "image.png"
    Image.new("RGB", (5, 5), "white").save(picture)
    baseline = tmp_path / "baseline.docx"
    reviewed = tmp_path / "reviewed.docx"
    source = tmp_path / "source.tex"
    doc = Document()
    paragraph = doc.add_paragraph("The sample was cold today.")
    doc.add_comment(paragraph.runs, "Check this wording", author="Reviewer")
    doc.add_paragraph("See Supplementary Note 2.")
    doc.add_paragraph().add_run().add_picture(str(picture))
    doc.save(baseline)
    doc = Document(baseline)
    doc.paragraphs[0].runs[0].text = "The sample was warm today."
    doc.save(reviewed)
    source.write_text(
        "\\begin{document}\nThe sample was cold today.\n"
        "See Supplementary Note~\\ref{S-sn:example}.\n\\end{document}\n",
        encoding="utf-8",
    )
    return baseline, reviewed, source


def _parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_review_export_and_noop_reuses_exact_reviewed_bytes(inputs, tmp_path):
    baseline, reviewed, source = inputs
    bundle = tmp_path / "bundle"
    manifest = export_review_bundle(baseline, reviewed, source, bundle)
    text = (bundle / "document.tex").read_text(encoding="utf-8")
    assert "The sample was warm today." in text
    assert r"\ref{S-sn:example}" in text
    assert manifest["reviewed_changed_paragraphs"] == 1
    output = tmp_path / "out.docx"
    assert build_reviewed_docx(bundle, output)["mode"] == "EXACT_REUSE"
    assert output.read_bytes() == reviewed.read_bytes()


def test_plain_tex_edit_patches_only_reviewed_text_and_keeps_package(inputs, tmp_path):
    baseline, reviewed, source = inputs
    bundle = tmp_path / "bundle"
    export_review_bundle(baseline, reviewed, source, bundle)
    editable = bundle / "document.tex"
    editable.write_text(editable.read_text(encoding="utf-8").replace("warm", "mild"), encoding="utf-8")
    output = tmp_path / "out.docx"
    report = build_reviewed_docx(bundle, output)
    assert report == {"mode": "SOURCE_PATCH", "changed_paragraphs": 1}
    assert Document(output).paragraphs[0].text == "The sample was mild today."
    original, rebuilt = _parts(reviewed), _parts(output)
    assert original.keys() == rebuilt.keys()
    assert all(original[name] == rebuilt[name] for name in original if name != "word/document.xml")
    a, b = Package.load(reviewed).xml("word/document.xml"), Package.load(output).xml("word/document.xml")
    for selector in (".//w:commentRangeStart", ".//w:commentRangeEnd", ".//w:commentReference", ".//w:drawing"):
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        assert len(a.xpath(selector, namespaces=ns)) == len(b.xpath(selector, namespaces=ns)) == 1


@pytest.mark.parametrize("edit", [
    lambda text: text.replace("S-sn:example", "S-sn:changed"),
    lambda text: text.replace("warm", "warm%"),
])
def test_unmapped_or_tex_syntax_edit_leaves_existing_output_untouched(inputs, tmp_path, edit):
    baseline, reviewed, source = inputs
    bundle = tmp_path / "bundle"
    export_review_bundle(baseline, reviewed, source, bundle)
    editable = bundle / "document.tex"
    editable.write_text(edit(editable.read_text(encoding="utf-8")), encoding="utf-8")
    output = tmp_path / "out.docx"
    output.write_bytes(b"previous output")
    with pytest.raises(ReviewRoundtripError):
        build_reviewed_docx(bundle, output, force=True)
    assert output.read_bytes() == b"previous output"


def test_duplicate_source_text_and_complex_review_fail_before_publish(inputs, tmp_path):
    baseline, reviewed, source = inputs
    source.write_text(source.read_text(encoding="utf-8") + "The sample was cold today.\n", encoding="utf-8")
    bundle = tmp_path / "bundle"
    with pytest.raises(ReviewRoundtripError, match="unique"):
        export_review_bundle(baseline, reviewed, source, bundle)
    assert not bundle.exists()


def test_tracked_change_accepts_visible_text_but_preserves_revision_xml(inputs, tmp_path):
    baseline, reviewed, source = inputs
    package = Package.load(baseline)
    root = package.xml("word/document.xml")
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraph = root.xpath("./w:body/w:p", namespaces=ns)[0]
    original_run = paragraph.xpath("./w:r[w:t]", namespaces=ns)[0]
    position = paragraph.index(original_run)
    paragraph.remove(original_run)
    for tag, word in ((None, "The sample was "), ("del", "cold"), ("ins", "warm"), (None, " today.")):
        run = OxmlElement("w:r")
        text = OxmlElement("w:delText" if tag == "del" else "w:t")
        text.text = word
        if word[:1].isspace() or word[-1:].isspace():
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        run.append(text)
        if tag:
            wrapper = OxmlElement("w:" + tag)
            wrapper.set(qn("w:id"), "21" if tag == "ins" else "20")
            wrapper.set(qn("w:author"), "Reviewer")
            wrapper.append(run)
            paragraph.insert(position, wrapper)
        else:
            paragraph.insert(position, run)
        position += 1
    package.set_xml("word/document.xml", root)
    package.write(reviewed, overwrite=True)
    bundle = tmp_path / "bundle"
    export_review_bundle(baseline, reviewed, source, bundle)
    output = tmp_path / "out.docx"
    assert build_reviewed_docx(bundle, output)["mode"] == "EXACT_REUSE"
    assert output.read_bytes() == reviewed.read_bytes()
    editable = bundle / "document.tex"
    editable.write_text(editable.read_text(encoding="utf-8").replace("warm", "mild"), encoding="utf-8")
    patched = tmp_path / "patched.docx"
    build_reviewed_docx(bundle, patched)
    root = Package.load(patched).xml("word/document.xml")
    assert root.xpath("count(.//w:ins)", namespaces=ns) == 1
    assert root.xpath("count(.//w:del)", namespaces=ns) == 1
    assert root.xpath("string(.//w:ins//w:t)", namespaces=ns) == "mild"


def test_cross_run_edit_and_changed_snapshot_fail_closed(inputs, tmp_path):
    baseline, reviewed, source = inputs
    doc = Document(reviewed)
    p = doc.paragraphs[0]
    p.runs[0].text = "The sample was "
    p.add_run("warm today.")
    doc.save(reviewed)
    bundle = tmp_path / "bundle"
    export_review_bundle(baseline, reviewed, source, bundle)
    editable = bundle / "document.tex"
    editable.write_text(editable.read_text(encoding="utf-8").replace("was warm", "became hot"), encoding="utf-8")
    with pytest.raises(ReviewRoundtripError, match="boundary"):
        build_reviewed_docx(bundle, tmp_path / "out.docx")
    (bundle / "reviewed.docx").write_bytes(b"tampered")
    with pytest.raises(ReviewRoundtripError, match="hash mismatch"):
        build_reviewed_docx(bundle, tmp_path / "out.docx")


def test_review_formatting_change_is_not_silently_dropped(inputs, tmp_path):
    baseline, reviewed, source = inputs
    doc = Document(reviewed)
    doc.paragraphs[0].runs[0].bold = True
    doc.save(reviewed)
    with pytest.raises(ReviewRoundtripError, match="formatting"):
        export_review_bundle(baseline, reviewed, source, tmp_path / "bundle")
    assert not (tmp_path / "bundle").exists()


def test_plain_text_inside_other_tex_environment_is_not_a_prose_anchor(inputs, tmp_path):
    baseline, reviewed, source = inputs
    source.write_text(
        "\\begin{document}\n\\begin{verbatim}\nThe sample was cold today.\n"
        "\\end{verbatim}\n\\end{document}\n", encoding="utf-8",
    )
    with pytest.raises(ReviewRoundtripError, match="unique plain"):
        export_review_bundle(baseline, reviewed, source, tmp_path / "bundle")


def test_text_with_a_word_field_is_not_silently_rewritten(inputs, tmp_path):
    baseline, reviewed, source = inputs
    for path in (baseline, reviewed):
        package = Package.load(path)
        root = package.xml("word/document.xml")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        p = root.xpath("./w:body/w:p", namespaces=ns)[0]
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), 'REF example')
        p.append(field)
        package.set_xml("word/document.xml", root)
        package.write(path, overwrite=True)
    with pytest.raises(ReviewRoundtripError, match="simple text"):
        export_review_bundle(baseline, reviewed, source, tmp_path / "bundle")


def test_cli_exports_and_rebuilds_bundle(inputs, tmp_path):
    baseline, reviewed, source = inputs
    bundle = tmp_path / "bundle"
    output = tmp_path / "out.docx"
    assert main(["tex-review-export", str(baseline), str(reviewed), str(source), "--workdir", str(bundle)]) == 0
    assert main(["tex-review-build", str(bundle), "-o", str(output)]) == 0
    assert output.read_bytes() == reviewed.read_bytes()
