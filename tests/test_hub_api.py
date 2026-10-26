"""Hub API, legacy parity, and fidelity behavior on synthetic documents."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree
from docx.shared import Pt

from docforge import convert, read, write
from docforge.adapters.protocol import ConversionError
from docforge.cli import main as cli_main
from docforge.style import load_profile, resolved_paragraph_styles, restyle_docx


def test_public_commands_remain_available() -> None:
    from docforge.cli import build_parser

    choices = set(build_parser()._subparsers._group_actions[0].choices)
    required = {
        "md2docx", "tex2docx", "docx2md", "docx2tex", "md-export", "md-build",
        "roundtrip", "redline", "contract", "deliver", "convert", "restyle", "audit",
        "export", "build",
    }
    assert required <= choices


def test_md2docx_api_matches_cli(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    source.write_text(
        "# Sample\n\nA paragraph with **bold** and *italic* text.\n\n- alpha\n- beta\n\n"
        "1. first\n2. second\n",
        encoding="utf-8",
    )
    via_cli = tmp_path / "cli.docx"
    via_api = tmp_path / "api.docx"
    assert cli_main(["md2docx", str(source), "-o", str(via_cli), "--force"]) == 0
    result = convert(source, via_api, force=True)
    assert result.execution_mode == "REBUILD"
    assert _docx_view(via_cli) == _docx_view(via_api)
    assert hashlib.sha256(source.read_bytes()).hexdigest() == hashlib.sha256(source.read_bytes()).hexdigest()


def test_tex2docx_api_matches_cli(tmp_path: Path) -> None:
    source = tmp_path / "sample.tex"
    source.write_text(
        "\\documentclass{article}\n\\begin{document}\n\\section{Introduction}\n"
        "A short paragraph with \\textbf{bold} text.\n\\end{document}\n",
        encoding="utf-8",
    )
    via_cli = tmp_path / "cli.docx"
    via_api = tmp_path / "api.docx"
    assert cli_main(["tex2docx", str(source), "-o", str(via_cli), "--force"]) == 0
    convert(source, via_api, force=True)
    assert _docx_view(via_cli) == _docx_view(via_api)


def test_generic_final_filename_is_accepted(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    source.write_text("Hello\n", encoding="utf-8")
    output = tmp_path / "paper_final.docx"
    assert cli_main(["md2docx", str(source), "-o", str(output), "--force"]) == 0
    assert output.is_file()


def test_markdown_latex_round_trip_projection(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    source.write_text(
        "# Methods\n\nWater is **bold** and $x_i$ is inline.\n\n- one\n- two\n\n"
        "$$\nE = mc^2\n$$\n\nSee \\cite{smith2020}.\n",
        encoding="utf-8",
    )
    tex = tmp_path / "sample.tex"
    back = tmp_path / "back.md"
    convert(source, tex, force=True)
    convert(tex, back, force=True)
    assert read(source).document.projection() == read(back).document.projection()


def test_preserve_rejects_cross_format_and_unsupported_tex(tmp_path: Path) -> None:
    docx = tmp_path / "sample.docx"
    Document().save(docx)
    with pytest.raises(ConversionError, match="preserve fidelity"):
        convert(docx, tmp_path / "out.md", fidelity="preserve")
    assert not (tmp_path / "out.md").exists()

    source = tmp_path / "custom.tex"
    source.write_text(
        "\\documentclass{article}\n\\begin{document}\n\\foobar{hidden}\n\\end{document}\n",
        encoding="utf-8",
    )
    output = tmp_path / "custom.md"
    with pytest.raises(ConversionError, match="loss policy"):
        convert(source, output, fidelity="preserve")
    assert not output.exists()


def test_unchanged_docx_is_byte_identical(tmp_path: Path) -> None:
    source = tmp_path / "sample.docx"
    document = Document()
    document.add_paragraph("Unchanged sentence.")
    document.save(source)
    output = tmp_path / "copy.docx"
    before = source.read_bytes()
    result = convert(source, output, fidelity="preserve", force=True)
    assert result.execution_mode == "EXACT_REUSE"
    assert output.read_bytes() == before
    assert source.read_bytes() == before


def test_write_unchanged_markdown_matches_direct_render(tmp_path: Path) -> None:
    source = tmp_path / "sample.md"
    source.write_text("Plain sentence.\n", encoding="utf-8")
    loaded = read(source)
    via_write = tmp_path / "write.docx"
    via_convert = tmp_path / "convert.docx"
    write(loaded, via_write, force=True)
    convert(source, via_convert, force=True)
    assert _docx_view(via_write) == _docx_view(via_convert)


def test_style_profile_yaml_and_restyle_keeps_text(tmp_path: Path) -> None:
    profile_path = tmp_path / "style.yaml"
    profile_path.write_text(
        "schema: docforge.style.v1\n"
        "roles:\n"
        "  body:\n"
        "    word_style: Normal\n"
        "    font:\n"
        "      latin: Times New Roman\n"
        "      east_asia: SimSun\n"
        "      size_pt: 14\n"
        "    paragraph:\n"
        "      alignment: justify\n"
        "      line_spacing: 1.15\n"
        "  caption:\n"
        "    word_style: Caption\n"
        "    font: {size_pt: 9}\n"
        "page:\n"
        "  size: A4\n"
        "  margins_mm: {top: 25, bottom: 25, left: 20, right: 20}\n"
        "policy:\n"
        "  missing_style: create\n"
        "  direct_formatting: clear_conflicting\n",
        encoding="utf-8",
    )
    profile = load_profile(profile_path)
    assert profile.roles["body"].font.size_pt == 14
    assert profile.roles["body"].font.east_asia == "SimSun"
    source = _fixture(tmp_path / "styled.docx")
    output = tmp_path / "restyled.docx"
    again = tmp_path / "again.docx"
    summary = restyle_docx(source, output, profile, force=True)
    assert summary["content_unchanged"] is True
    assert _texts(source) == _texts(output)
    resolved = resolved_paragraph_styles(output)
    body = next(item for item in resolved if item["text"].startswith("Body"))
    caption = next(item for item in resolved if item["text"] == "Caption text")
    assert body["size_pt"] == 14
    assert body["latin"] == "Times New Roman"
    assert body["east_asia"] == "SimSun"
    assert body["italic"] is True
    assert body["subscript"] is True
    assert _page_size(output) == ("11906", "16838")
    assert caption["size_pt"] == 9
    assert caption["text"] == "Caption text"
    restyle_docx(output, again, profile, force=True)
    assert resolved_paragraph_styles(output) == resolved_paragraph_styles(again)
    assert cli_main(["restyle", str(source), str(tmp_path / "cli.docx"), "--profile", str(profile_path), "--force"]) == 0


def test_direct_formatting_policies(tmp_path: Path) -> None:
    source = _fixture(tmp_path / "direct.docx")
    preserved = tmp_path / "preserved.docx"
    cleared = tmp_path / "cleared.docx"
    profile = {
        "schema": "docforge.style.v1",
        "roles": {"body": {"word_style": "Normal", "font": {"size_pt": 16}}},
        "policy": {"direct_formatting": "preserve", "missing_style": "error"},
    }
    restyle_docx(source, preserved, profile, force=True)
    body = next(item for item in resolved_paragraph_styles(preserved) if item["text"].startswith("Body"))
    assert body["direct_size"] is True
    assert body["size_pt"] == 20
    profile["policy"]["direct_formatting"] = "clear_all"
    restyle_docx(source, cleared, profile, force=True)
    body = next(item for item in resolved_paragraph_styles(cleared) if item["text"].startswith("Body"))
    assert body["direct_italic"] is False
    assert body["subscript"] is True
    assert body["size_pt"] == 16


def test_style_inheritance_and_missing_role(tmp_path: Path) -> None:
    source = _fixture(tmp_path / "inherit.docx")
    with pytest.raises(ConversionError, match="no Word style"):
        restyle_docx(source, tmp_path / "nope.docx", {
            "schema": "docforge.style.v1",
            "roles": {"missing": {"word_style": "NotAStyle", "font": {"size_pt": 12}}},
            "policy": {"missing_style": "error"},
        })
    assert not (tmp_path / "nope.docx").exists()
    created = tmp_path / "created.docx"
    restyle_docx(source, created, {
        "schema": "docforge.style.v1",
        "roles": {"missing": {"word_style": "NotAStyle", "font": {"size_pt": 12}}},
        "policy": {"missing_style": "create"},
    }, force=True)
    assert created.is_file()


def _docx_view(path: Path) -> list[tuple[str, list[tuple[str, bool, bool]]]]:
    document = Document(path)
    return [
        (paragraph.text, [(run.text, bool(run.bold), bool(run.italic)) for run in paragraph.runs])
        for paragraph in document.paragraphs
    ]


def _page_size(path: Path) -> tuple[str | None, str | None]:
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    size = root.find(".//w:sectPr/w:pgSz", namespace)
    assert size is not None
    return size.get(f"{{{namespace['w']}}}w"), size.get(f"{{{namespace['w']}}}h")


def _texts(path: Path) -> tuple[str, ...]:
    return tuple(paragraph.text for paragraph in Document(path).paragraphs)


def _fixture(path: Path) -> Path:
    document = Document()
    caption_style = document.styles["Caption"]
    caption_style.font.size = Pt(9)
    body = document.add_paragraph()
    body.style = document.styles["Normal"]
    lead = body.add_run("Body ")
    lead.font.size = Pt(20)
    term = body.add_run("term")
    term.italic = True
    term.font.size = Pt(20)
    subscript = body.add_run("2")
    subscript.font.subscript = True
    subscript.font.size = Pt(20)
    caption = document.add_paragraph("Caption text")
    caption.style = caption_style
    document.save(path)
    return path
