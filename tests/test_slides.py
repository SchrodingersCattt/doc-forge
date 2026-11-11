from __future__ import annotations

import json
import zipfile
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

import pytest

from docforge.cli import main as cli_main
from docforge.slides import write_pptx

_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

_DECK = [
    {"title": "Alpha & Beta", "bullets": ["first point", "../outside.xml"]},
    {"title": "Second", "bullets": ["<kept>"]},
]


def _inside(name: str) -> None:
    path = PurePosixPath(name)
    assert not path.is_absolute()
    assert ".." not in path.parts
    assert "\\" not in name
    assert ":" not in name


def test_write_pptx_builds_two_visible_slides(tmp_path) -> None:
    path = tmp_path / "deck.pptx"
    write_pptx(path, _DECK)
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        for name in names:
            _inside(name)
        slide_names = [name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")]
        assert slide_names == ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]
        raw = "\n".join(archive.read(name).decode("utf-8") for name in slide_names)
        assert "Alpha &amp; Beta" in raw
        assert "Second" in raw
        assert "first point" in raw
        assert "../outside.xml" in raw
        assert "&lt;kept&gt;" in raw
        assert "<kept>" not in raw
        assert "outside.xml" not in names
        slides = [ET.fromstring(archive.read(name)) for name in slide_names]
        texts = [node.text for slide in slides for node in slide.iter(f"{{{_A}}}t")]
        assert "Alpha & Beta" in texts
        assert "Second" in texts
        assert "first point" in texts
        offsets = list(slides[0].iter(f"{{{_A}}}off"))
        extents = list(slides[0].iter(f"{{{_A}}}ext"))
        assert offsets and extents
        assert any(int(node.attrib["x"]) > 0 and int(node.attrib["y"]) > 0 for node in offsets)
        assert any(int(node.attrib["cx"]) > 0 and int(node.attrib["cy"]) > 0 for node in extents)
        types = ET.fromstring(archive.read("[Content_Types].xml"))
        for override in types.findall(f"{{{_CT}}}Override"):
            part = override.attrib["PartName"].lstrip("/")
            _inside(part)
            assert part in names
        for name in names:
            if not name.endswith(".rels"):
                continue
            rels = ET.fromstring(archive.read(name))
            for rel in rels.findall(f"{{{_REL}}}Relationship"):
                target = rel.attrib["Target"]
                assert ".." not in PurePosixPath(target).parts
                assert target.startswith("/")
                member = target.lstrip("/")
                _inside(member)
                assert member in names
    assert not (tmp_path / "outside.xml").exists()
    assert not list(tmp_path.rglob("*.partial"))


def test_second_write_requires_force(tmp_path) -> None:
    path = tmp_path / "deck.pptx"
    write_pptx(path, [{"title": "Kept", "bullets": ["stay"]}])
    original = path.read_bytes()
    with pytest.raises(FileExistsError, match="--force"):
        write_pptx(path, [{"title": "Replaced", "bullets": ["no"]}])
    assert path.read_bytes() == original
    write_pptx(path, [{"title": "Replaced", "bullets": ["yes"]}], force=True)
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "Replaced" in xml
    assert "Kept" not in xml


def test_empty_title_is_refused(tmp_path) -> None:
    path = tmp_path / "deck.pptx"
    for title in ("", "   "):
        with pytest.raises(ValueError, match="title"):
            write_pptx(path, [{"title": title, "bullets": []}])
    assert not path.exists()


def test_cli_slides_command(tmp_path, capsys) -> None:
    source = tmp_path / "slides.json"
    source.write_text(json.dumps(_DECK), encoding="utf-8")
    output = tmp_path / "deck.pptx"
    assert cli_main(["slides", str(source), "-o", str(output)]) == 0
    assert output.is_file()
    assert str(output) in capsys.readouterr().out
    assert cli_main(["slides", str(source), "-o", str(output)]) == 2
    assert cli_main(["slides", str(source), "-o", str(output), "--force"]) == 0
