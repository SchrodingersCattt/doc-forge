from __future__ import annotations

import errno
import json
import os
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

import pytest

from docforge.cli import main as cli_main
from docforge.slides import write_pptx

_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

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
    assert path.is_file()
    assert not path.is_symlink()
    if os.name == "posix":
        assert path.stat().st_mode & 0o777 == 0o600
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
    assert not output.is_symlink()
    assert str(output) in capsys.readouterr().out
    assert cli_main(["slides", str(source), "-o", str(output)]) == 2
    assert cli_main(["slides", str(source), "-o", str(output), "--force"]) == 0


def test_title_box_and_body_box_are_the_only_shapes(tmp_path) -> None:
    path = tmp_path / "deck.pptx"
    write_pptx(path, [
        {"title": "Title", "bullets": ["one", "two", "three", "four"]},
        {"title": "Only title", "bullets": []},
    ])
    with zipfile.ZipFile(path) as archive:
        first = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        second = ET.fromstring(archive.read("ppt/slides/slide2.xml"))
    filled = _shapes(first)
    empty = _shapes(second)
    assert [shape[0] for shape in filled] == ["Title", "Body"]
    assert [_paragraph_text(paragraph) for paragraph in filled[0][1]] == ["Title"]
    assert [_paragraph_text(paragraph) for paragraph in filled[1][1]] == ["one", "two", "three", "four"]
    assert [shape[0] for shape in empty] == ["Title"]
    assert [_paragraph_text(paragraph) for paragraph in empty[0][1]] == ["Only title"]
    for shape in (*filled, *empty):
        assert shape[2].attrib["val"] == "000000"
        assert int(shape[3].attrib["x"]) > 0 and int(shape[3].attrib["y"]) > 0
        assert int(shape[4].attrib["cx"]) > 0 and int(shape[4].attrib["cy"]) > 0


def test_publish_leaves_planted_temp_names_untouched(tmp_path) -> None:
    victim = tmp_path / "victim.txt"
    victim.write_text("keep-me")
    victim_inode = victim.stat().st_ino
    planted_link = tmp_path / ".deck.pptx.partial"
    planted_link.symlink_to(victim)
    planted_file = tmp_path / ".other.pptx.partial"
    planted_file.write_text("keep-file")
    output = tmp_path / "deck.pptx"
    write_pptx(output, [{"title": "T", "bullets": ["b"]}])
    assert victim.read_text() == "keep-me"
    assert victim.stat().st_ino == victim_inode
    assert planted_link.is_symlink()
    assert os.readlink(planted_link) == str(victim)
    assert planted_file.read_text() == "keep-file"
    assert output.is_file() and not output.is_symlink()
    assert zipfile.is_zipfile(output)
    names = {path.name for path in tmp_path.iterdir()}
    assert names == {victim.name, planted_link.name, planted_file.name, output.name}


def test_refused_write_does_not_create_a_temp_file(tmp_path, monkeypatch) -> None:
    output = tmp_path / "deck.pptx"
    write_pptx(output, [{"title": "Kept", "bullets": ["stay"]}])
    original = output.read_bytes()

    def fail_temp(*args, **kwargs):
        raise AssertionError("temporary file created")

    monkeypatch.setattr("docforge.slides.tempfile.mkstemp", fail_temp)
    with pytest.raises(FileExistsError, match="--force"):
        write_pptx(output, [{"title": "Next", "bullets": ["no"]}])
    with pytest.raises(ValueError, match="XML"):
        write_pptx(output, [{"title": "\ufffe", "bullets": ["no"]}])
    directory = tmp_path / "directory"
    directory.mkdir()
    with pytest.raises(ValueError, match="Is a directory"):
        write_pptx(directory, [{"title": "Next", "bullets": []}], force=True)
    assert output.read_bytes() == original
    assert not list(tmp_path.rglob("*.partial"))


def test_failed_replace_keeps_destination_and_removes_temp(tmp_path, monkeypatch) -> None:
    output = tmp_path / "deck.pptx"
    write_pptx(output, [{"title": "Kept", "bullets": ["stay"]}])
    original = output.read_bytes()
    planted = tmp_path / ".deck.pptx.partial"
    planted.write_text("keep-file")

    def fail_replace(source, destination):
        raise PermissionError(13, "Permission denied", "/tmp/.deck.pptx.hidden.partial")

    monkeypatch.setattr("docforge.slides.os.replace", fail_replace)
    with pytest.raises(ValueError, match="Permission denied") as caught:
        write_pptx(output, [{"title": "Next", "bullets": ["no"]}], force=True)
    message = str(caught.value)
    assert str(output) in message
    assert ".partial" not in message
    assert ".hidden" not in message
    assert output.read_bytes() == original
    assert planted.read_text() == "keep-file"
    assert {path.name for path in tmp_path.iterdir()} == {output.name, planted.name}

    def fail_link(source, destination, *args, **kwargs):
        raise PermissionError(13, "Permission denied", "/tmp/.deck.pptx.hidden.partial")

    def replace_is_a_bug(*args, **kwargs):
        raise AssertionError("no-clobber used replace")

    monkeypatch.setattr("docforge.slides.os.link", fail_link)
    monkeypatch.setattr("docforge.slides.os.replace", replace_is_a_bug)
    fresh = tmp_path / "fresh.pptx"
    with pytest.raises(ValueError, match="Permission denied") as caught:
        write_pptx(fresh, [{"title": "Next", "bullets": ["no"]}])
    assert str(fresh) in str(caught.value)
    assert ".partial" not in str(caught.value)
    assert ".hidden" not in str(caught.value)
    assert not fresh.exists()
    assert {path.name for path in tmp_path.iterdir()} == {output.name, planted.name}


def test_symlink_destination_replaces_the_link_not_the_target(tmp_path) -> None:
    target = tmp_path / "target.pptx"
    target.write_bytes(b"original-target")
    target_inode = target.stat().st_ino
    output = tmp_path / "deck.pptx"
    output.symlink_to(target)
    with pytest.raises(FileExistsError, match="--force"):
        write_pptx(output, [{"title": "T", "bullets": ["b"]}])
    assert output.is_symlink()
    assert target.read_bytes() == b"original-target"
    write_pptx(output, [{"title": "T", "bullets": ["b"]}], force=True)
    assert output.is_file() and not output.is_symlink()
    assert zipfile.is_zipfile(output)
    assert target.stat().st_ino == target_inode
    assert target.read_bytes() == b"original-target"

    real_dir = tmp_path / "real-dir"
    real_dir.mkdir()
    (real_dir / "keep.txt").write_text("keep")
    linked_dir = tmp_path / "linked.pptx"
    linked_dir.symlink_to(real_dir, target_is_directory=True)
    with pytest.raises(ValueError, match="Is a directory"):
        write_pptx(linked_dir, [{"title": "T", "bullets": []}], force=True)
    assert linked_dir.is_symlink()
    assert (real_dir / "keep.txt").read_text() == "keep"

    broken = tmp_path / "broken.pptx"
    missing = tmp_path / "missing-target.pptx"
    broken.symlink_to(missing)
    with pytest.raises(FileExistsError, match="--force") as caught:
        write_pptx(broken, [{"title": "Broken", "bullets": ["b"]}])
    assert str(broken) in str(caught.value)
    assert ".partial" not in str(caught.value)
    assert broken.is_symlink()
    assert os.readlink(broken) == str(missing)
    assert not missing.exists()
    write_pptx(broken, [{"title": "Broken", "bullets": ["b"]}], force=True)
    assert broken.is_file() and not broken.is_symlink()
    assert not missing.exists()
    assert zipfile.is_zipfile(broken)


@pytest.mark.parametrize("title,bullet", [
    ("\u0000", "ok"),
    ("\u0008", "ok"),
    ("\u000b", "ok"),
    ("\u000c", "ok"),
    ("\u001f", "ok"),
    ("\ufffe", "ok"),
    ("\uffff", "ok"),
    ("\ud800", "ok"),
    ("\udfff", "ok"),
    ("ok", "\ufffe"),
    ("ok", "\ud800"),
])
def test_xml_illegal_characters_are_refused_before_publish(tmp_path, title, bullet) -> None:
    output = tmp_path / "deck.pptx"
    output.write_bytes(b"keep")
    planted = tmp_path / ".deck.pptx.partial"
    planted.write_text("keep-file")
    with pytest.raises(ValueError, match="XML"):
        write_pptx(output, [{"title": title, "bullets": [bullet]}])
    assert output.read_bytes() == b"keep"
    assert planted.read_text() == "keep-file"
    assert {path.name for path in tmp_path.iterdir()} == {output.name, planted.name}


def test_xml_10_characters_round_trip(tmp_path) -> None:
    title = "A\tB\nC\r \u007f\u00a0\ud7ff\ue000\ufffd"
    bullet = "\U0001f600\U00010000\U0010ffff"
    output = tmp_path / "deck.pptx"
    write_pptx(output, [{"title": title, "bullets": [bullet]}])
    with zipfile.ZipFile(output) as archive:
        raw = archive.read("ppt/slides/slide1.xml")
        slide = ET.fromstring(raw)
    assert b"\r" in raw
    texts = [node.text for node in slide.iter(f"{{{_A}}}t")]
    # An XML parser normalizes a literal CR to LF. The other characters stay.
    assert texts == [title.replace("\r", "\n"), bullet]


def test_output_filesystem_errors_name_the_requested_path(tmp_path, capsys) -> None:
    source = tmp_path / "slides.json"
    source.write_text(json.dumps([{"title": "T", "bullets": ["b"]}]), encoding="utf-8")
    missing = tmp_path / "missing" / "deck.pptx"
    assert cli_main(["slides", str(source), "-o", str(missing)]) == 2
    missing_err = capsys.readouterr().err
    assert f"error: Cannot write {missing}:" in missing_err
    assert "No such file or directory" in missing_err
    assert ".partial" not in missing_err
    assert "Traceback" not in missing_err
    assert not missing.exists()
    assert not missing.parent.exists()

    parent = tmp_path / "not-a-directory"
    parent.write_text("keep")
    nested = parent / "deck.pptx"
    assert cli_main(["slides", str(source), "-o", str(nested)]) == 2
    nested_err = capsys.readouterr().err
    assert f"error: Cannot write {nested}:" in nested_err
    assert "Not a directory" in nested_err
    assert ".partial" not in nested_err
    assert "Traceback" not in nested_err
    assert parent.read_text() == "keep"

    directory = tmp_path / "outdir"
    directory.mkdir()
    marker = directory / "keep.txt"
    marker.write_text("keep")
    assert cli_main(["slides", str(source), "-o", str(directory)]) == 2
    assert cli_main(["slides", str(source), "-o", str(directory), "--force"]) == 2
    directory_err = capsys.readouterr().err
    assert directory_err.count(f"error: Cannot write {directory}:") == 2
    assert "Is a directory" in directory_err
    assert "--force" not in directory_err
    assert ".partial" not in directory_err
    assert "Traceback" not in directory_err
    assert marker.read_text() == "keep"

    absent = tmp_path / "no-input.json"
    assert cli_main(["slides", str(absent), "-o", str(tmp_path / "deck.pptx")]) == 2
    absent_err = capsys.readouterr().err
    assert str(absent) in absent_err
    assert "Cannot write" not in absent_err


def test_cli_illegal_xml_exits_2_without_a_file(tmp_path, capsys) -> None:
    source = tmp_path / "slides.json"
    source.write_text(json.dumps([{"title": "\ufffe", "bullets": ["ok"]}]), encoding="utf-8")
    output = tmp_path / "deck.pptx"
    assert cli_main(["slides", str(source), "-o", str(output)]) == 2
    captured = capsys.readouterr()
    assert "XML" in captured.err
    assert "Traceback" not in captured.err
    assert not output.exists()
    assert not list(tmp_path.rglob("*.partial"))


def test_publish_keeps_file_created_after_preflight(tmp_path, monkeypatch) -> None:
    output = tmp_path / "deck.pptx"
    competitor = b"competitor"
    collisions = []

    def link_after_competitor(source, destination, *args, _link=os.link, **kwargs):
        Path(destination).write_bytes(competitor)
        assert Path(destination).is_file()
        try:
            return _link(source, destination, *args, **kwargs)
        except FileExistsError as exc:
            collisions.append(exc.errno)
            raise

    def replace_is_a_bug(*args, **kwargs):
        raise AssertionError("no-clobber used replace")

    monkeypatch.setattr("docforge.slides.os.link", link_after_competitor)
    monkeypatch.setattr("docforge.slides.os.replace", replace_is_a_bug)
    with pytest.raises(FileExistsError, match="--force") as caught:
        write_pptx(output, [{"title": "T", "bullets": ["b"]}])
    message = str(caught.value)
    assert collisions == [errno.EEXIST]
    assert str(output) in message
    assert ".partial" not in message
    assert output.read_bytes() == competitor
    assert not list(tmp_path.rglob("*.partial"))


def test_publish_keeps_directory_created_after_preflight(tmp_path, monkeypatch) -> None:
    directory = tmp_path / "late-dir"
    collisions = []

    def link_after_directory(source, destination, *args, _link=os.link, **kwargs):
        path = Path(destination)
        path.mkdir()
        assert path.is_dir()
        try:
            return _link(source, destination, *args, **kwargs)
        except FileExistsError as exc:
            collisions.append(exc.errno)
            raise

    def replace_is_a_bug(*args, **kwargs):
        raise AssertionError("no-clobber used replace")

    monkeypatch.setattr("docforge.slides.os.link", link_after_directory)
    monkeypatch.setattr("docforge.slides.os.replace", replace_is_a_bug)
    with pytest.raises(ValueError, match="Is a directory") as caught:
        write_pptx(directory, [{"title": "T", "bullets": ["b"]}])
    assert collisions == [errno.EEXIST]
    assert str(directory) in str(caught.value)
    assert ".partial" not in str(caught.value)
    assert directory.is_dir()
    assert not any(directory.iterdir())
    assert not list(tmp_path.rglob("*.partial"))


def test_exclusive_link_failure_does_not_replace(tmp_path, monkeypatch) -> None:
    """Simulated platforms. These checks do not run on Windows."""

    output = tmp_path / "deck.pptx"

    def unavailable(*args, **kwargs):
        raise NotImplementedError("link")

    def replace_is_a_bug(*args, **kwargs):
        raise AssertionError("unsupported link used replace")

    monkeypatch.setattr("docforge.slides.os.link", unavailable)
    monkeypatch.setattr("docforge.slides.os.replace", replace_is_a_bug)
    with pytest.raises(ValueError, match="atomic publish is unavailable") as caught:
        write_pptx(output, [{"title": "T", "bullets": ["b"]}])
    assert str(output) in str(caught.value)
    assert ".partial" not in str(caught.value)
    assert not output.exists()

    monkeypatch.delattr("docforge.slides.os", "link", raising=False)
    with pytest.raises(ValueError, match="atomic publish is unavailable"):
        write_pptx(output, [{"title": "T", "bullets": ["b"]}])
    assert not output.exists()
    assert not list(tmp_path.rglob("*.partial"))


def test_export_does_not_use_fchmod_or_umask(tmp_path, monkeypatch) -> None:
    """Removes os.fchmod to resemble a module without that Unix API.

    This process is not Windows, and the test does not execute Windows.
    The owner-read mode check runs only when ``os.name`` is ``posix``.
    """

    def umask_is_a_bug(*args, **kwargs):
        raise AssertionError("umask called")

    monkeypatch.setattr("docforge.slides.os.umask", umask_is_a_bug)
    monkeypatch.delattr("docforge.slides.os", "fchmod", raising=False)
    output = tmp_path / "deck.pptx"
    write_pptx(output, [{"title": "T", "bullets": ["b"]}])
    assert output.is_file() and not output.is_symlink()
    if os.name == "posix":
        assert output.stat().st_mode & 0o777 == 0o600
    with zipfile.ZipFile(output) as archive:
        assert b"T" in archive.read("ppt/slides/slide1.xml")


def test_programming_errors_are_not_rewritten(tmp_path, monkeypatch) -> None:
    source = tmp_path / "slides.json"
    source.write_text(json.dumps([{"title": "T", "bullets": []}]), encoding="utf-8")
    output = tmp_path / "deck.pptx"

    def fail_package(*args, **kwargs):
        raise TypeError("programmer")

    monkeypatch.setattr("docforge.slides._package", fail_package)
    with pytest.raises(TypeError, match="programmer"):
        write_pptx(output, [{"title": "T", "bullets": []}])
    with pytest.raises(TypeError, match="programmer"):
        cli_main(["slides", str(source), "-o", str(output)])
    assert not output.exists()


def _shapes(slide: ET.Element) -> list[tuple]:
    shapes = []
    for shape in slide.iter(f"{{{_P}}}sp"):
        name = shape.find(f".//{{{_P}}}cNvPr").attrib["name"]
        paragraphs = list(shape.iter(f"{{{_A}}}p"))
        color = shape.find(f".//{{{_A}}}srgbClr")
        offset = shape.find(f".//{{{_A}}}off")
        extent = shape.find(f".//{{{_A}}}ext")
        shapes.append((name, paragraphs, color, offset, extent))
    return shapes


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(f"{{{_A}}}t"))
