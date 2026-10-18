from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from docforge.output import (
    INVALID_FORMAT,
    UNSUPPORTED_FORMAT,
    InvalidFormatError,
    UnsupportedFormatError,
    _detect_format,
)
from docforge.cli import main as cli_main


def _package(path: Path, part: str) -> Path:
    with ZipFile(path, "w") as archive:
        archive.writestr(part, "")
    return path


def test_detect_format_uses_package_magic_and_extension(tmp_path: Path) -> None:
    source = _package(tmp_path / "source.docx", "word/document.xml")
    assert _detect_format(source) == "docx"
    assert _detect_format(source).name == "docx"

    renamed = _package(tmp_path / "source.txt", "word/document.xml")
    with pytest.raises(InvalidFormatError) as error:
        _detect_format(renamed)
    assert error.value.code == INVALID_FORMAT


@pytest.mark.parametrize(
    ("filename", "part"),
    [("slides.docx", "ppt/presentation.xml"), ("sheet.docx", "xl/workbook.xml")],
)
def test_detect_format_rejects_non_word_ooxml(tmp_path: Path, filename: str, part: str) -> None:
    source = _package(tmp_path / filename, part)
    with pytest.raises(UnsupportedFormatError) as error:
        _detect_format(source)
    assert error.value.code == UNSUPPORTED_FORMAT
    assert str(tmp_path) not in str(error.value)


def test_detect_format_rejects_ole2_without_parser_text(tmp_path: Path) -> None:
    source = tmp_path / "legacy.doc"
    source.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64)
    with pytest.raises(UnsupportedFormatError) as error:
        _detect_format(source)
    assert error.value.code == UNSUPPORTED_FORMAT
    assert "parser" not in str(error.value).casefold()


def test_cli_roundtrip_rejects_before_creating_workdir(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    source = _package(tmp_path / "slides.docx", "ppt/presentation.xml")
    workdir = tmp_path / "work"
    output = tmp_path / "out.docx"
    assert cli_main(["roundtrip", str(source), "--workdir", str(workdir), "--output", str(output)]) == 2
    assert not workdir.exists()
    assert not output.exists()
    error = capsys.readouterr().err
    assert "UNSUPPORTED_FORMAT" in error
    assert str(source) not in error
