"""OOXML package plumbing: load/save a DOCX as an in-memory zip of parts."""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree


class Package:
    """A DOCX opened as a dict of part-name → bytes, with XML helpers."""

    def __init__(self, parts: dict[str, bytes]) -> None:
        self.parts = parts

    @classmethod
    def load(cls, path: Path) -> "Package":
        with zipfile.ZipFile(path) as archive:
            return cls({name: archive.read(name) for name in archive.namelist()})

    def xml(self, name: str) -> etree._Element:
        return etree.fromstring(self.parts[name])

    def set_xml(self, name: str, root: etree._Element) -> None:
        self.parts[name] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def write(self, path: Path, overwrite: bool = False) -> None:
        if path.exists() and not overwrite:
            raise FileExistsError(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, payload in self.parts.items():
                archive.writestr(name, payload)


__all__ = ["Package"]