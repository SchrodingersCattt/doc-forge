"""Color the accepted revision text of a DOCX comparison without Word tracking."""

from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path

from lxml import etree

from .package import Package
from .redline import M, NS, W, _accepted_revision_view, _final_blocks, create_tracked_docx, visible_text
from ..output import validate_output_path


def _paragraph_color(text: str, rules: list[tuple[re.Pattern[str], str]], fallback: str) -> str:
    matched = {color for pattern, color in rules if pattern.search(text)}
    return matched.pop() if len(matched) == 1 else fallback


def _set_color(run: etree._Element, value: str) -> None:
    props = run.find(f"{{{W}}}rPr")
    if props is None:
        props = etree.Element(f"{{{W}}}rPr")
        run.insert(0, props)
    color = props.find(f"{{{W}}}color")
    if color is None:
        color = etree.SubElement(props, f"{{{W}}}color")
    color.attrib.clear()
    color.set(f"{{{W}}}val", value)


def _rules(path: Path | None) -> list[tuple[re.Pattern[str], str]]:
    if path is None:
        return []
    records = json.loads(path.read_text(encoding="utf-8"))
    return [(re.compile(row["pattern"]), row["color"].lstrip("#").upper()) for row in records]


def create_colored_docx(
    base: Path, current: Path, output: Path, *, color_map_path: Path | None = None,
    default_color: str = "FF0000", overwrite: bool = False,
) -> dict[str, int]:
    """Color only inserted/modified current text; keep all clean-document layout.

    A temporary native comparison supplies the text alignment. Its deleted
    material is discarded, and only the colored current runs are copied into
    the original clean package; paragraph, table and section layout stays put.
    """
    validate_output_path(output)
    if output.exists() and not overwrite:
        raise FileExistsError(output)
    fallback = default_color.lstrip("#").upper()
    rules = _rules(color_map_path)
    if any(not re.fullmatch(r"[0-9A-F]{6}", color) for color in [fallback, *(v for _, v in rules)]):
        raise ValueError("Color must have six hexadecimal digits")
    with tempfile.TemporaryDirectory() as directory:
        redlined = Path(directory) / "diff.docx"
        summary = create_tracked_docx(base, current, redlined, overwrite=True,
                                      preserve_base_revisions=False, workers=1)
        difference = Package.load(redlined).xml("word/document.xml")
        changed_runs = 0
        body = difference.find(".//w:body", NS)
        caption = ""
        for block in body:
            if block.tag == f"{{{W}}}p":
                text = visible_text(block)
                caption = text if re.match(r"^(?:Table|Figure) S?\d+[.]", text) else ""
                paragraphs = [(block, text)]
            elif block.tag == f"{{{W}}}tbl":
                paragraphs = [(p, caption + " " + visible_text(p)) for p in block.findall(".//w:p", NS)]
                caption = ""
            else:
                continue
            for paragraph, context in paragraphs:
                color = _paragraph_color(context, rules, fallback)
                for insertion in paragraph.findall(".//w:ins", NS) + paragraph.findall(".//w:moveTo", NS):
                    for run in (*insertion.findall(".//w:r", NS), *insertion.findall(f".//{{{M}}}r")):
                        _set_color(run, color)
                        changed_runs += 1
                for run in paragraph.xpath(".//w:r[w:rPr/w:rPrChange]", namespaces=NS):
                    _set_color(run, color)
                    changed_runs += 1
                row = paragraph.getparent()
                while row is not None and row.tag != f"{{{W}}}tr":
                    row = row.getparent()
                if row is not None and row.find("./w:trPr/w:ins", NS) is not None:
                    for run in (*paragraph.findall(".//w:r", NS), *paragraph.findall(f".//{{{M}}}r")):
                        _set_color(run, color)
                        changed_runs += 1
        accepted = _accepted_revision_view(difference)
        package = Package.load(current)
        clean = package.xml("word/document.xml")
        if _final_blocks(accepted) != _final_blocks(clean):
            raise AssertionError("Colored diff does not reproduce the revised document text")
        colored_paragraphs = accepted.findall(".//w:body//w:p", NS)
        clean_paragraphs = clean.findall(".//w:body//w:p", NS)
        if len(colored_paragraphs) != len(clean_paragraphs):
            raise AssertionError("Colored diff changed the paragraph count")
        for colored, original in zip(colored_paragraphs, clean_paragraphs):
            if visible_text(colored) != visible_text(original):
                raise AssertionError("Colored diff changed paragraph order or text")
            properties = original.find(f"{{{W}}}pPr")
            for child in list(original):
                if child is not properties:
                    original.remove(child)
            for child in colored:
                if child.tag != f"{{{W}}}pPr":
                    original.append(child)
        package.set_xml("word/document.xml", clean)
        package.write(output, overwrite=overwrite)
    return {**summary, "colored_runs": changed_runs}
