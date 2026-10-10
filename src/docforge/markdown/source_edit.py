"""Apply a Markdown paragraph delta onto a source DOCX.

Template assembly redraws every paragraph and drops layout that Markdown
cannot carry: title-page styles, index leaders, equations, and tables.
This editor leaves the source package untouched except for paragraphs whose
baseline Markdown text uniquely matches a source paragraph and whose edited
Markdown text changed. New paragraphs are inserted beside that anchor and
copy its paragraph and run properties.
"""

from __future__ import annotations

import copy
import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

from lxml import etree

from ..docxdiff.package import Package
from .inline import display, key as inline_key
from ..output import validate_output_path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
TAG_RE = re.compile(r"</?[^>]+>")
ESCAPE_RE = re.compile(r"\\([\\`*{}\[\]()#+\-.!_>~|])")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
LEAD_NUMBER_RE = re.compile(r"^\d+\.\s+")
INLINE_RE = re.compile(
    r"\*\*(.+?)\*\*|\*(.+?)\*|<sub>(.+?)</sub>|<sup>(.+?)</sup>",
    re.DOTALL,
)
DASHES = str.maketrans(
    {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
        "\u00ad": "",
    }
)


def apply_markdown_delta(
    source: Path,
    baseline_markdown: Path,
    edited_markdown: Path,
    output: Path,
    *,
    overwrite: bool = False,
) -> dict[str, int]:
    """Write ``output`` as ``source`` plus the paragraph delta of the Markdown.

    Paragraphs, drawings, tables, and equations that the delta does not name
    stay byte-for-byte in ``word/document.xml`` apart from the inserted
    siblings. A paragraph is editable only when its baseline Markdown block
    matches exactly one source paragraph.
    """

    validate_output_path(output)
    package = Package.load(source)
    root = package.xml("word/document.xml")
    operations = plan_markdown_delta(
        root,
        baseline_markdown.read_text(encoding="utf-8"),
        edited_markdown.read_text(encoding="utf-8"),
        styles_xml=package.parts.get("word/styles.xml"),
    )
    _apply(root, operations)
    package.set_xml("word/document.xml", root)
    package.write(output, overwrite=overwrite)
    summary = {"replaced": 0, "inserted": 0}
    for operation in operations:
        summary["replaced" if operation[0] == "replace" else "inserted"] += 1
    return summary


def plan_markdown_delta(
    root: etree._Element,
    baseline: str,
    edited: str,
    *,
    styles_xml: bytes | None = None,
) -> list[tuple]:
    """Return replace and insert operations against ``root``'s body children."""

    body = root.find(f"{{{W}}}body")
    if body is None:
        raise ValueError("DOCX has no document body")
    children = [child for child in list(body) if etree.QName(child).localname != "sectPr"]
    docx_keys = [(index, _paragraph_key(child)) for index, child in enumerate(children)]
    baseline_blocks = _blocks(baseline)
    edited_blocks = _blocks(edited)
    mapping = _map_blocks(baseline_blocks, docx_keys)
    heading_ids = _heading_style_ids(styles_xml)
    prototypes = _prototypes(children, heading_ids)
    operations: list[tuple] = []
    matcher = SequenceMatcher(
        a=[_opcode_key(block) for block in baseline_blocks],
        b=[_opcode_key(block) for block in edited_blocks],
        autojunk=False,
    )
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "delete":
            raise ValueError("markdown delta deletes a block; refusing to drop source layout")
        new_blocks = edited_blocks[j1:j2]
        if any(block.kind != "text" for block in new_blocks):
            raise ValueError("markdown delta changes a table or image; those stay in the source DOCX")
        if tag == "insert":
            anchor = _anchor_before(baseline_blocks, mapping, i1)
            operations.append(("insert", anchor, new_blocks, prototypes))
            continue
        mapped = [mapping.get(index) for index in range(i1, i2)]
        if not mapped or any(index is None for index in mapped):
            missing = baseline_blocks[i1].raw[:80]
            raise ValueError(f"edited paragraph does not match a unique source paragraph: {missing}")
        operations.append(("replace", mapped, new_blocks, prototypes))
    return operations


class _Block:
    def __init__(self, kind: str, raw: str) -> None:
        self.kind = kind
        self.raw = raw

    @property
    def key(self) -> str:
        return _key(self.raw)


def _blocks(text: str) -> list[_Block]:
    chunks = re.split(r"\n\s*\n", text.replace("\r\n", "\n"))
    blocks: list[_Block] = []
    for chunk in chunks:
        lines = [line.strip() for line in chunk.splitlines() if line.strip()]
        if not lines:
            continue
        if all(line.startswith("|") for line in lines) or (len(lines) == 1 and lines[0].startswith("![")):
            blocks.append(_Block("opaque", "\n".join(lines)))
            continue
        blocks.append(_Block("text", " ".join(lines)))
    return blocks


def _opcode_key(block: _Block) -> str:
    return block.kind + "\0" + block.key


def _key(raw: str) -> str:
    return inline_key(raw)


def _paragraph_key(element: etree._Element) -> str:
    if etree.QName(element).localname != "p":
        return ""
    if element.find(f".//{{{W}}}drawing") is not None:
        return ""
    text = "".join(node.text or "" for node in element.iter(f"{{{W}}}t"))
    return _key(text)


def _map_blocks(blocks: list[_Block], docx_keys: list[tuple[int, str]]) -> dict[int, int]:
    unused = {index: key for index, key in docx_keys if key}
    mapping: dict[int, int] = {}
    for index, block in enumerate(blocks):
        if block.kind != "text" or len(block.key) < 8:
            continue
        exact = [child for child, key in unused.items() if key == block.key]
        if len(exact) == 1:
            mapping[index] = exact[0]
            del unused[exact[0]]
            continue
        if len(exact) > 1:
            continue
        ranked = sorted(
            ((SequenceMatcher(a=block.key, b=key, autojunk=False).ratio(), child) for child, key in unused.items()),
            reverse=True,
        )
        if not ranked:
            continue
        best_ratio, best_child = ranked[0]
        second = ranked[1][0] if len(ranked) > 1 else 0.0
        if best_ratio >= 0.92 and best_ratio - second >= 0.04:
            mapping[index] = best_child
            del unused[best_child]
    return mapping


def _anchor_before(blocks: list[_Block], mapping: dict[int, int], start: int) -> int:
    for index in range(start - 1, -1, -1):
        if index in mapping:
            return mapping[index]
    raise ValueError("inserted Markdown has no matched paragraph before it")


def _heading_style_ids(styles_xml: bytes | None) -> dict[int, str]:
    if not styles_xml:
        return {}
    root = etree.fromstring(styles_xml)
    found: dict[int, str] = {}
    for style in root.findall(f"{{{W}}}style"):
        name = style.find(f"{{{W}}}name")
        if name is None:
            continue
        match = re.fullmatch(r"heading ([1-6])", (name.get(f"{{{W}}}val") or "").strip().lower())
        if match:
            found[int(match.group(1))] = style.get(f"{{{W}}}styleId") or ""
    return {level: style_id for level, style_id in found.items() if style_id}


def _prototypes(children: list[etree._Element], heading_ids: dict[int, str]) -> dict:
    body = next((child for child in children if _paragraph_key(child) and not _style_id(child)), None)
    if body is None:
        body = next((child for child in children if etree.QName(child).localname == "p"), None)
    headings = {}
    for level, style_id in heading_ids.items():
        match = next((child for child in children if _style_id(child) == style_id), None)
        if match is not None:
            headings[level] = match
    return {
        "body": body,
        "headings": headings,
        "heading_ids": heading_ids,
        "list_headings": _list_headings(children),
    }


def _style_id(element: etree._Element) -> str:
    style = element.find(f"{{{W}}}pPr/{{{W}}}pStyle")
    return style.get(f"{{{W}}}val") if style is not None else ""


def _list_headings(children: list[etree._Element]) -> dict[int, etree._Element]:
    """Numbered bold paragraphs keyed by numbering level.

    Documents without Heading styles still encode section and subsection
    formatting in those paragraphs. The first example of each level is kept.
    """

    found: dict[int, etree._Element] = {}
    for child in children:
        numbering = child.find(f"{{{W}}}pPr/{{{W}}}numPr")
        if numbering is None or child.find(f"{{{W}}}r/{{{W}}}rPr/{{{W}}}b") is None:
            continue
        if child.find(f"{{{W}}}r/{{{W}}}rPr/{{{W}}}highlight") is not None:
            continue
        level_node = numbering.find(f"{{{W}}}ilvl")
        level = int(level_node.get(f"{{{W}}}val") or "0") if level_node is not None else 0
        found.setdefault(level, child)
    return found


def _apply(root: etree._Element, operations: list[tuple]) -> None:
    body = root.find(f"{{{W}}}body")
    children = [child for child in list(body) if etree.QName(child).localname != "sectPr"]
    tails = {index: child for index, child in enumerate(children)}
    for operation in operations:
        if operation[0] == "insert":
            anchor = operation[1]
            tails[anchor] = _insert_after(tails[anchor], operation[2], operation[3])
            continue
        indexes, new_blocks, prototypes = operation[1], operation[2], operation[3]
        first = children[indexes[0]]
        _write_paragraph(first, new_blocks[0].raw, prototypes, replacement=first)
        for extra in indexes[1:]:
            parent = children[extra].getparent()
            if parent is not None:
                parent.remove(children[extra])
        cursor = first
        for block in new_blocks[1:]:
            cursor = _insert_after(cursor, [block], prototypes)
        tails[indexes[0]] = cursor


def _insert_after(anchor: etree._Element, blocks: list[_Block], prototypes: dict) -> etree._Element:
    cursor = anchor
    for block in blocks:
        paragraph = etree.Element(f"{{{W}}}p")
        properties = _properties_for(block.raw, prototypes, anchor)
        if properties is not None:
            paragraph.append(properties)
        heading = _is_heading(block.raw)
        run_source = _heading_prototype(block.raw, prototypes) if heading else None
        if run_source is None:
            run_source = anchor
        _add_runs(
            paragraph,
            _display_text(block.raw),
            _run_properties(run_source),
            extra_marks={"b"} if heading else frozenset(),
        )
        cursor.addnext(paragraph)
        cursor = paragraph
    return cursor


def _write_paragraph(paragraph: etree._Element, raw: str, prototypes: dict, *, replacement: etree._Element) -> None:
    for child in list(paragraph):
        if etree.QName(child).localname != "pPr":
            paragraph.remove(child)
    _add_runs(paragraph, _display_text(raw), _run_properties(replacement))


def _is_heading(raw: str) -> bool:
    return HEADING_RE.match(raw.strip()) is not None


def _heading_prototype(raw: str, prototypes: dict) -> etree._Element | None:
    match = HEADING_RE.match(raw.strip())
    if match is None:
        return None
    level = min(len(match.group(1)), 6)
    headings = prototypes["headings"]
    if level in headings:
        return headings[level]
    if 1 in headings:
        return headings[1]
    lists = prototypes.get("list_headings") or {}
    # Markdown ## is a section; ### is the subsection under it.
    target = 0 if level <= 2 else level - 2
    if target in lists:
        return lists[target]
    if lists:
        return lists[min(lists)]
    return None


def _properties_for(raw: str, prototypes: dict, anchor: etree._Element) -> etree._Element | None:
    heading = HEADING_RE.match(raw.strip())
    if heading:
        level = min(len(heading.group(1)), 6)
        prototype = _heading_prototype(raw, prototypes)
        if prototype is not None and (level in prototypes["headings"] or 1 in prototypes["headings"]):
            return _copy_paragraph_properties(prototype)
        if prototype is not None:
            return _copy_paragraph_properties(prototype, drop_style=True)
        style_id = prototypes["heading_ids"].get(level) or prototypes["heading_ids"].get(1)
        if style_id:
            properties = etree.Element(f"{{{W}}}pPr")
            etree.SubElement(properties, f"{{{W}}}pStyle").set(f"{{{W}}}val", style_id)
            return properties
    source = anchor if _paragraph_key(anchor) else prototypes["body"]
    if source is None:
        return None
    return _copy_paragraph_properties(source)


def _copy_paragraph_properties(
    paragraph: etree._Element,
    *,
    drop_style: bool = False,
) -> etree._Element | None:
    properties = paragraph.find(f"{{{W}}}pPr")
    if properties is None:
        return None
    copied = copy.deepcopy(properties)
    for numbering in copied.findall(f"{{{W}}}numPr"):
        copied.remove(numbering)
    if drop_style:
        for style in copied.findall(f"{{{W}}}pStyle"):
            copied.remove(style)
    return copied


def _display_text(raw: str) -> str:
    heading = HEADING_RE.match(raw.strip())
    text = heading.group(2).strip() if heading else raw
    return text


def _run_properties(paragraph: etree._Element | None) -> etree._Element | None:
    if paragraph is None:
        return None
    properties = paragraph.find(f"{{{W}}}r/{{{W}}}rPr")
    return copy.deepcopy(properties) if properties is not None else None


def _add_runs(
    paragraph: etree._Element,
    text: str,
    base_properties: etree._Element | None,
    extra_marks: frozenset[str] = frozenset(),
) -> None:
    for run in display(text):
        marks = set(extra_marks)
        if run.bold:
            marks.add("b")
        if run.italic:
            marks.add("i")
        if run.subscript:
            marks.add("sub")
        elif run.superscript:
            marks.add("sup")
        _append_run(paragraph, run.text, base_properties, marks)


def _append_run(paragraph: etree._Element, text: str, base_properties: etree._Element | None, marks: set[str]) -> None:
    if not text:
        return
    run = etree.SubElement(paragraph, f"{{{W}}}r")
    properties = copy.deepcopy(base_properties) if base_properties is not None else etree.Element(f"{{{W}}}rPr")
    for name in ("b", "i", "vertAlign"):
        for node in properties.findall(f"{{{W}}}{name}"):
            properties.remove(node)
    if "b" in marks:
        etree.SubElement(properties, f"{{{W}}}b")
    if "i" in marks:
        etree.SubElement(properties, f"{{{W}}}i")
    if "sub" in marks or "sup" in marks:
        align = etree.SubElement(properties, f"{{{W}}}vertAlign")
        align.set(f"{{{W}}}val", "subscript" if "sub" in marks else "superscript")
    if len(properties):
        run.append(properties)
    node = etree.SubElement(run, f"{{{W}}}t")
    if text.startswith(" ") or text.endswith(" "):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
