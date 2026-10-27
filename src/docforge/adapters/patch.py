"""SOURCE_PATCH for supported semantic edits of a DOCX-origin document.

Supported edits are confined to simple body paragraphs and headings:
replacing their text, inserting a plain paragraph, or deleting one.
Tables, figures, fields, math, revisions, drawings, and story parts are
refused. Preservation mode must not rebuild the package instead.
"""

from __future__ import annotations

import copy
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from ..adapters.protocol import ConversionError
from ..semantic import BlockNode, InlineNode, SemanticDocument


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML = "http://www.w3.org/XML/1998/namespace"
_BODY_ANCHOR = re.compile(r"^word/document\.xml:child:(\d+)$")
_SUPPORTED_INLINE = frozenset({"text", "emph", "strong"})
_HAZARD = frozenset({
    "drawing", "pict", "object", "hyperlink", "fldChar", "instrText", "fldSimple",
    "ins", "del", "moveFrom", "moveTo", "sdt", "tbl", "sectPr", "oMath", "oMathPara",
    "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd",
    "br", "tab", "cr", "sym", "ruby", "footnoteReference", "endnoteReference",
    "commentReference", "AlternateContent",
})


def patch_docx(source: Path, output: Path, edited: SemanticDocument, *, force: bool = False) -> dict[str, int | str]:
    """Apply a supported semantic edit to ``source`` and publish ``output``.

    The destination is replaced only after the package audit passes. A refused
    edit leaves it untouched.
    """

    source, output = Path(source), Path(output)
    if source.resolve() == output.resolve():
        raise ConversionError("patch output must be a different file", code="PATH_COLLISION")
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    from .docx import project_docx

    original, _diagnostics, _resources, _features, _opaque = project_docx(source)
    operations = _plan(original, edited)
    with zipfile.ZipFile(source) as package:
        document_xml = package.read("word/document.xml")
        original_parts = {name: package.read(name) for name in package.namelist()}
    root = etree.fromstring(document_xml)
    body = root.find(f"{{{W}}}body")
    if body is None:
        raise ConversionError("DOCX has no document body", code="DOCX_INVALID")
    children = list(body)
    _apply(body, children, operations)
    replacement = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{output.stem}-", suffix=".docx", dir=output.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(temporary, "w") as outgoing:
            for info in incoming.infolist():
                payload = replacement if info.filename == "word/document.xml" else incoming.read(info.filename)
                outgoing.writestr(info, payload)
        with zipfile.ZipFile(temporary) as package:
            changed = [name for name in original_parts if package.read(name) != original_parts[name]]
        if changed != ["word/document.xml"]:
            raise ConversionError("patch changed a part other than the document body", code="PATCH_UNCONFINED")
        from .docx import audit_docx

        validation = audit_docx(temporary)
        if not validation.ok:
            raise ConversionError("patch audit failed", code="AUDIT_FAILED", diagnostics=validation.diagnostics)
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
    replaced = sum(1 for item in operations if item[0] == "replace")
    inserted = sum(1 for item in operations if item[0] == "insert")
    deleted = sum(1 for item in operations if item[0] == "delete")
    return {"execution_mode": "SOURCE_PATCH", "replaced": replaced, "inserted": inserted, "deleted": deleted}


def _plan(original: SemanticDocument, edited: SemanticDocument) -> list[tuple]:
    if _resource_key(original) != _resource_key(edited) or original.references != edited.references:
        _refuse("resources or references changed")
    original_groups = _groups(original.blocks)
    edited_groups = _groups(edited.blocks)
    original_anchors = [anchor for anchor, _blocks in original_groups if anchor]
    if len(original_anchors) != len(set(original_anchors)):
        _refuse("source anchors are not unique")
    by_anchor = {anchor: blocks for anchor, blocks in original_groups if anchor}
    edited_anchors = [anchor for anchor, _blocks in edited_groups if anchor]
    unknown = [anchor for anchor in edited_anchors if anchor not in by_anchor]
    if unknown or len(edited_anchors) != len(set(edited_anchors)):
        _refuse("edited anchors do not match the source")
    if edited_anchors != [anchor for anchor in original_anchors if anchor in set(edited_anchors)]:
        _refuse("paragraphs were reordered")
    operations: list[tuple] = []
    for anchor, blocks in edited_groups:
        if not anchor:
            if len(blocks) != 1 or not _insertable(blocks[0]):
                _refuse("only a plain paragraph can be inserted")
            operations.append(("insert", blocks[0]))
            continue
        current = by_anchor[anchor]
        if _same(current, blocks):
            operations.append(("keep", anchor))
            continue
        if len(current) != 1 or len(blocks) != 1:
            _refuse("a grouped source block cannot be edited")
        left, right = current[0], blocks[0]
        if left.kind != right.kind or left.kind not in {"paragraph", "heading"}:
            _refuse(f"{left.kind} edits are not a source patch")
        if left.kind == "heading" and left.attr("level") != right.attr("level"):
            _refuse("heading level changes are not a source patch")
        if not _BODY_ANCHOR.match(anchor) or not _inlines_supported(right.inlines):
            _refuse("this paragraph edit is not a source patch")
        operations.append(("replace", anchor, right))
    edited_anchor_set = set(edited_anchors)
    for anchor in original_anchors:
        if anchor in edited_anchor_set:
            continue
        group = by_anchor[anchor]
        if len(group) != 1 or group[0].kind not in {"paragraph", "heading"} or not _BODY_ANCHOR.match(anchor):
            _refuse("this deletion is not a source patch")
        operations.append(("delete", anchor))
    if not any(item[0] in {"replace", "insert", "delete"} for item in operations):
        _refuse("the semantic edit has no supported paragraph change")
    return operations


def _apply(body: etree._Element, children: list[etree._Element], operations: list[tuple]) -> None:
    def element_for(anchor: str) -> etree._Element:
        match = _BODY_ANCHOR.match(anchor)
        if match is None:
            _refuse("the edit is outside the document body")
        index = int(match.group(1))
        if index >= len(children):
            _refuse("the source anchor is missing from the package")
        return children[index]

    for kind, anchor, *_rest in [(item[0], item[1], *item[2:]) for item in operations if item[0] in {"replace", "delete"}]:
        if not _simple_paragraph(element_for(anchor)):
            _refuse("the source paragraph is not a supported patch target")
    pending: list[BlockNode] = []
    for item in operations:
        if item[0] == "insert":
            pending.append(item[1])
            continue
        if item[0] == "delete":
            continue
        element = element_for(item[1])
        for block in pending:
            element.addprevious(_new_paragraph(block, _prototype(element)))
        pending.clear()
        if item[0] == "replace":
            _rewrite_paragraph(element, item[2].inlines)
    host = body.find(f"{{{W}}}sectPr")
    prototype = _prototype(host) if host is not None else _prototype(body)
    for block in pending:
        paragraph = _new_paragraph(block, prototype)
        if host is not None:
            host.addprevious(paragraph)
        else:
            body.append(paragraph)
    for item in operations:
        if item[0] == "delete":
            body.remove(element_for(item[1]))


def _new_paragraph(block: BlockNode, prototype: etree._Element | None) -> etree._Element:
    paragraph = etree.Element(f"{{{W}}}p")
    if prototype is not None and etree.QName(prototype).localname == "p":
        properties = prototype.find(f"{{{W}}}pPr")
        if properties is not None:
            paragraph.append(copy.deepcopy(properties))
        first = prototype.find(f"{{{W}}}r")
        if first is not None and first.find(f"{{{W}}}rPr") is not None:
            run = etree.SubElement(paragraph, f"{{{W}}}r")
            run.append(copy.deepcopy(first.find(f"{{{W}}}rPr")))
    _rewrite_paragraph(paragraph, block.inlines)
    return paragraph


def _rewrite_paragraph(paragraph: etree._Element, inlines: tuple[InlineNode, ...]) -> None:
    properties = paragraph.find(f"{{{W}}}pPr")
    first = paragraph.find(f"{{{W}}}r")
    base = copy.deepcopy(first.find(f"{{{W}}}rPr")) if first is not None and first.find(f"{{{W}}}rPr") is not None else None
    for child in list(paragraph):
        if child is not properties:
            paragraph.remove(child)
    runs = _marked_runs(inlines) or [("", False, False)]
    for text, bold, italic in runs:
        run = etree.SubElement(paragraph, f"{{{W}}}r")
        properties_copy = copy.deepcopy(base) if base is not None else None
        if properties_copy is None:
            properties_copy = etree.Element(f"{{{W}}}rPr")
        _set_mark(properties_copy, "b", bold)
        _set_mark(properties_copy, "i", italic)
        if len(properties_copy):
            run.append(properties_copy)
        node = etree.SubElement(run, f"{{{W}}}t")
        if text[:1].isspace() or text[-1:].isspace():
            node.set(f"{{{XML}}}space", "preserve")
        node.text = text


def _prototype(element: etree._Element | None) -> etree._Element | None:
    current = element
    while current is not None:
        if etree.QName(current).localname == "p" and _simple_paragraph(current):
            return current
        current = current.getprevious()
    return None


def _simple_paragraph(element: etree._Element) -> bool:
    if etree.QName(element).localname != "p":
        return False
    for node in element.iter():
        name = etree.QName(node)
        if name.localname in _HAZARD or name.namespace == M:
            return False
    return True


def _groups(blocks: tuple[BlockNode, ...]) -> list[tuple[str, list[BlockNode]]]:
    groups: list[tuple[str, list[BlockNode]]] = []
    for block in blocks:
        anchor = block.provenance.anchor if block.provenance is not None else ""
        if anchor and groups and groups[-1][0] == anchor:
            groups[-1][1].append(block)
            continue
        groups.append((anchor, [block]))
    return groups


def _same(left: list[BlockNode], right: list[BlockNode]) -> bool:
    return [block.project() for block in left] == [block.project() for block in right]


def _insertable(block: BlockNode) -> bool:
    return block.kind == "paragraph" and _inlines_supported(block.inlines) and not block.children


def _inlines_supported(nodes: tuple[InlineNode, ...]) -> bool:
    for node in nodes:
        if node.kind not in _SUPPORTED_INLINE:
            return False
        if node.children and not _inlines_supported(node.children):
            return False
    return True


def _marked_runs(nodes: tuple[InlineNode, ...], *, bold: bool = False, italic: bool = False) -> list[tuple[str, bool, bool]]:
    runs: list[tuple[str, bool, bool]] = []
    for node in nodes:
        if node.kind == "text":
            runs.append((node.text, bold, italic))
        elif node.kind == "strong":
            runs.extend(_marked_runs(node.children or (InlineNode("text", node.text),), bold=True, italic=italic))
        elif node.kind == "emph":
            runs.extend(_marked_runs(node.children or (InlineNode("text", node.text),), bold=bold, italic=True))
    return [(text, bold, italic) for text, bold, italic in runs if text]


def _set_mark(properties: etree._Element, tag: str, enabled: bool) -> None:
    for name in (tag, f"{tag}Cs"):
        existing = properties.find(f"{{{W}}}{name}")
        if existing is not None:
            properties.remove(existing)
    if enabled:
        properties.append(etree.Element(f"{{{W}}}{tag}"))


def _resource_key(document: SemanticDocument) -> tuple:
    return tuple((item.id, item.sha256, item.path) for item in document.resources)


def _refuse(reason: str) -> None:
    raise ConversionError(f"preserve fidelity refuses this DOCX edit: {reason}", code="UNSUPPORTED_PATCH")


__all__ = ["patch_docx"]
