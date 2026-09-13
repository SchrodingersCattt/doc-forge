"""Word tracked-changes diffing with comment preservation.

``create_tracked_docx`` aligns the paragraphs/tables of a reviewed baseline
with a freshly generated DOCX, then rebuilds the current package so that:

- unchanged paragraphs are copied intact (preserving page breaks, tabs,
  fields, drawings, and run-level formatting);
- text edits become native Word revisions (``w:ins`` / ``w:del`` /
  ``w:moveFrom`` / ``w:moveTo`` / ``w:rPrChange`` / ``w:pPrChange``);
- review comments and their anchors are carried forward from the baseline;
- the resulting package is validated (relationship ids, tracking ordering,
  and acceptance-identity against the fresh document).
"""

from __future__ import annotations

import copy
import difflib
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from .package import Package
from ..output import validate_output_path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W}
TOKEN_RE = re.compile(r"\s+|[^\W_]+(?:[’'][^\W_]+)*|_|[^\w\s]", re.UNICODE)
CLOSING_PUNCTUATION = frozenset("，。；：！？、）】》〉」』〕〗〙〛”’)]}")
REVISION_NAMES = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange")
COMMENT_NAMES = (
    "comments.xml",
    "commentsExtended.xml",
    "commentsExtensible.xml",
    "commentsIds.xml",
    "people.xml",
)
PASSTHROUGH_LOCAL_NAMES = {
    "drawing",
    "pict",
    "object",
    "fldChar",
    "instrText",
    "hyperlink",
    "footnoteReference",
    "endnoteReference",
    "sym",
    "noBreakHyphen",
    "softHyphen",
    "ptab",
    "sdt",
    "oMath",
    "sectPr",
}


@dataclass(frozen=True)
class Style:
    rpr: etree._Element | None


@dataclass(frozen=True)
class Token:
    text: str
    start: int
    end: int
    style: Style


@dataclass(frozen=True)
class Event:
    kind: str
    offset: int
    order: int
    element: etree._Element


@dataclass(frozen=True)
class Block:
    element: etree._Element
    kind: str
    text: str
    style: str
    drawing: bool


class Context:
    def __init__(self, author: str, next_id: int):
        self.author = author
        self.next_id = next_id
        self.date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def attrs(self) -> dict[str, str]:
        result = {
            f"{{{W}}}id": str(self.next_id),
            f"{{{W}}}author": self.author,
            f"{{{W}}}date": self.date,
        }
        self.next_id += 1
        return result


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", text)).strip()


def _inside(node: etree._Element, local_name: str) -> bool:
    target = f"{{{W}}}{local_name}"
    parent = node.getparent()
    while parent is not None:
        if parent.tag == target:
            return True
        parent = parent.getparent()
    return False


def visible_text(element: etree._Element, view: str = "final") -> str:
    out: list[str] = []

    def walk(node: etree._Element, hidden: bool = False) -> None:
        if node.tag in (f"{{{W}}}del", f"{{{W}}}moveFrom"):
            hidden = view == "final"
        elif node.tag in (f"{{{W}}}ins", f"{{{W}}}moveTo"):
            hidden = view == "original"
        if hidden:
            return
        if node.tag == f"{{{W}}}t":
            out.append(node.text or "")
            return
        if node.tag == f"{{{W}}}delText":
            if view != "final":
                out.append(node.text or "")
            return
        if node.tag == f"{{{W}}}tab":
            out.append("\t")
            return
        if node.tag in (f"{{{W}}}br", f"{{{W}}}cr"):
            out.append("\n")
            return
        for child in node:
            walk(child, hidden)

    walk(element)
    return "".join(out)


def _paragraph_style(paragraph: etree._Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.get(f"{{{W}}}val", "") if style is not None else ""


def _blocks(root: etree._Element) -> tuple[etree._Element, list[Block]]:
    body = root.find(".//w:body", NS)
    if body is None:
        raise ValueError("DOCX has no document body")
    blocks: list[Block] = []
    for element in body:
        if element.tag == f"{{{W}}}sectPr":
            continue
        kind = "p" if element.tag == f"{{{W}}}p" else "tbl" if element.tag == f"{{{W}}}tbl" else "other"
        blocks.append(
            Block(
                element,
                kind,
                visible_text(element),
                _paragraph_style(element) if kind == "p" else "",
                element.find(".//w:drawing", NS) is not None,
            )
        )
    return body, blocks


def _score(left: Block, right: Block) -> float:
    if left.kind != right.kind:
        return -10.0
    a, b = _normalize(left.text), _normalize(right.text)
    if left.kind != "p":
        if a == b:
            return 5.0
        return 2.0 if left.kind == "tbl" else -2.0
    if not a and not b:
        return 4.0 if left.drawing == right.drawing and left.drawing else 0.5
    if not a or not b:
        return -2.0
    if a == b:
        return 6.0
    ratio = difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()
    return -2.5 if ratio < 0.22 else 5.0 * ratio - 1.5 + (0.5 if left.style == right.style else 0.0)


def _align(base: list[Block], current: list[Block]) -> list[tuple[str, int | None, int | None]]:
    n, m, gap = len(base), len(current), -1.35
    scores = [[0.0] * (m + 1) for _ in range(n + 1)]
    steps: list[list[str | None]] = [[None] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        scores[i][0], steps[i][0] = scores[i - 1][0] + gap, "delete"
    for j in range(1, m + 1):
        scores[0][j], steps[0][j] = scores[0][j - 1] + gap, "insert"
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            scores[i][j], steps[i][j] = max(
                (scores[i - 1][j - 1] + _score(base[i - 1], current[j - 1]), "match"),
                (scores[i - 1][j] + gap, "delete"),
                (scores[i][j - 1] + gap, "insert"),
                key=lambda item: item[0],
            )
    edits: list[tuple[str, int | None, int | None]] = []
    i, j = n, m
    while i or j:
        action = steps[i][j]
        if action == "match":
            edits.append((action, i - 1, j - 1))
            i -= 1
            j -= 1
        elif action == "delete":
            edits.append((action, i - 1, None))
            i -= 1
        else:
            edits.append(("insert", None, j - 1))
            j -= 1
    return list(reversed(edits))


def _segments(paragraph: etree._Element, view: str) -> tuple[str, list[tuple[int, int, Style]]]:
    pieces: list[str] = []
    spans: list[tuple[int, int, Style]] = []
    offset = 0
    for node in paragraph.iter():
        if node.tag in (f"{{{W}}}tab", f"{{{W}}}br", f"{{{W}}}cr"):
            value = "\t" if node.tag == f"{{{W}}}tab" else "\n"
        elif node.tag in (f"{{{W}}}t", f"{{{W}}}delText"):
            value = node.text or ""
        else:
            continue
        if view == "final" and (_inside(node, "del") or _inside(node, "moveFrom")):
            continue
        if not value:
            continue
        run = node
        while run is not None and run.tag != f"{{{W}}}r":
            run = run.getparent()
        rpr = run.find("./w:rPr", NS) if run is not None else None
        style = Style(copy.deepcopy(rpr) if rpr is not None else None)
        pieces.append(value)
        spans.append((offset, offset + len(value), style))
        offset += len(value)
    return "".join(pieces), spans


def _style_at(spans: list[tuple[int, int, Style]], offset: int) -> Style:
    for start, end, style in spans:
        if start <= offset < end:
            return style
    return spans[-1][2] if spans else Style(None)


def _tokenize(
    text: str, spans: list[tuple[int, int, Style]], split_offsets: set[int] | None = None
) -> list[Token]:
    split_offsets = set(split_offsets or ())
    # Preserve character-level formatting boundaries (e.g. true Word
    # subscript/superscript inside chemical formulae) even when the tokenizer
    # would otherwise treat the whole alphanumeric formula as one token.
    split_offsets.update(start for start, _, _ in spans)
    split_offsets.update(end for _, end, _ in spans)
    result: list[Token] = []
    for match in TOKEN_RE.finditer(text):
        points = [
            match.start(),
            *sorted(x for x in split_offsets if match.start() < x < match.end()),
            match.end(),
        ]
        for start, end in zip(points, points[1:]):
            token = Token(text[start:end], start, end, _style_at(spans, start))
            if token.text in CLOSING_PUNCTUATION and result and result[-1].end == start:
                previous = result[-1]
                # Chinese closing punctuation must stay with the preceding
                # token. Word frequently stores it in a separate run (or at a
                # review-format boundary); preserving that artificial boundary
                # creates a standalone revision run that can wrap to the next
                # line as an orphan punctuation mark.
                result[-1] = Token(previous.text + token.text, previous.start, token.end, previous.style)
            else:
                result.append(token)
    return result


def _events(paragraph: etree._Element) -> list[Event]:
    tags = {
        f"{{{W}}}commentRangeStart": "start",
        f"{{{W}}}commentRangeEnd": "end",
        f"{{{W}}}commentReference": "reference",
    }
    events: list[Event] = []
    offset = 0

    def walk(node: etree._Element) -> None:
        nonlocal offset
        if node.tag in tags:
            owner = node.getparent() if tags[node.tag] == "reference" else node
            events.append(Event(tags[node.tag], offset, len(events), copy.deepcopy(owner)))
            return
        if node.tag in (f"{{{W}}}t", f"{{{W}}}delText"):
            offset += len(node.text or "")
            return
        if node.tag in (f"{{{W}}}tab", f"{{{W}}}br", f"{{{W}}}cr"):
            offset += 1
            return
        for child in node:
            walk(child)

    walk(paragraph)
    return events


def _needs_passthrough(paragraph: etree._Element) -> bool:
    """Identify paragraphs whose layout-bearing XML must remain intact."""
    return any(etree.QName(node).localname in PASSTHROUGH_LOCAL_NAMES for node in paragraph.iter())


def _run(token: Token, deleted: bool = False) -> etree._Element:
    run = etree.Element(f"{{{W}}}r")
    if token.style.rpr is not None:
        run.append(copy.deepcopy(token.style.rpr))
    if token.text == "\t":
        etree.SubElement(run, f"{{{W}}}tab")
        return run
    if token.text == "\n":
        etree.SubElement(run, f"{{{W}}}br")
        return run
    text = etree.SubElement(run, f"{{{W}}}{'delText' if deleted else 't'}")
    if token.text[:1].isspace() or token.text[-1:].isspace():
        text.set(f"{{{XML}}}space", "preserve")
    text.text = token.text
    return run


def _revision(kind: str, tokens: list[Token], context: Context) -> etree._Element:
    wrapper = etree.Element(f"{{{W}}}{kind}", attrib=context.attrs())
    for token in tokens:
        wrapper.append(_run(token, deleted=kind == "del"))
    return wrapper


def _merge_paragraph(
    base: etree._Element, current: etree._Element, context: Context
) -> etree._Element:
    base_text, base_spans = _segments(base, "final")
    current_text, current_spans = _segments(current, "final")
    events = _events(base)
    # Preserve structurally complex paragraphs as complete current-package XML.
    # Character-level redlining would otherwise discard layout controls that
    # have no w:t representation. Review comments remain available by carrying
    # their anchors onto the preserved paragraph.
    if _needs_passthrough(base) or _needs_passthrough(current):
        return _carry_comment_markers(base, current) if events else copy.deepcopy(current)
    # Identical, uncommented paragraphs already have the desired final XML in
    # the freshly generated document. Copying them intact preserves page
    # breaks, tabs, fields and other run-level controls that carry no text and
    # therefore do not participate in the token diff below.
    if base_text == current_text and not events:
        return copy.deepcopy(current)
    base_tokens = _tokenize(base_text, base_spans, {event.offset for event in events})
    current_tokens = _tokenize(current_text, current_spans)
    matcher = difflib.SequenceMatcher(
        None, [t.text for t in base_tokens], [t.text for t in current_tokens], autojunk=False
    )
    nodes: list[etree._Element] = []
    boundary: dict[int, int] = {0: 0}
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        boundary.setdefault(i1, len(nodes))
        if tag == "equal":
            for offset, token in enumerate(current_tokens[j1:j2], i1):
                boundary.setdefault(offset, len(nodes))
                nodes.append(_run(token))
                boundary[offset + 1] = len(nodes)
        elif tag == "delete":
            nodes.append(_revision("del", base_tokens[i1:i2], context))
        elif tag == "insert":
            nodes.append(_revision("ins", current_tokens[j1:j2], context))
        else:
            nodes.append(_revision("del", base_tokens[i1:i2], context))
            nodes.append(_revision("ins", current_tokens[j1:j2], context))
        boundary[i2] = len(nodes)
    offset_map = {0: 0}
    for index, token in enumerate(base_tokens, 1):
        offset_map[token.end] = index
    placed: dict[int, list[Event]] = {}
    for event in events:
        token_boundary = offset_map.get(event.offset, 0)
        placed.setdefault(boundary.get(token_boundary, len(nodes)), []).append(event)
    result = etree.Element(f"{{{W}}}p", nsmap=current.nsmap)
    ppr = current.find("./w:pPr", NS)
    ppr_from_current = ppr is not None
    if ppr is None:
        ppr = base.find("./w:pPr", NS)
    if ppr is not None:
        copied_ppr = copy.deepcopy(ppr)
        # Relationship ids are package-local. A section definition copied from
        # the reviewed package can therefore point to an unrelated part in the
        # newly generated package and make Word reject the document. Section
        # definitions originating in the current package are safe and must
        # remain, because they carry the intermediate section break.
        if not ppr_from_current:
            for sectpr in copied_ppr.findall("./w:sectPr", NS):
                copied_ppr.remove(sectpr)
        result.append(copied_ppr)
    for position in range(len(nodes) + 1):
        for event in sorted(placed.get(position, []), key=lambda item: item.order):
            result.append(copy.deepcopy(event.element))
        if position < len(nodes):
            result.append(nodes[position])
    return result


def _mark_paragraph(
    paragraph: etree._Element, kind: str, context: Context
) -> etree._Element:
    if kind == "del":
        # A paragraph-level deletion marker hides the paragraph in the final
        # view.  Keep the original runs in place so Word's original view can
        # still show the deleted paragraph; diffing against an empty paragraph
        # would otherwise discard those runs entirely.
        result = copy.deepcopy(paragraph)
        result_ppr = result.find("./w:pPr", NS)
        if result_ppr is None:
            result_ppr = etree.Element(f"{{{W}}}pPr")
            result.insert(0, result_ppr)
        rpr = result_ppr.find("./w:rPr", NS)
        if rpr is None:
            rpr = etree.SubElement(result_ppr, f"{{{W}}}rPr")
        rpr.insert(0, etree.Element(f"{{{W}}}del", attrib=context.attrs()))
        return result
    result = copy.deepcopy(paragraph)
    ppr = result.find("./w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(f"{{{W}}}pPr")
        result.insert(0, ppr)
    rpr = ppr.find("./w:rPr", NS)
    if rpr is None:
        rpr = etree.SubElement(ppr, f"{{{W}}}rPr")
    rpr.insert(0, etree.Element(f"{{{W}}}ins", attrib=context.attrs()))
    children = [child for child in result if child.tag != f"{{{W}}}pPr"]
    for child in children:
        result.remove(child)
    wrapper = etree.Element(f"{{{W}}}ins", attrib=context.attrs())
    for child in children:
        wrapper.append(child)
    result.append(wrapper)
    return result


def _carry_comment_markers(
    base: etree._Element, current: etree._Element
) -> etree._Element:
    """Attach comment anchors from a non-text paragraph to its current drawing."""
    result = copy.deepcopy(current)
    events = _events(base)
    if not events:
        return result
    insert_at = 1 if len(result) and result[0].tag == f"{{{W}}}pPr" else 0
    starts = [copy.deepcopy(event.element) for event in events if event.kind == "start"]
    endings = [copy.deepcopy(event.element) for event in events if event.kind != "start"]
    for element in reversed(starts):
        result.insert(insert_at, element)
    for element in endings:
        result.append(element)
    return result


def _merge_table(
    base: etree._Element, current: etree._Element, context: Context
) -> etree._Element:
    """Keep the current table layout and track cell-text edits where possible."""
    result = copy.deepcopy(current)
    base_paragraphs = base.findall(".//w:p", NS)
    current_paragraphs = result.findall(".//w:p", NS)

    # The official form normally keeps the same table geometry between drafts.
    # In that common case, diff every corresponding cell paragraph so changes
    # to metadata and abstracts remain visible as native Word revisions.
    if len(base_paragraphs) == len(current_paragraphs):
        for base_paragraph, current_paragraph in zip(base_paragraphs, list(current_paragraphs)):
            merged = _merge_paragraph(base_paragraph, current_paragraph, context)
            current_paragraph.getparent().replace(current_paragraph, merged)
        return result

    # If the table geometry changed, retain the current table and salvage any
    # review comments that can be matched to unchanged cell text.
    used: set[int] = set()
    for base_paragraph in base_paragraphs:
        if not _events(base_paragraph):
            continue
        target_text = _normalize(visible_text(base_paragraph))
        match = next(
            (
                (index, paragraph)
                for index, paragraph in enumerate(current_paragraphs)
                if index not in used and _normalize(visible_text(paragraph)) == target_text
            ),
            None,
        )
        if match is None:
            continue
        index, paragraph = match
        used.add(index)
        merged = _merge_paragraph(base_paragraph, paragraph, context)
        paragraph.getparent().replace(paragraph, merged)
        current_paragraphs[index] = merged
    return result


def _next_id(root: etree._Element) -> int:
    values: list[int] = []
    for name in REVISION_NAMES:
        for node in root.findall(f".//w:{name}", NS):
            try:
                values.append(int(node.get(f"{{{W}}}id", "-1")))
            except ValueError:
                pass
    return max(values, default=-1) + 1


def _comment_counts(root: etree._Element) -> dict[str, tuple[int, int, int]]:
    result: dict[str, list[int]] = {}
    for index, name in enumerate(("commentRangeStart", "commentRangeEnd", "commentReference")):
        for node in root.findall(f".//w:{name}", NS):
            result.setdefault(node.get(f"{{{W}}}id", ""), [0, 0, 0])[index] += 1
    return {key: tuple(value) for key, value in result.items()}


def _copy_comment_parts(base: Package, current: Package) -> None:
    def is_review_relationship(rel_type: str, target: str = "") -> bool:
        rel_type = rel_type.lower()
        target = target.lower()
        return "comment" in rel_type or rel_type.endswith("/people") or "people.xml" in target

    for short_name in COMMENT_NAMES:
        name = f"word/{short_name}"
        if name in base.parts:
            current.parts[name] = base.parts[name]
    rel_name = "word/_rels/document.xml.rels"
    if rel_name in base.parts and rel_name in current.parts:
        base_rel = etree.fromstring(base.parts[rel_name])
        current_rel = etree.fromstring(current.parts[rel_name])
        existing_ids = {node.get("Id") for node in current_rel}
        existing_types = {node.get("Type") for node in current_rel}
        for relationship in base_rel:
            rel_type = relationship.get("Type", "")
            if not is_review_relationship(rel_type, relationship.get("Target", "")) or rel_type in existing_types:
                continue
            copied = copy.deepcopy(relationship)
            if copied.get("Id") in existing_ids:
                number = 1
                while f"rId{number}" in existing_ids:
                    number += 1
                copied.set("Id", f"rId{number}")
            existing_ids.add(copied.get("Id"))
            existing_types.add(rel_type)
            current_rel.append(copied)
        current.parts[rel_name] = etree.tostring(
            current_rel, xml_declaration=True, encoding="UTF-8", standalone=True
        )
    content_name = "[Content_Types].xml"
    if content_name in base.parts and content_name in current.parts:
        base_types = etree.fromstring(base.parts[content_name])
        current_types = etree.fromstring(current.parts[content_name])
        existing = {node.get("PartName") for node in current_types}
        for node in base_types:
            part_name = node.get("PartName", "")
            lowered = part_name.lower()
            if ("comment" in lowered or lowered.endswith("/people.xml")) and part_name not in existing:
                current_types.append(copy.deepcopy(node))
        current.parts[content_name] = etree.tostring(
            current_types, xml_declaration=True, encoding="UTF-8", standalone=True
        )


def _enable_tracking(package: Package) -> None:
    name = "word/settings.xml"
    root = package.xml(name)
    if root.find("./w:trackRevisions", NS) is None:
        marker = etree.Element(f"{{{W}}}trackRevisions")
        # CT_Settings is an ordered sequence. trackRevisions belongs after
        # proofState/revisionView and before defaultTabStop and later settings.
        default_tab = root.find("./w:defaultTabStop", NS)
        if default_tab is not None:
            root.insert(root.index(default_tab), marker)
        else:
            root.append(marker)
    package.set_xml(name, root)


def _validate_package_relationships(package: Package) -> None:
    """Reject package-local relationship ids copied from another DOCX."""
    document = package.xml("word/document.xml")
    relationships = package.xml("word/_rels/document.xml.rels")
    by_id = {node.get("Id", ""): node.get("Type", "") for node in relationships}
    checks = (
        (f".//{{{W}}}headerReference", "header"),
        (f".//{{{W}}}footerReference", "footer"),
        (f".//{{{W}}}hyperlink", "hyperlink"),
        (f".//{{{A}}}blip", "image"),
    )
    for path, expected in checks:
        for element in document.findall(path):
            rid = element.get(f"{{{R}}}id") or element.get(f"{{{R}}}embed")
            if not rid:
                continue
            rel_type = by_id.get(rid, "")
            if not rel_type.endswith(f"/{expected}"):
                raise AssertionError(
                    f"Relationship {rid!r} used by {etree.QName(element).localname} "
                    f"has type {rel_type!r}, expected {expected!r}"
                )


def _validate_tracking_position(package: Package) -> None:
    settings = package.xml("word/settings.xml")
    names = [etree.QName(child).localname for child in settings]
    index = names.index("trackRevisions")
    if "defaultTabStop" in names and index > names.index("defaultTabStop"):
        raise AssertionError("trackRevisions is out of order in word/settings.xml")
    for earlier in ("zoom", "proofState", "revisionView"):
        if earlier in names and index < names.index(earlier):
            raise AssertionError("trackRevisions is out of order in word/settings.xml")


def _final_blocks(root: etree._Element) -> list[str]:
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    result: list[str] = []
    for element in body:
        if element.tag not in (f"{{{W}}}p", f"{{{W}}}tbl"):
            continue
        if element.tag == f"{{{W}}}p" and element.find("./w:pPr/w:rPr/w:del", NS) is not None:
            continue
        result.append(visible_text(element, "final"))
    return result


def _xml_signature(element: etree._Element) -> tuple:
    """Return a prefix-independent structural signature for an OOXML node."""
    name = etree.QName(element).localname
    attrs = tuple(sorted((etree.QName(key).localname, value) for key, value in element.attrib.items()))
    return name, attrs, tuple(_xml_signature(child) for child in element)


def _final_special_structure(root: etree._Element) -> list[tuple]:
    """Capture final-view controls whose layout effect is absent from text."""
    body = root.find(".//w:body", NS)
    if body is None:
        return []
    result: list[tuple] = []

    def walk(node: etree._Element, hidden: bool, items: list[tuple]) -> None:
        if node.tag in (f"{{{W}}}del", f"{{{W}}}moveFrom"):
            hidden = True
        elif node.tag in (f"{{{W}}}ins", f"{{{W}}}moveTo"):
            hidden = False
        if hidden:
            return
        local = etree.QName(node).localname
        if node.tag == f"{{{W}}}sectPr":
            items.append(("sectPr", _xml_signature(node)))
            return
        if node.tag in (f"{{{W}}}br", f"{{{W}}}cr", f"{{{W}}}tab"):
            attrs = tuple(
                sorted((etree.QName(key).localname, value) for key, value in node.attrib.items())
            )
            items.append((local, attrs))
            return
        if local == "drawing":
            items.append(("drawing",))
            return
        for child in node:
            walk(child, hidden, items)

    for block in body:
        if block.tag == f"{{{W}}}sectPr":
            continue
        if block.tag == f"{{{W}}}p" and block.find("./w:pPr/w:rPr/w:del", NS) is not None:
            continue
        items: list[tuple] = []
        walk(block, False, items)
        result.append(tuple(items))
    body_section = body.find("./w:sectPr", NS)
    result.append(("bodySectPr", _xml_signature(body_section) if body_section is not None else None))
    return result


def _accepted_revision_view(root: etree._Element) -> etree._Element:
    """Return a comment-preserving copy with all earlier revisions accepted.

    A reviewed baseline may itself contain tracked changes. Diffing its raw
    revision tree can align deleted cover metadata with similarly worded body
    paragraphs and leak stale deletion text into the next redline. The next
    comparison must therefore start from the baseline's accepted view while
    retaining comment anchors and the separate comment parts.
    """
    result = copy.deepcopy(root)

    # Paragraph deletions use a revision marker in pPr rather than a wrapper
    # around the paragraph text. Remove those paragraphs before stripping the
    # remaining revision metadata, otherwise they would survive as empty blocks.
    for paragraph in list(result.xpath(".//w:p[w:pPr/w:rPr/w:del]", namespaces=NS)):
        parent = paragraph.getparent()
        if parent is not None:
            parent.remove(paragraph)

    # Reject deleted/moved-from content.
    for name in ("del", "moveFrom"):
        for node in list(result.findall(f".//w:{name}", NS)):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)

    # Accept inserted/moved-to content by unwrapping it in place. Iterate until
    # no wrappers remain so nested revisions are handled deterministically.
    for name in ("ins", "moveTo"):
        while True:
            wrappers = list(result.findall(f".//w:{name}", NS))
            if not wrappers:
                break
            for node in wrappers:
                parent = node.getparent()
                if parent is None:
                    continue
                position = parent.index(node)
                for child in list(node):
                    node.remove(child)
                    parent.insert(position, child)
                    position += 1
                parent.remove(node)

    # The current formatting already represents the accepted formatting state.
    for name in ("rPrChange", "pPrChange", "sectPrChange"):
        for node in list(result.findall(f".//w:{name}", NS)):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
    return result


def create_tracked_docx(
    base_path: Path,
    current_path: Path,
    output_path: Path,
    *,
    author: str = "M.Y.G.",
    overwrite: bool = False,
) -> dict[str, int]:
    validate_output_path(output_path)
    base_package = Package.load(base_path)
    current_package = Package.load(current_path)
    raw_base_root = base_package.xml("word/document.xml")
    base_root = _accepted_revision_view(raw_base_root)
    # A freshly supplied current document may itself contain an earlier
    # review pass.  Diff only its accepted view so stale w:ins/w:del and
    # formatting-change markers cannot leak into the new redline.
    raw_current_root = current_package.xml("word/document.xml")
    current_root = _accepted_revision_view(raw_current_root)
    base_body, base_blocks = _blocks(base_root)
    current_body, current_blocks = _blocks(current_root)
    context = Context(author, _next_id(raw_base_root))
    summary = {"matched": 0, "changed": 0, "inserted": 0, "deleted": 0, "tables_replaced": 0}
    children: list[etree._Element] = []
    for action, base_index, current_index in _align(base_blocks, current_blocks):
        if action == "match":
            old = base_blocks[base_index]  # type: ignore[index]
            new = current_blocks[current_index]  # type: ignore[index]
            if old.kind == "p":
                children.append(
                    _carry_comment_markers(old.element, new.element)
                    if old.drawing or new.drawing
                    else _merge_paragraph(old.element, new.element, context)
                )
                summary["matched" if old.text == new.text else "changed"] += 1
            elif old.kind == "tbl":
                children.append(_merge_table(old.element, new.element, context))
                summary["matched" if _normalize(old.text) == _normalize(new.text) else "tables_replaced"] += 1
            else:
                children.append(
                    copy.deepcopy(old.element if _normalize(old.text) == _normalize(new.text) else new.element)
                )
                summary["matched" if _normalize(old.text) == _normalize(new.text) else "tables_replaced"] += 1
        elif action == "delete":
            old = base_blocks[base_index]  # type: ignore[index]
            children.append(
                _mark_paragraph(old.element, "del", context) if old.kind == "p" else copy.deepcopy(old.element)
            )
            summary["deleted"] += 1
        else:
            new = current_blocks[current_index]  # type: ignore[index]
            children.append(
                _mark_paragraph(new.element, "ins", context) if new.kind == "p" else copy.deepcopy(new.element)
            )
            summary["inserted"] += 1
    current_sectpr = current_body.find("./w:sectPr", NS)
    for child in list(current_body):
        current_body.remove(child)
    for child in children:
        current_body.append(child)
    if current_sectpr is not None:
        current_body.append(copy.deepcopy(current_sectpr))
    original_markers = _comment_counts(base_root)
    final_markers = _comment_counts(current_root)
    if final_markers != original_markers:
        missing = {key: value for key, value in original_markers.items() if final_markers.get(key) != value}
        if missing:
            raise AssertionError(f"Comment anchors were not preserved: {missing}")
    current_package.set_xml("word/document.xml", current_root)
    _copy_comment_parts(base_package, current_package)
    _enable_tracking(current_package)
    _validate_package_relationships(current_package)
    _validate_tracking_position(current_package)
    if _final_blocks(current_root) != _final_blocks(current_package.xml("word/document.xml")):
        raise AssertionError("Tracked final view changed while packaging the document")
    original_current_root = Package.load(current_path).xml("word/document.xml")
    if _final_blocks(current_root) != _final_blocks(original_current_root):
        raise AssertionError("Accepting all revisions would not reproduce the generated DOCX")
    if _final_special_structure(current_root) != _final_special_structure(original_current_root):
        raise AssertionError("Tracked final view changed non-text layout controls from the generated DOCX")
    current_package.write(output_path, overwrite=overwrite)
    return summary


__all__ = ["create_tracked_docx", "Package"]
