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
import hashlib
import json
import os
import posixpath
import re
import unicodedata
import zipfile
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from .package import Package
from ..output import validate_output_path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W}
TOKEN_RE = re.compile(r"\s+|[^\W_]+(?:[’'][^\W_]+)*|_|[^\w\s]", re.UNICODE)
CLOSING_PUNCTUATION = frozenset("，。；：！？、）】》〉」』〕〗〙〛”’)]}")
# Markdown round-trips and Word's Unicode-script normalizer can represent the
# same visible character in several ways.  Redline alignment must compare the
# semantic text, not those transport details.  In particular, Pandoc emits
# escaped punctuation (``\<``/``\[``) and the template converter turns
# ``N₅⁻`` into an ASCII ``N5–`` run with true script formatting.
_ESCAPED_PUNCTUATION = re.compile(r"\\([\\`*{}\[\]()#+\-.!_>~|<%$&=:?@])")
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
# Internal citation links are ordinary runs plus a w:hyperlink wrapper.
# Treating every hyperlink as opaque layout XML forces a whole-paragraph
# replacement, so Word shows the old paragraph as live text and the new one
# as a separate insertion. Relationship-bearing hyperlinks are still copied
# only from the current package, whose relationship ids remain valid.


@dataclass(frozen=True)
class Style:
    rpr: etree._Element | None
    revision: etree._Element | None = None
    link: tuple[tuple[str, str], ...] | None = None


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
    companion: etree._Element | None = None


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


def _semantic_text(text: str) -> str:
    """Canonicalize text for alignment while preserving source XML on output.

    This is deliberately used only for comparison.  The emitted runs still
    come from the current DOCX, so punctuation, Unicode scripts, and
    formatting are not silently rewritten in the deliverable.
    """
    value = unicodedata.normalize("NFKC", text)
    value = _ESCAPED_PUNCTUATION.sub(r"\1", value)
    # NFKC produces U+2212 for superscript minus, while the DOCX converter
    # uses an en dash in superscript chemical charges.  Dash typography is not
    # a semantic manuscript edit for redline purposes.
    value = value.replace("−", "-").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", value).strip()


_NOISE_PARAGRAPH = re.compile(
    r"^(?:suppnote|supprecord|landscape|small(?:\s+\d+pt)?)$",
    re.IGNORECASE,
)


def _hyperlink_attrs(node: etree._Element) -> tuple[tuple[str, str], ...] | None:
    parent = node.getparent()
    while parent is not None:
        if parent.tag == f"{{{W}}}hyperlink":
            return tuple(sorted(parent.attrib.items()))
        if parent.tag == f"{{{W}}}p":
            return None
        parent = parent.getparent()
    return None


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

    if (
        view == "final"
        and element.tag == f"{{{W}}}p"
        and element.find("./w:pPr/w:rPr/w:del", NS) is not None
    ):
        return ""

    def walk(node: etree._Element, hidden: bool = False) -> None:
        if (
            view == "final"
            and node.tag == f"{{{W}}}p"
            and node.find("./w:pPr/w:rPr/w:del", NS) is not None
        ):
            return
        if node.tag in (f"{{{W}}}del", f"{{{W}}}moveFrom"):
            hidden = view == "final"
        elif node.tag in (f"{{{W}}}ins", f"{{{W}}}moveTo"):
            hidden = view == "original"
        elif node.tag == f"{{{W}}}tr":
            # Word stores row-level insertions/deletions in ``w:trPr``.
            # Treat those markers like run-level revisions for both text
            # alignment and acceptance validation.
            tr_pr = node.find("./w:trPr", NS)
            if tr_pr is not None:
                if tr_pr.find("./w:del", NS) is not None:
                    hidden = view == "final"
                elif tr_pr.find("./w:ins", NS) is not None:
                    hidden = view == "original"
        if hidden:
            return
        if node.tag in (f"{{{W}}}t", f"{{{M}}}t"):
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
    raw = [element for element in body if element.tag != f"{{{W}}}sectPr"]
    blocks: list[Block] = []
    index = 0
    while index < len(raw):
        element = raw[index]
        kind = "p" if element.tag == f"{{{W}}}p" else "tbl" if element.tag == f"{{{W}}}tbl" else "other"
        drawing = element.find(".//w:drawing", NS) is not None
        companion = None
        if drawing and index + 1 < len(raw):
            following = visible_text(raw[index + 1]).strip()
            if re.match(r"Figure S\d+\.", following):
                companion = raw[index + 1]
        text_element = companion if companion is not None else element
        blocks.append(
            Block(
                element,
                kind,
                visible_text(text_element),
                _paragraph_style(element) if kind == "p" else "",
                drawing,
                companion,
            )
        )
        index += 2 if companion is not None else 1
    return body, blocks


def _pair_ratio(pair: tuple[str, str]) -> float:
    return difflib.SequenceMatcher(None, pair[0], pair[1], autojunk=False).ratio()


def _ratio_key(a: str, b: str) -> str:
    return hashlib.sha1(a.encode("utf-8")).hexdigest() + hashlib.sha1(b.encode("utf-8")).hexdigest()


class _RatioTable:
    """Paragraph similarity ratios, optionally persisted across runs.

    The ratio of two texts never changes, so a cache keyed by both text hashes
    returns exactly what ``SequenceMatcher`` would compute.
    """

    def __init__(self, cache_path: Path | None) -> None:
        self.cache_path = cache_path
        self.values: dict[str, float] = {}
        self.used: set[str] = set()
        self.dirty = False
        if cache_path is not None and cache_path.is_file():
            try:
                self.values = json.loads(cache_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                self.values = {}

    def fill(self, pairs: set[tuple[str, str]], workers: int | None) -> None:
        keys = {pair: _ratio_key(*pair) for pair in pairs}
        self.used.update(keys.values())
        if len(self.used) != len(self.values):
            self.dirty = True
        missing = [pair for pair, key in keys.items() if key not in self.values]
        if not missing:
            return
        if workers != 1 and len(missing) >= 200:
            with ProcessPoolExecutor(max_workers=workers) as pool:
                ratios = list(pool.map(_pair_ratio, missing, chunksize=32))
        else:
            ratios = [_pair_ratio(pair) for pair in missing]
        for pair, ratio in zip(missing, ratios):
            self.values[_ratio_key(*pair)] = ratio
        self.dirty = True

    def get(self, a: str, b: str) -> float:
        key = _ratio_key(a, b)
        if key not in self.used:
            self.used.add(key)
            self.dirty = True
        if key not in self.values:
            self.values[key] = _pair_ratio((a, b))
            self.dirty = True
        return self.values[key]

    def save(self) -> None:
        if self.cache_path is None or not self.dirty:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.cache_path.with_suffix(".tmp")
        kept = {key: self.values[key] for key in self.used if key in self.values}
        temporary.write_text(json.dumps(kept), encoding="utf-8")
        os.replace(temporary, self.cache_path)


def _float_label(text: str) -> str | None:
    match = re.match(r"((?:Figure|Table)\s+S?\d+)\b", text)
    return match.group(1) if match else None


def _score(left: Block, right: Block, ratios: _RatioTable | None = None) -> float:
    if left.kind != right.kind:
        return -10.0
    # An image paragraph must not be aligned with body text, and two different
    # images must not inherit each other's place. Otherwise a later caption is
    # word-merged onto the wrong figure.
    if left.drawing != right.drawing:
        return -8.0
    a, b = _semantic_text(left.text), _semantic_text(right.text)
    label_a, label_b = _float_label(a), _float_label(b)
    if label_a and label_b and label_a != label_b:
        return -8.0
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
    ratio = ratios.get(a, b) if ratios is not None else _pair_ratio((a, b))
    # Two gaps cost -2.70, so a weak match must score below that or the
    # aligner splices unrelated sentences into one paragraph. Long paragraphs
    # need a higher bar: a 0.3 overlap is still a different sentence, and
    # word-merging it is unreadable. Short strings stay on the 0.22 bar so
    # small edits such as "before"/"after" remain in-paragraph revisions.
    heading = "heading" in left.style.lower() or "heading" in right.style.lower()
    cutoff = 0.22 if heading or min(len(a), len(b)) < 80 else 0.55
    return -8.0 if ratio < cutoff else 5.0 * ratio - 1.5 + (0.5 if left.style == right.style else 0.0)


def _align(
    base: list[Block],
    current: list[Block],
    *,
    cache_path: Path | None = None,
    workers: int | None = None,
) -> list[tuple[str, int | None, int | None]]:
    ratios = _RatioTable(cache_path)
    base_text = [_semantic_text(block.text) if block.kind == "p" else "" for block in base]
    current_text = [_semantic_text(block.text) if block.kind == "p" else "" for block in current]
    ratios.fill(
        {
            (a, b)
            for a in set(base_text) if a
            for b in set(current_text) if b and b != a
        },
        workers,
    )
    try:
        return _align_scored(base, current, ratios)
    finally:
        ratios.save()


def _align_scored(
    base: list[Block],
    current: list[Block],
    ratios: _RatioTable,
) -> list[tuple[str, int | None, int | None]]:
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
                (scores[i - 1][j - 1] + _score(base[i - 1], current[j - 1], ratios), "match"),
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


def _revision_shell(node: etree._Element) -> etree._Element | None:
    """Copy an enclosing insertion marker without its children."""
    ancestor = node.getparent()
    while ancestor is not None and ancestor.tag != f"{{{W}}}p":
        if ancestor.tag in (f"{{{W}}}ins", f"{{{W}}}moveTo"):
            return etree.Element(ancestor.tag, attrib=dict(ancestor.attrib), nsmap=ancestor.nsmap)
        ancestor = ancestor.getparent()
    return None


def _has_revisions(element: etree._Element) -> bool:
    return any(
        next(element.iter(f"{{{W}}}{name}"), None) is not None
        for name in ("ins", "del", "moveFrom", "moveTo")
    )


def _hidden_revision_nodes(paragraph: etree._Element) -> list[tuple[int, etree._Element]]:
    """Return baseline deletions with their final-view character offsets."""
    result: list[tuple[int, etree._Element]] = []
    offset = 0

    def walk(node: etree._Element) -> None:
        nonlocal offset
        if node.tag in (f"{{{W}}}del", f"{{{W}}}moveFrom"):
            if node.find(f".//{{{W}}}delText") is not None or node.find(f".//{{{W}}}t") is not None:
                result.append((offset, copy.deepcopy(node)))
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
    return result


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
        style = Style(
            copy.deepcopy(rpr) if rpr is not None else None,
            revision=_revision_shell(node),
            link=_hyperlink_attrs(node),
        )
        pieces.append(value)
        spans.append((offset, offset + len(value), style))
        offset += len(value)
    return "".join(pieces), spans


def _vert_align(style: Style) -> str | None:
    if style.rpr is None:
        return None
    marker = style.rpr.find(f"{{{W}}}vertAlign")
    if marker is None:
        return None
    return marker.get(f"{{{W}}}val")


def _style_at(spans: list[tuple[int, int, Style]], offset: int) -> Style:
    for start, end, style in spans:
        if start <= offset < end:
            return style
    return spans[-1][2] if spans else Style(None)


def _style_signature(style: Style) -> tuple[bytes | None, tuple[tuple[str, str], ...] | None, bytes | None]:
    """Return the formatting identity that matters when tokenizing text.

    Word commonly splits one visible word across several runs while keeping
    the same formatting.  Run-level revision IDs and rsid attributes are
    transport metadata, so treating those boundaries as semantic formatting
    changes creates false delete/insert pairs.  Keep boundaries only when the
    effective run properties, hyperlink, or enclosing revision actually
    differs.
    """

    # These properties change the character's visible script or emphasis.  The
    # remaining run properties in this workflow (font hints, language, size,
    # theme color, and rsids) are routinely rewritten by Pandoc/Word and must
    # not turn an unchanged word into a textual replacement.
    semantic_rpr = {
        "b",
        "i",
        "u",
        "strike",
        "dstrike",
        "vertAlign",
        "outline",
        "shadow",
        "emboss",
        "imprint",
        "caps",
        "smallCaps",
        "vanish",
        "rtl",
        "oMath",
    }

    def canonical(node: etree._Element | None, *, strip_rsid: bool = False) -> bytes | None:
        if node is None:
            return None
        if node.tag == f"{{{W}}}rPr":
            clone = etree.Element(node.tag, nsmap=node.nsmap)
            for child in node:
                if etree.QName(child).localname in semantic_rpr:
                    clone.append(copy.deepcopy(child))
        else:
            clone = copy.deepcopy(node)
        if strip_rsid:
            for element in clone.iter():
                for key in list(element.attrib):
                    if etree.QName(key).localname.startswith("rsid"):
                        del element.attrib[key]
        return etree.tostring(clone, method="c14n")

    return (
        canonical(style.rpr, strip_rsid=True),
        style.link,
        canonical(style.revision),
    )


def _tokenize(
    text: str, spans: list[tuple[int, int, Style]], split_offsets: set[int] | None = None
) -> list[Token]:
    split_offsets = set(split_offsets or ())
    # Preserve character-level formatting boundaries (e.g. true Word
    # subscript/superscript inside chemical formulae) even when the tokenizer
    # would otherwise treat the whole alphanumeric formula as one token.  Do
    # not preserve arbitrary run boundaries when the effective formatting is
    # identical: Word and Markdown round-trips frequently split one word into
    # several equivalent runs, and diffing those fragments produces false
    # replacements such as ``sorptionsorption``.
    for previous, current in zip(spans, spans[1:]):
        if _style_signature(previous[2]) != _style_signature(current[2]):
            split_offsets.add(previous[1])
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
                # line as an orphan punctuation mark. Do not glue across a
                # script change: a baseline ")" or "]" after a subscript would
                # otherwise inherit that subscript.
                if _vert_align(previous.style) == _vert_align(token.style):
                    result[-1] = Token(previous.text + token.text, previous.start, token.end, previous.style)
                else:
                    result.append(token)
            else:
                result.append(token)
    return result


def _events(paragraph: etree._Element) -> list[Event]:
    tags = {
        f"{{{W}}}commentRangeStart": "start",
        f"{{{W}}}commentRangeEnd": "end",
        f"{{{W}}}commentReference": "reference",
        f"{{{W}}}bookmarkStart": "start",
        f"{{{W}}}bookmarkEnd": "end",
    }
    events: list[Event] = []
    offset = 0

    def walk(node: etree._Element) -> None:
        nonlocal offset
        if node.tag in (f"{{{W}}}del", f"{{{W}}}moveFrom"):
            return
        if node.tag in tags:
            owner = node.getparent() if tags[node.tag] == "reference" else node
            events.append(Event(tags[node.tag], offset, len(events), copy.deepcopy(owner)))
            return
        if node.tag in (f"{{{W}}}t", f"{{{W}}}delText", f"{{{M}}}t"):
            offset += len(node.text or "")
            return
        if node.tag in (f"{{{W}}}tab", f"{{{W}}}br", f"{{{W}}}cr"):
            offset += 1
            return
        for child in node:
            walk(child)

    walk(paragraph)
    return events


_FIELD_LOCAL_NAMES = frozenset({"fldChar", "instrText"})


def _only_internal_hyperlink_fields(paragraph: etree._Element) -> bool:
    """True when every field in the paragraph is an internal ``HYPERLINK \\l`` link.

    Word saves a reviewed document's citation links as field codes, while the
    generated document uses ``w:hyperlink`` wrappers. Such a field carries no
    layout of its own, so it must not force a whole-paragraph replacement.
    """
    instructions = [node.text or "" for node in paragraph.iter(f"{{{W}}}instrText")]
    if not instructions:
        return False
    return all(re.match(r"\s*HYPERLINK\s+\\l\s", text) for text in instructions)


def _needs_passthrough(paragraph: etree._Element) -> bool:
    """Identify paragraphs whose layout-bearing XML must remain intact."""
    fields_are_links = _only_internal_hyperlink_fields(paragraph)
    return any(
        (
            etree.QName(node).localname in PASSTHROUGH_LOCAL_NAMES
            and not (fields_are_links and etree.QName(node).localname in _FIELD_LOCAL_NAMES)
        )
        or (node.tag == f"{{{W}}}br" and node.get(f"{{{W}}}type") in ("page", "column"))
        for node in paragraph.iter()
    )


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


def _materialize(token: Token, *, deleted: bool = False) -> etree._Element:
    node = _run(token, deleted=deleted)
    if token.style.revision is not None and not deleted:
        wrapper = copy.deepcopy(token.style.revision)
        wrapper.append(node)
        return wrapper
    return node


def _revision(kind: str, tokens: list[Token], context: Context) -> etree._Element:
    wrapper = etree.Element(f"{{{W}}}{kind}", attrib=context.attrs())
    for token in tokens:
        wrapper.append(_materialize(token, deleted=kind == "del"))
    return wrapper


def _emit_link(token: Token, kind: str | None) -> tuple[tuple[str, str], ...] | None:
    link = token.style.link
    if link is None:
        return None
    # Deleted text is copied from the reviewed package. An r:id on that
    # hyperlink points at the baseline relationships, not the package being
    # written, so keep the run formatting and drop the relationship.
    if kind == "del" and any(key == f"{{{R}}}id" for key, _value in link):
        return None
    return link


def _append_tracked(
    nodes: list[etree._Element],
    tokens: list[Token],
    kind: str | None,
    context: Context,
    boundary: dict[int, int],
    index0: int | None,
) -> None:
    """Append runs, grouping consecutive tokens that share one hyperlink."""
    index = 0
    while index < len(tokens):
        link = _emit_link(tokens[index], kind)
        end = index + 1
        if link is not None:
            while end < len(tokens) and _emit_link(tokens[end], kind) == link:
                end += 1
        else:
            while end < len(tokens) and _emit_link(tokens[end], kind) is None:
                end += 1
        group = tokens[index:end]
        if link is None and kind is None:
            for offset, token in enumerate(group):
                if index0 is not None:
                    boundary.setdefault(index0 + index + offset, len(nodes))
                nodes.append(_run(token))
                if index0 is not None:
                    boundary[index0 + index + offset + 1] = len(nodes)
            index = end
            continue
        if index0 is not None:
            for offset in range(index, end):
                boundary.setdefault(index0 + offset, len(nodes))
        payload = (
            [_run(token) for token in group]
            if kind is None
            else [_revision(kind, group, context)]
        )
        if link is None:
            nodes.extend(payload)
        else:
            hyperlink = etree.Element(f"{{{W}}}hyperlink")
            for key, value in link:
                hyperlink.set(key, value)
            for node in payload:
                hyperlink.append(node)
            nodes.append(hyperlink)
        if index0 is not None:
            boundary[index0 + end] = len(nodes)
        index = end


def _merge_paragraph(
    base: etree._Element,
    current: etree._Element,
    context: Context,
    *,
    preserve_base_revisions: bool = False,
) -> etree._Element:
    base_text, base_spans = _segments(base, "final")
    current_text, current_spans = _segments(current, "final")
    events = _events(base)
    base_has_revisions = preserve_base_revisions and _has_revisions(base)
    # Unchanged final text keeps the reviewed w:ins/w:del author, date, and
    # boundaries. Rebuilding the paragraph would drop metadata deletions such
    # as a removed corresponding-author mark or email.
    if base_has_revisions and base_text == current_text and not (
        _needs_passthrough(base) or _needs_passthrough(current)
    ):
        result = copy.deepcopy(base)
        current_ppr = current.find("./w:pPr", NS)
        old_ppr = result.find("./w:pPr", NS)
        if current_ppr is not None:
            if old_ppr is not None:
                result.replace(old_ppr, copy.deepcopy(current_ppr))
            else:
                result.insert(0, copy.deepcopy(current_ppr))
        elif old_ppr is not None:
            result.remove(old_ppr)
        return result
    # Preserve structurally complex paragraphs as complete current-package XML.
    # Character-level redlining would otherwise discard layout controls that
    # have no w:t representation. Review comments remain available by carrying
    # their anchors onto the preserved paragraph.
    if _needs_passthrough(base) or _needs_passthrough(current):
        return _carry_comment_markers(base, current) if events else copy.deepcopy(current)
    # Identical final text is not a content edit. Re-rendered Markdown often
    # splits the same words into different runs (for example around bold
    # figure references, bookmarks, or subscript spans); diffing those runs
    # creates false deletions/insertions such as ``multi-ionicmulti-ionic`` in
    # Word's All Markup view. Keep the reviewed baseline paragraph intact so
    # only semantic text changes become revisions. Layout-bearing paragraphs
    # are handled by the passthrough branch above.
    if base_text == current_text:
        return copy.deepcopy(base)
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
            if base_has_revisions:
                for offset, token in enumerate(current_tokens[j1:j2], i1):
                    boundary.setdefault(offset, len(nodes))
                    base_token = base_tokens[offset] if offset < len(base_tokens) else None
                    if base_token is not None:
                        nodes.append(_materialize(base_token))
                    else:
                        nodes.append(_run(token))
                    boundary[offset + 1] = len(nodes)
            else:
                _append_tracked(nodes, current_tokens[j1:j2], None, context, boundary, i1)
        elif tag == "delete":
            _append_tracked(nodes, base_tokens[i1:i2], "del", context, boundary, i1)
        elif tag == "insert":
            _append_tracked(nodes, current_tokens[j1:j2], "ins", context, boundary, None)
        else:
            _append_tracked(nodes, base_tokens[i1:i2], "del", context, boundary, i1)
            _append_tracked(nodes, current_tokens[j1:j2], "ins", context, boundary, None)
        boundary[i2] = len(nodes)
    offset_map = {0: 0}
    for index, token in enumerate(base_tokens, 1):
        offset_map[token.end] = index
    placed: dict[int, list[Event]] = {}
    for event in events:
        token_boundary = offset_map.get(event.offset, 0)
        placed.setdefault(boundary.get(token_boundary, len(nodes)), []).append(event)
    if base_has_revisions:
        for offset, revision in _hidden_revision_nodes(base):
            token_boundary = offset_map.get(offset, 0)
            placed.setdefault(boundary.get(token_boundary, len(nodes)), []).append(
                Event("baseline-revision", offset, -1, revision)
            )
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


def _append_block(children: list[etree._Element], element: etree._Element, companion: etree._Element | None) -> None:
    children.append(element)
    if companion is not None:
        children.append(companion)


def _drop_complex_field_codes(element: etree._Element) -> None:
    """Keep a deleted field's visible text and drop its field codes.

    A hyperlink field is several runs (begin, instruction, separate, result,
    end). Wrapping each run in its own ``w:del`` makes Word report the file
    as unreadable. The result text is what the deletion needs to show.
    """
    for run in list(element.findall(".//w:r", NS)):
        if run.find("./w:fldChar", NS) is not None or run.find("./w:instrText", NS) is not None:
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)


def _mark_deleted_runs(element: etree._Element, context: Context) -> None:
    """Mark surviving runs deleted without pulling hyperlinks inside ``w:del``."""
    for child in list(element):
        local = etree.QName(child).localname
        if local in ("pPr", "del", "moveFrom"):
            continue
        if local == "r":
            for text in child.iter(f"{{{W}}}t"):
                text.tag = f"{{{W}}}delText"
            parent = child.getparent()
            if parent is None:
                continue
            index = parent.index(child)
            parent.remove(child)
            wrapper = etree.Element(f"{{{W}}}del", attrib=context.attrs())
            wrapper.append(child)
            parent.insert(index, wrapper)
            continue
        if len(child):
            _mark_deleted_runs(child, context)


_MATH_STRUCTURES = {
    "f", "sSub", "sSup", "sSubSup", "nary", "d", "rad", "acc", "bar", "func",
    "eqArr", "m", "limLow", "limUpp", "groupChr", "box", "borderBox", "sPre", "phant",
}


def _mark_math_structures(element: etree._Element, kind: str, context: Context) -> None:
    """Hide empty math slots when only the text of an equation is revised."""
    tag = "del" if kind == "del" else "ins"
    for node in element.iter():
        if etree.QName(node).localname not in _MATH_STRUCTURES:
            continue
        if etree.QName(node).namespace != M:
            continue
        props_tag = f"{{{M}}}{etree.QName(node).localname}Pr"
        props = node.find(props_tag)
        if props is None:
            props = etree.Element(props_tag)
            node.insert(0, props)
        ctrl = props.find(f"{{{M}}}ctrlPr")
        if ctrl is None:
            ctrl = etree.SubElement(props, f"{{{M}}}ctrlPr")
        rpr = ctrl.find(f"{{{W}}}rPr")
        if rpr is None:
            rpr = etree.SubElement(ctrl, f"{{{W}}}rPr")
        if rpr.find(f"{{{W}}}{tag}") is None:
            rpr.insert(0, etree.Element(f"{{{W}}}{tag}", attrib=context.attrs()))


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
        # A paragraph-mark deletion does not hide drawings. Word keeps rendering
        # the old figure beside the inserted replacement, so drop the picture
        # from the deleted paragraph and leave the deletion on the paragraph mark.
        for drawing in list(result.findall(".//w:drawing", NS)):
            parent = drawing.getparent()
            if parent is not None:
                parent.remove(drawing)
        for pict in list(result.findall(".//w:pict", NS)):
            parent = pict.getparent()
            if parent is not None:
                parent.remove(pict)
        _drop_complex_field_codes(result)
        # The paragraph-mark marker only joins this paragraph into the next
        # one when the revision is accepted. The runs themselves stay live
        # unless they are wrapped in w:del.
        _mark_deleted_runs(result, context)
        _mark_math_structures(result, "del", context)
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
    _mark_math_structures(result, "ins", context)
    return result


def _mark_table(table: etree._Element, kind: str, context: Context) -> etree._Element:
    """Track a whole table through native Word row insertions or deletions."""
    result = copy.deepcopy(table)
    for row in result.findall("./w:tr", NS):
        properties = row.find("./w:trPr", NS)
        if properties is None:
            properties = etree.Element(f"{{{W}}}trPr")
            row.insert(0, properties)
        properties.append(etree.Element(f"{{{W}}}{kind}", attrib=context.attrs()))
    return result


def _base_relationships(base_package: Package, current_package: Package):
    """Copy relationships referenced by deleted baseline paragraphs into the output."""
    rel_name = "word/_rels/document.xml.rels"
    original = base_package.xml(rel_name)
    current = current_package.xml(rel_name)
    # Deleted baseline drawings may reference media formats that are absent
    # from the freshly rendered package (for example TIFF images).  When the
    # media part is carried into the redline package, carry its content-type
    # declaration as well so the resulting DOCX remains a valid OPC package.
    base_content_types = base_package.xml("[Content_Types].xml")
    current_content_types = current_package.xml("[Content_Types].xml")
    content_type_defaults = {
        node.get("Extension", "").lower(): node.get("ContentType", "")
        for node in current_content_types.findall(f"{{{CONTENT_TYPES}}}Default")
    }
    content_type_overrides = {
        node.get("PartName", ""): node.get("ContentType", "")
        for node in current_content_types.findall(f"{{{CONTENT_TYPES}}}Override")
    }

    def ensure_media_content_type(target: str) -> None:
        part_name = "/" + target.replace("\\", "/")
        suffix = posixpath.splitext(part_name)[1].lstrip(".").lower()
        if suffix in content_type_defaults or part_name in content_type_overrides:
            return
        base_part_name = part_name.replace("/redline_base_", "/image", 1)
        base_override = next(
            (
                node
                for node in base_content_types.findall(f"{{{CONTENT_TYPES}}}Override")
                if node.get("PartName", "").lower() == base_part_name.lower()
            ),
            None,
        )
        if base_override is not None:
            current_content_types.append(copy.deepcopy(base_override))
            current_content_types[-1].set("PartName", part_name)
            content_type_overrides[part_name] = base_override.get("ContentType", "")
            return
        base_default = next(
            (
                node
                for node in base_content_types.findall(f"{{{CONTENT_TYPES}}}Default")
                if node.get("Extension", "").lower() == suffix
            ),
            None,
        )
        if base_default is not None:
            current_content_types.append(copy.deepcopy(base_default))
            content_type_defaults[suffix] = base_default.get("ContentType", "")
    sources = {rel.get("Id"): rel for rel in original}
    used = {rel.get("Id") for rel in current}
    copied: dict[str, str] = {}

    def carry(paragraph: etree._Element) -> etree._Element:
        result = copy.deepcopy(paragraph)
        for node in result.iter():
            for attr in (f"{{{R}}}id", f"{{{R}}}embed", f"{{{R}}}link"):
                old_id = node.get(attr)
                if old_id is None:
                    continue
                if old_id not in sources:
                    raise ValueError(f"Missing baseline relationship: {old_id}")
                if old_id not in copied:
                    rel = copy.deepcopy(sources[old_id])
                    number = 1
                    while f"rId{number}" in used:
                        number += 1
                    new_id = f"rId{number}"
                    used.add(new_id)
                    rel.set("Id", new_id)
                    if rel.get("TargetMode") != "External":
                        target = posixpath.normpath(posixpath.join("word", rel.get("Target", "")))
                        if not target.startswith("word/media/") or target not in base_package.parts:
                            raise ValueError(f"Unsupported baseline relationship target: {target}")
                        suffix = posixpath.splitext(target)[1]
                        new_target = f"media/redline_base_{number}{suffix}"
                        current_package.parts[f"word/{new_target}"] = base_package.parts[target]
                        ensure_media_content_type(f"word/{new_target}")
                        rel.set("Target", new_target)
                    current.append(rel)
                    copied[old_id] = new_id
                node.set(attr, copied[old_id])
        return result

    def save() -> None:
        current_package.set_xml(rel_name, current)
        current_package.set_xml("[Content_Types].xml", current_content_types)

    return carry, save


def _drawing_payloads(paragraph: etree._Element, package: Package) -> list[bytes]:
    rels = {rel.get("Id"): rel for rel in package.xml("word/_rels/document.xml.rels")}
    result = []
    for node in paragraph.findall(".//a:blip", {"a": A}):
        rel = rels.get(node.get(f"{{{R}}}embed"))
        if rel is not None and rel.get("TargetMode") != "External":
            target = posixpath.normpath(posixpath.join("word", rel.get("Target", "")))
            result.append(package.parts[target])
    return result


def _table_is_deleted(table: etree._Element) -> bool:
    rows = table.findall("./w:tr", NS)
    return bool(rows) and all(row.find("./w:trPr/w:del", NS) is not None for row in rows)


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


def _row_revision(row: etree._Element, kind: str, context: Context) -> etree._Element:
    """Mark a table row with a valid Word row-level revision property."""
    result = copy.deepcopy(row)
    tr_pr = result.find("./w:trPr", NS)
    if tr_pr is None:
        tr_pr = etree.Element(f"{{{W}}}trPr")
        result.insert(0, tr_pr)
    # A row can carry at most one insertion/deletion marker in a single pass.
    for name in ("ins", "del", "moveFrom", "moveTo"):
        for marker in list(tr_pr.findall(f"./w:{name}", NS)):
            tr_pr.remove(marker)
    tr_pr.append(etree.Element(f"{{{W}}}{kind}", attrib=context.attrs()))
    return result


def _merge_row_cells(
    base_row: etree._Element,
    current_row: etree._Element,
    context: Context,
    carry_base,
    *,
    preserve_base_revisions: bool = False,
) -> etree._Element:
    """Merge corresponding cell paragraphs while retaining current row XML."""
    result = copy.deepcopy(current_row)
    base_cells = base_row.findall("./w:tc", NS)
    current_cells = result.findall("./w:tc", NS)
    if len(base_cells) != len(current_cells):
        return result
    for base_cell, current_cell in zip(base_cells, current_cells):
        base_paragraphs = base_cell.findall("./w:p", NS)
        current_paragraphs = current_cell.findall("./w:p", NS)
        if len(base_paragraphs) != len(current_paragraphs):
            continue
        for base_paragraph, current_paragraph in zip(base_paragraphs, current_paragraphs):
            if (visible_text(base_paragraph) != visible_text(current_paragraph) and
                    (_needs_passthrough(base_paragraph) or _needs_passthrough(current_paragraph))):
                parent = current_paragraph.getparent()
                index = parent.index(current_paragraph)
                parent.remove(current_paragraph)
                parent.insert(index, _mark_paragraph(carry_base(base_paragraph), "del", context))
                parent.insert(index + 1, _mark_paragraph(current_paragraph, "ins", context))
                continue
            merged = _merge_paragraph(
                base_paragraph,
                current_paragraph,
                context,
                preserve_base_revisions=preserve_base_revisions,
            )
            current_paragraph.getparent().replace(current_paragraph, merged)
    return result


def _merge_table(
    base: etree._Element,
    current: etree._Element,
    context: Context,
    carry_base,
    *,
    preserve_base_revisions: bool = False,
) -> tuple[etree._Element, bool]:
    """Diff table rows and cell text, tracking inserted/deleted rows natively."""
    result = copy.deepcopy(current)
    base_rows = base.findall("./w:tr", NS)
    current_rows = result.findall("./w:tr", NS)
    if not base_rows or not current_rows:
        return result, False

    base_keys = [_normalize(visible_text(row)) for row in base_rows]
    current_keys = [_normalize(visible_text(row)) for row in current_rows]
    matcher = difflib.SequenceMatcher(None, base_keys, current_keys, autojunk=False)
    output_rows: list[etree._Element] = []
    row_revision = False
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for base_row, current_row in zip(base_rows[i1:i2], current_rows[j1:j2]):
                output_rows.append(
                    _merge_row_cells(
                        base_row,
                        current_row,
                        context,
                        carry_base,
                        preserve_base_revisions=preserve_base_revisions,
                    )
                )
        elif tag == "replace" and (i2 - i1) == (j2 - j1):
            # Same geometry, changed cell text: preserve row formatting and
            # expose the cell-level insertions/deletions.
            for base_row, current_row in zip(base_rows[i1:i2], current_rows[j1:j2]):
                output_rows.append(
                    _merge_row_cells(
                        base_row,
                        current_row,
                        context,
                        carry_base,
                        preserve_base_revisions=preserve_base_revisions,
                    )
                )
        else:
            for base_row in base_rows[i1:i2]:
                output_rows.append(_row_revision(base_row, "del", context))
                row_revision = True
            for current_row in current_rows[j1:j2]:
                output_rows.append(_row_revision(current_row, "ins", context))
                row_revision = True

    # Replace only the row children; table properties, grid, bookmarks and
    # relationships remain sourced from the current DOCX.
    for row in result.findall("./w:tr", NS):
        result.remove(row)
    insert_at = len(result)
    for index, child in enumerate(result):
        if child.tag == f"{{{W}}}tblGrid":
            insert_at = index + 1
    for offset, row in enumerate(output_rows):
        result.insert(insert_at + offset, row)
    return result, row_revision


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


def _carry_missing_comment_markers(raw: etree._Element, merged: etree._Element) -> None:
    """Carry anchors from accepted-away paragraphs to corresponding text.

    A paragraph can disappear while accepting an earlier review pass.  Word's
    comment part still contains that thread, so dropping its XML anchors makes
    the comment look orphaned.  Match the old paragraph text first; when the
    text was deleted, attach the anchor to the nearest surviving paragraph in
    document order.  The old implementation used the first paragraph for all
    misses, which visibly moved unrelated comments to the document cover.
    """
    wanted = _comment_counts(raw)
    present = _comment_counts(merged)
    missing_ids = {
        identifier for identifier, counts in wanted.items()
        if present.get(identifier, (0, 0, 0)) != counts
    }
    if not missing_ids:
        return
    raw_paragraphs = raw.findall(".//w:body/w:p", NS)
    current_body = merged.find(".//w:body", NS)
    if current_body is None:
        return
    current_paragraphs = current_body.findall("./w:p", NS)
    if not current_paragraphs:
        return
    targets: dict[str, etree._Element] = {}
    for identifier in missing_ids:
        old_index = None
        old_text = ""
        for index, paragraph in enumerate(raw_paragraphs):
            if any(node.get(f"{{{W}}}id") == identifier for node in paragraph.findall(".//w:commentRangeStart", NS) + paragraph.findall(".//w:commentRangeEnd", NS) + paragraph.findall(".//w:commentReference", NS)):
                old_index, old_text = index, _normalize(visible_text(paragraph))
                break
        target = None
        if old_text:
            for paragraph in current_paragraphs:
                if _normalize(visible_text(paragraph)) == old_text:
                    target = paragraph
                    break
        if target is None:
            index = old_index if old_index is not None else 0
            # Prefer a corresponding paragraph after a deleted first block;
            # never relocate a review thread to the cover's first paragraph.
            if index == 0 and len(current_paragraphs) > 1:
                index = 1
            target = current_paragraphs[min(index, len(current_paragraphs) - 1)]
        targets[identifier] = target
    by_kind = (("commentRangeStart", 0), ("commentRangeEnd", 1), ("commentReference", 2))
    for name, position in by_kind:
        for node in raw.findall(f".//w:{name}", NS):
            identifier = node.get(f"{{{W}}}id", "")
            if identifier not in targets:
                continue
            counts = present.get(identifier, (0, 0, 0))
            if counts[position] >= wanted.get(identifier, (0, 0, 0))[position]:
                continue
            target = targets[identifier]
            clone = copy.deepcopy(node)
            if name == "commentRangeStart":
                insert_at = 1 if target.find("./w:pPr", NS) is not None else 0
                target.insert(insert_at, clone)
            else:
                target.append(clone)
            updated = list(counts)
            updated[position] += 1
            present[identifier] = tuple(updated)


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
        if element.tag == f"{{{W}}}tbl" and _table_is_deleted(element):
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
        elif node.tag == f"{{{W}}}tr":
            tr_pr = node.find("./w:trPr", NS)
            if tr_pr is not None and tr_pr.find("./w:del", NS) is not None:
                hidden = True
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
        if block.tag == f"{{{W}}}tbl" and _table_is_deleted(block):
            continue
        items: list[tuple] = []
        walk(block, False, items)
        # Empty paragraphs carry no controls whose accepted-view identity can
        # be audited here.  Redline alignment may legitimately retain a blank
        # paragraph around a moved figure or section break, so compare only
        # paragraphs that actually contain a drawing, break, tab, or section
        # definition.
        if items:
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

    # Word tracks whole-table changes by marking each row. Removing the last
    # deleted row also removes its now-empty table from the accepted view.
    for table in list(result.findall(".//w:tbl", NS)):
        for row in list(table.findall("./w:tr", NS)):
            if row.find("./w:trPr/w:del", NS) is not None:
                table.remove(row)
        if not table.findall("./w:tr", NS):
            parent = table.getparent()
            if parent is not None:
                parent.remove(table)

    # Paragraph deletions use a revision marker in pPr rather than a wrapper
    # around the paragraph text. Remove those paragraphs before stripping the
    # remaining revision metadata, otherwise they would survive as empty blocks.
    for paragraph in list(result.xpath(".//w:p[w:pPr/w:rPr/w:del]", namespaces=NS)):
        parent = paragraph.getparent()
        if parent is not None:
            parent.remove(paragraph)

    # Table-row revisions are stored in w:trPr rather than as wrappers.
    for row in list(result.xpath(".//w:tr[w:trPr/w:del or w:trPr/w:moveFrom]", namespaces=NS)):
        parent = row.getparent()
        if parent is not None:
            parent.remove(row)
    for row in result.xpath(".//w:tr[w:trPr/w:ins or w:trPr/w:moveTo]", namespaces=NS):
        tr_pr = row.find("./w:trPr", NS)
        if tr_pr is not None:
            for name in ("ins", "moveTo"):
                for marker in list(tr_pr.findall(f"./w:{name}", NS)):
                    tr_pr.remove(marker)

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
    preserve_base_revisions: bool = True,
    ratio_cache: Path | None = None,
    workers: int | None = None,
    accept_baseline: bool = False,
    report_path: Path | None = None,
) -> dict[str, int]:
    """Write ``output_path`` as ``current_path`` tracked against ``base_path``.

    ``ratio_cache`` persists paragraph similarity ratios between runs, and
    ``workers`` sets the process count for ratios not yet cached (``1`` keeps
    the computation in this process). Neither changes the result. Set
    ``accept_baseline`` to materialize earlier baseline revisions first;
    ``report_path`` optionally receives the alignment decisions.
    """
    validate_output_path(output_path)
    validate_output_path(report_path, label="report")
    base_package = Package.load(base_path)
    current_package = Package.load(current_path)
    raw_base_root = base_package.xml("word/document.xml")
    # Keep the reviewed author's w:ins/w:del tree. Alignment still uses the
    # final view, so only new differences are added on top of those traces.
    baseline_preserve = preserve_base_revisions and not accept_baseline
    base_root = raw_base_root if baseline_preserve else _accepted_revision_view(raw_base_root)
    # A freshly supplied current document may itself contain an earlier
    # review pass.  Diff only its accepted view so stale w:ins/w:del and
    # formatting-change markers cannot leak into the new redline.
    raw_current_root = current_package.xml("word/document.xml")
    current_root = _accepted_revision_view(raw_current_root)
    base_body, base_blocks = _blocks(base_root)
    current_body, current_blocks = _blocks(current_root)
    # Converter leftovers such as "suppnote" are not manuscript text. Leaving
    # them in the baseline makes Word paint them on top of the new cover.
    base_blocks = [
        block for block in base_blocks
        if block.drawing or not _NOISE_PARAGRAPH.fullmatch(_normalize(block.text))
    ]
    context = Context(author, _next_id(raw_base_root))
    carry_base, save_base_relationships = _base_relationships(base_package, current_package)
    summary = {"matched": 0, "changed": 0, "inserted": 0, "deleted": 0, "tables_replaced": 0}
    children: list[etree._Element] = []
    aligned = _align(
        base_blocks, current_blocks, cache_path=ratio_cache, workers=workers
    )
    alignment_report: list[dict[str, object]] = []
    for action, base_index, current_index in aligned:
        old = base_blocks[base_index] if base_index is not None else None
        new = current_blocks[current_index] if current_index is not None else None
        alignment_report.append({
            "action": action,
            "baseline_index": base_index,
            "current_index": current_index,
            "baseline_text": old.text if old is not None else None,
            "current_text": new.text if new is not None else None,
        })
    matched = sum(action == "match" for action, _, _ in aligned)
    # A rewritten document has too few shared paragraphs to interleave. Word
    # then drops old sentences into the new cover and headings. Keep the new
    # pages intact and put the old document after them as deletions.
    block_replace = matched < 0.25 * max(len(aligned), 1)
    if block_replace:
        aligned = [("insert", None, index) for index in range(len(current_blocks))]
        aligned += [("delete", index, None) for index in range(len(base_blocks))]
    for action, base_index, current_index in aligned:
        if action == "match":
            old = base_blocks[base_index]  # type: ignore[index]
            new = current_blocks[current_index]  # type: ignore[index]
            if old.kind == "p" and (old.companion is not None or new.companion is not None):
                _append_block(children, copy.deepcopy(new.element), None)
                if old.companion is not None and new.companion is not None:
                    children.append(_merge_paragraph(
                        old.companion, new.companion, context,
                        preserve_base_revisions=baseline_preserve,
                    ))
                elif new.companion is not None:
                    children.append(copy.deepcopy(new.companion))
                summary["matched" if old.text == new.text else "changed"] += 1
            elif old.kind == "p":
                changed_drawing = (old.drawing or new.drawing) and (
                    _drawing_payloads(old.element, base_package)
                    != _drawing_payloads(new.element, current_package)
                )
                if (_normalize(old.text) != _normalize(new.text) and (_needs_passthrough(old.element) or _needs_passthrough(new.element))) or changed_drawing:
                    children.append(_mark_paragraph(carry_base(old.element), "del", context))
                    children.append(_mark_paragraph(new.element, "ins", context))
                    summary["changed"] += 1
                    continue
                children.append(
                    _carry_comment_markers(old.element, new.element)
                    if old.drawing or new.drawing
                    else _merge_paragraph(
                        old.element,
                        new.element,
                        context,
                        preserve_base_revisions=baseline_preserve,
                    )
                )
                summary["matched" if old.text == new.text else "changed"] += 1
            elif old.kind == "tbl":
                merged_table, row_revision = _merge_table(
                    old.element,
                    new.element,
                    context,
                    carry_base,
                    preserve_base_revisions=baseline_preserve,
                )
                children.append(merged_table)
                if _normalize(old.text) == _normalize(new.text):
                    summary["matched"] += 1
                elif row_revision:
                    # Row-level ``w:ins``/``w:del`` markers make the table
                    # diff reviewable; it is not an opaque table replacement.
                    summary["changed"] += 1
                else:
                    summary["tables_replaced"] += 1
            else:
                children.append(
                    copy.deepcopy(old.element if _normalize(old.text) == _normalize(new.text) else new.element)
                )
                summary["matched" if _normalize(old.text) == _normalize(new.text) else "tables_replaced"] += 1
        elif action == "delete":
            old = base_blocks[base_index]  # type: ignore[index]
            if old.kind == "p":
                _append_block(
                    children,
                    _mark_paragraph(carry_base(old.element), "del", context),
                    _mark_paragraph(carry_base(old.companion), "del", context) if old.companion is not None else None,
                )
            elif old.kind == "tbl":
                children.append(_mark_table(old.element, "del", context))
            else:
                children.append(copy.deepcopy(old.element))
            summary["deleted"] += 1
        else:
            new = current_blocks[current_index]  # type: ignore[index]
            if new.kind == "p":
                _append_block(
                    children,
                    _mark_paragraph(new.element, "ins", context),
                    _mark_paragraph(new.companion, "ins", context) if new.companion is not None else None,
                )
            elif new.kind == "tbl":
                children.append(_mark_table(new.element, "ins", context))
            else:
                children.append(copy.deepcopy(new.element))
            summary["inserted"] += 1
    # Interleaved redlines need a break after each caption so a deleted heading
    # cannot sit on the figure. A block replacement already keeps the new pages
    # intact, and an extra break there opens a blank page.
    if not block_replace:
        for index, child in enumerate(children[1:], start=1):
            previous = visible_text(children[index - 1]).strip()
            if child.tag != f"{{{W}}}p" or not re.match(r"Figure S\d+\.", previous):
                continue
            props = child.find("./w:pPr", NS)
            if props is None:
                props = etree.Element(f"{{{W}}}pPr")
                child.insert(0, props)
            if props.find("./w:pageBreakBefore", NS) is None:
                etree.SubElement(props, f"{{{W}}}pageBreakBefore")
    current_sectpr = current_body.find("./w:sectPr", NS)
    for child in list(current_body):
        current_body.remove(child)
    for child in children:
        current_body.append(child)
    if current_sectpr is not None:
        current_body.append(copy.deepcopy(current_sectpr))
    _carry_missing_comment_markers(raw_base_root, current_root)
    original_markers = _comment_counts(raw_base_root)
    final_markers = _comment_counts(current_root)
    if final_markers != original_markers:
        missing = {key: value for key, value in original_markers.items() if final_markers.get(key) != value}
        if missing:
            raise AssertionError(f"Comment anchors were not preserved: {missing}")
    current_package.set_xml("word/document.xml", current_root)
    save_base_relationships()
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
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(
                {
                    "schema": "docforge.redline.v1",
                    "baseline": str(base_path),
                    "current": str(current_path),
                    "output": str(output_path),
                    "accept_baseline": accept_baseline,
                    "revision_author": author,
                    "actions": alignment_report,
                    "summary": summary,
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
    return summary


__all__ = ["create_tracked_docx", "Package"]
