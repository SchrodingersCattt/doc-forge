"""Flatten WordprocessingML property elements into ordered text tokens.

A property element such as ``w:pPr`` becomes a list of ``(key, value)``
tokens.  ``spacing@after`` addresses the ``w:after`` attribute of the
``w:spacing`` child, ``jc@val`` its ``w:val`` attribute (written ``jc=...``),
``keepNext`` an empty child, and ``rPr.rFonts@cs`` a nested attribute.  A
repeated sibling is addressed as ``tab#2``.  Rebuilding sorts children in
schema order, so a flatten/rebuild cycle reproduces the element.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from typing import Iterable, Iterator

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

KNOWN_NAMESPACES = {
    "w": W,
    "r": R,
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "xml": XML_NS,
}

Token = tuple[str, "str | None"]

_ORDERS: dict[str, tuple[str, ...]] = {
    "pPr": (
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
        "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
        "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
        "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
        "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    ),
    "rPr": (
        "ins", "del", "moveFrom", "moveTo",
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
        "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
        "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
        "eastAsianLayout", "specVanish", "oMath", "rPrChange",
    ),
    "sectPr": (
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz",
        "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
        "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
        "printerSettings", "sectPrChange",
    ),
    "tblPr": (
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
        "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
    ),
    "trPr": (
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
        "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange",
    ),
    "tcPr": (
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
        "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns",
        "cellDel", "cellMerge", "tcPrChange",
    ),
    "tblBorders": ("top", "left", "start", "bottom", "right", "end", "insideH", "insideV"),
    "tcBorders": (
        "top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl",
    ),
    "pBdr": ("top", "left", "bottom", "right", "between", "bar"),
    "tblCellMar": ("top", "left", "start", "bottom", "right", "end"),
    "tcMar": ("top", "left", "start", "bottom", "right", "end"),
    "numPr": ("ilvl", "numId", "numberingChange", "ins"),
}

_BARE_VALUE = re.compile(r'^[^\s"{}|]+$')


def qname(prefix_name: str, nsmap: dict[str, str] | None = None) -> str:
    """Clark name for ``w:name``-style text; an unprefixed name is ``w:``."""
    if ":" in prefix_name:
        prefix, local = prefix_name.split(":", 1)
        if prefix == "":
            return local
        uri = (nsmap or {}).get(prefix) or KNOWN_NAMESPACES.get(prefix)
        if uri is None:
            raise ValueError(f"Unknown namespace prefix: {prefix!r}")
        return f"{{{uri}}}{local}"
    return f"{{{W}}}{prefix_name}"


def short_name(name: str, node: etree._Element | None = None, *, attribute: bool = False) -> str:
    """Inverse of :func:`qname`: ``w:`` names become bare local names."""
    q = etree.QName(name)
    if q.namespace is None:
        return f":{q.localname}" if attribute else q.localname
    if q.namespace == W:
        return q.localname
    prefix = None
    if node is not None:
        for key, value in node.nsmap.items():
            if value == q.namespace and key:
                prefix = key
                break
    if prefix is None:
        for key, value in KNOWN_NAMESPACES.items():
            if value == q.namespace:
                prefix = key
                break
    if prefix is None:
        raise ValueError(f"No prefix for namespace {q.namespace}")
    return f"{prefix}:{q.localname}"


def flatten(element: etree._Element | None, *, skip: Iterable[str] = ()) -> list[Token]:
    """Return ordered tokens describing the children of ``element``."""
    if element is None:
        return []
    skipped = set(skip)
    tokens: list[Token] = []
    _flatten_into(element, "", tokens, skipped)
    return tokens


def _flatten_into(element: etree._Element, prefix: str, tokens: list[Token], skipped: set[str]) -> None:
    counts: dict[str, int] = {}
    for child in element:
        if not isinstance(child.tag, str):
            continue
        name = short_name(child.tag, child)
        if not prefix and name in skipped:
            continue
        counts[name] = counts.get(name, 0) + 1
        segment = name if counts[name] == 1 else f"{name}#{counts[name]}"
        path = prefix + segment
        attributes = [(key, value) for key, value in child.attrib.items()]
        children = [node for node in child if isinstance(node.tag, str)]
        if not attributes and not children:
            tokens.append((path, None))
        for key, value in attributes:
            tokens.append((f"{path}@{short_name(key, child, attribute=True)}", value))
        if children:
            _flatten_into(child, path + ".", tokens, set())


def _split_key(key: str) -> tuple[list[tuple[str, int]], str | None]:
    attribute = None
    if "@" in key:
        key, attribute = key.split("@", 1)
    segments: list[tuple[str, int]] = []
    for segment in key.split("."):
        if "#" in segment:
            name, index = segment.split("#", 1)
            segments.append((name, int(index)))
        else:
            segments.append((segment, 1))
    return segments, attribute


def build(tag: str, tokens: Iterable[Token], nsmap: dict[str, str] | None = None) -> etree._Element:
    """Rebuild a property element from tokens, children in schema order."""
    root = etree.Element(tag, nsmap=nsmap)
    index: dict[tuple, etree._Element] = {}
    for key, value in tokens:
        segments, attribute = _split_key(key)
        parent = root
        path: tuple = ()
        for name, number in segments:
            path = path + ((name, number),)
            node = index.get(path)
            if node is None:
                node = etree.SubElement(parent, qname(name, nsmap))
                index[path] = node
            parent = node
        if attribute is not None:
            parent.set(qname(attribute, nsmap) if not attribute.startswith(":") else attribute[1:], value or "")
    _sort_children(root)
    return root


def _sort_children(element: etree._Element) -> None:
    order = _ORDERS.get(etree.QName(element).localname)
    children = list(element)
    if order and len(children) > 1:
        position = {name: rank for rank, name in enumerate(order)}

        def rank(item: tuple[int, etree._Element]) -> tuple[int, int]:
            number, child = item
            q = etree.QName(child)
            known = position.get(q.localname) if q.namespace == W else None
            return (known if known is not None else len(order), number)

        ordered = [child for _, child in sorted(enumerate(children), key=rank)]
        for child in children:
            element.remove(child)
        for child in ordered:
            element.append(child)
    for child in element:
        _sort_children(child)


def as_dict(tokens: Iterable[Token]) -> "OrderedDict[str, str | None]":
    result: OrderedDict[str, str | None] = OrderedDict()
    for key, value in tokens:
        result[key] = value
    return result


def _covers(removal: str, key: str) -> bool:
    return key == removal or key.startswith(removal + "@") or key.startswith(removal + ".")


def merge(base: Iterable[Token], override: Iterable[Token]) -> list[Token]:
    """Apply ``override`` (``-key`` removals first, then sets) to ``base``."""
    result = as_dict(base)
    sets: list[Token] = []
    for key, value in override:
        if key.startswith("-"):
            removal = key[1:]
            for existing in [item for item in result if _covers(removal, item)]:
                del result[existing]
        else:
            sets.append((key, value))
    for key, value in sets:
        result[key] = value
    return list(result.items())


def diff(base: Iterable[Token], target: Iterable[Token]) -> list[Token]:
    """Override tokens such that ``merge(base, diff) == target`` as a mapping."""
    base_map = as_dict(base)
    target_map = as_dict(target)
    removals = [key for key in base_map if key not in target_map]
    removed = {key for key in base_map if any(_covers(item, key) for item in removals)}
    result: list[Token] = [("-" + key, None) for key in removals]
    for key, value in target_map.items():
        if key not in base_map or base_map[key] != value or key in removed:
            result.append((key, value))
    return result


def strip_keys(tokens: Iterable[Token], predicate) -> list[Token]:
    return [(key, value) for key, value in tokens if not predicate(key)]


# ----------------------------------------------------------------------------
# Text form of token lists: ``spacing@after=240 keepNext jc=center -b``.

def format_value(value: str) -> str:
    if value and _BARE_VALUE.match(value):
        return value
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def format_tokens(tokens: Iterable[Token], prefix: str = "") -> str:
    parts = []
    for key, value in tokens:
        if key.startswith("-"):
            parts.append(f"-{prefix}{key[1:]}")
            continue
        if value is None:
            parts.append(prefix + key)
        elif key.endswith("@val"):
            parts.append(f"{prefix}{key[:-4]}={format_value(value)}")
        else:
            parts.append(f"{prefix}{key}={format_value(value)}")
    return " ".join(parts)


def split_words(text: str) -> Iterator[str]:
    """Split on whitespace outside double quotes, keeping the quotes."""
    word: list[str] = []
    quoted = False
    index = 0
    while index < len(text):
        char = text[index]
        if quoted:
            word.append(char)
            if char == "\\" and index + 1 < len(text):
                word.append(text[index + 1])
                index += 1
            elif char == '"':
                quoted = False
        elif char == '"':
            quoted = True
            word.append(char)
        elif char.isspace():
            if word:
                yield "".join(word)
                word = []
        else:
            word.append(char)
        index += 1
    if quoted:
        raise ValueError(f"Unclosed quote in: {text!r}")
    if word:
        yield "".join(word)


def unquote(value: str) -> str:
    if len(value) >= 2 and value[0] == value[-1] == '"':
        body = value[1:-1]
        return re.sub(r"\\(.)", r"\1", body)
    return value


def parse_token(word: str) -> Token:
    if word.startswith("-"):
        return (word, None)
    if "=" not in word:
        return (word, None)
    key, value = word.split("=", 1)
    if "@" not in key:
        key += "@val"
    return (key, unquote(value))


def parse_tokens(text: str) -> list[Token]:
    return [parse_token(word) for word in split_words(text)]


__all__ = [
    "KNOWN_NAMESPACES", "R", "W", "XML_NS", "PKG_REL", "CONTENT_TYPES", "Token", "format_value",
    "as_dict", "build", "diff", "flatten", "format_tokens", "merge", "parse_token",
    "parse_tokens", "qname", "short_name", "split_words", "strip_keys", "unquote",
]
