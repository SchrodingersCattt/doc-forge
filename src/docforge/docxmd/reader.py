"""DOCX -> Markdown bundle.

The bundle keeps one Markdown copy of every part of the document plus a
references JSON, extracted media, and the source package that supplies
styles, numbering, headers, footers and settings when rebuilding.
"""

from __future__ import annotations

import copy
import hashlib
import json
import re
import shutil
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath

from lxml import etree

from ..docxdiff.package import Package
from . import drawing as drawing_mod
from .common import (
    BIB_SCHEMA, FORCED, SCHEMA, SHORTCUT_ELEMENT, Attrs, clean_tokens, format_attrs,
    format_citation_numbers, parse_citation_numbers, run_tokens_for, top_name,
)
from .inline import Atom, Text, normalize, write as write_inline
from .ooxml import KNOWN_NAMESPACES, PKG_REL, R, W, Token, diff, flatten, format_tokens, merge
from .styles import StyleSheet

W15 = KNOWN_NAMESPACES["w15"]
_REFERENCE_HEADING = re.compile(r"^\s*(references?|bibliography|literature cited|参考文献)\s*$", re.I)
_CITATION_PRECEDERS = set(".,;:”’\"")
_ROLE_KEYWORDS = (
    ("table", None), ("caption", None), ("corresponding", "correspondence"),
    ("authorinfo", None), ("footnote", None), ("subtitle", None), ("synopsis", None),
    ("title", "title"), ("address", "affiliation"), ("affiliation", "affiliation"),
    ("author", "authors"), ("abstract", "abstract"), ("keyword", "keywords"),
)


def _local(node: etree._Element) -> str:
    return etree.QName(node).localname if isinstance(node.tag, str) else ""


def _sample_text(text: str) -> str:
    """One character per font class present in ``text`` (visibility cache key)."""
    sample = []
    if any("a" <= char.lower() <= "z" for char in text):
        sample.append("a")
    if any(ord(char) < 0x80 and not char.isspace() and not "a" <= char.lower() <= "z" for char in text):
        sample.append("-")
    if any(0x80 <= ord(char) < 0x2E80 and not char.isspace() for char in text):
        sample.append("é")
    if any(0x2E80 <= ord(char) and not (0xAC00 <= ord(char) <= 0xD7AF) for char in text):
        sample.append("中")
    if any(0x0590 <= ord(char) <= 0x08FF for char in text):
        sample.append("א")
    if any(char.isspace() for char in text):
        sample.append(" ")
    return "".join(sample)


@dataclass
class Piece:
    kind: str  # "text" | "atom" | "milestone"
    value: object
    rpr: list[Token]
    wrappers: tuple = ()


@dataclass
class Para:
    element: etree._Element
    ppr: list[Token]
    style: str | None
    table_style: str | None
    pieces: list[Piece]
    base: list[Token] | None
    signature: tuple = ()
    cls: str | None = None
    override: list[Token] = field(default_factory=list)
    r_override: list[Token] = field(default_factory=list)
    level: int | None = None

    @property
    def text(self) -> str:
        return "".join(piece.value for piece in self.pieces if piece.kind == "text")


@dataclass
class ExportOptions:
    doc_id: str
    split: list[dict] | None = None
    bibliography: str = "references.json"
    bib_style: str | None = None
    citations: bool = True


_AUTO_BOOKMARK = re.compile(r"^(?:OLE_LINK\d+|_Hlk\d+|_GoBack)$")


def kept_bookmarks(root: etree._Element) -> dict[str, str]:
    """Bookmark id -> name for bookmarks worth keeping.

    Word adds ``OLE_LINK``/``_Hlk``/``_GoBack`` bookmarks on its own; they
    are dropped unless a field or hyperlink refers to them.
    """
    referenced = {node.get(f"{{{W}}}anchor") for node in root.iter(f"{{{W}}}hyperlink")}
    for node in root.iter(f"{{{W}}}instrText", f"{{{W}}}fldSimple"):
        code = node.text if node.tag == f"{{{W}}}instrText" else node.get(f"{{{W}}}instr")
        referenced.update(re.findall(r"[A-Za-z_][\w]*", code or ""))
    starts = {}
    for node in root.iter(f"{{{W}}}bookmarkStart"):
        name = node.get(f"{{{W}}}name") or ""
        if name and (not _AUTO_BOOKMARK.match(name) or name in referenced):
            starts[node.get(f"{{{W}}}id")] = name
    return starts


def _hoist_body_bookmarks(body: etree._Element) -> None:
    """Move body-level bookmark marks into the neighbouring paragraph."""
    for node in list(body):
        if _local(node) not in {"bookmarkStart", "bookmarkEnd"}:
            continue
        if _local(node) == "bookmarkStart":
            target = node.getnext()
            while target is not None and _local(target) != "p":
                target = target.getnext()
            if target is not None:
                ppr = target.find(f"./{{{W}}}pPr")
                target.insert(0 if ppr is None else 1, node)
                continue
        else:
            target = node.getprevious()
            while target is not None and _local(target) != "p":
                target = target.getprevious()
            if target is not None:
                target.append(node)
                continue


class _Reader:
    def __init__(self, source: Path, bundle: Path, options: ExportOptions) -> None:
        self.source = source
        self.bundle = bundle
        self.options = options
        self.package = Package.load(source)
        self.root = self.package.xml("word/document.xml")
        self.body = self.root.find(f"./{{{W}}}body")
        styles = self.package.parts.get("word/styles.xml")
        self.styles = StyleSheet(etree.fromstring(styles) if styles else None)
        self.rels = self._rels("word/_rels/document.xml.rels")
        self.media: dict[str, str] = {}
        self.visible_cache: dict = {}
        self.mark_cache: dict = {}
        self.comment_ids: dict[str, str] = {}
        self.range_ends = {
            node.get(f"{{{W}}}id") for node in self.root.iter(f"{{{W}}}commentRangeEnd")
        }
        self.comments = self._comments()
        self.report: dict[str, object] = {"raw": [], "dropped": Counter()}
        self.citation_keys: dict[int, str] = {}
        self.bookmarks = kept_bookmarks(self.root)
        _hoist_body_bookmarks(self.body)

    # -- package helpers ---------------------------------------------------
    def _rels(self, name: str) -> dict[str, tuple[str, str, str | None]]:
        if name not in self.package.parts:
            return {}
        root = etree.fromstring(self.package.parts[name])
        return {
            rel.get("Id"): (rel.get("Type"), rel.get("Target"), rel.get("TargetMode"))
            for rel in root.findall(f"{{{PKG_REL}}}Relationship")
        }

    def _media_path(self, rel_id: str) -> str:
        _, target, _ = self.rels[rel_id]
        part = str(PurePosixPath("word") / target) if not target.startswith("/") else target[1:]
        part = str(PurePosixPath(part))
        normalized = re.sub(r"[^/]+/\.\./", "", part)
        if normalized not in self.media:
            name = PurePosixPath(normalized).name
            relative = f"media/{self.options.doc_id}/{name}"
            destination = self.bundle / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(self.package.parts[normalized])
            self.media[normalized] = relative
        return self.media[normalized]

    def _comments(self) -> dict[str, dict]:
        if "word/comments.xml" not in self.package.parts:
            return {}
        root = etree.fromstring(self.package.parts["word/comments.xml"])
        extended: dict[str, dict] = {}
        if "word/commentsExtended.xml" in self.package.parts:
            ext = etree.fromstring(self.package.parts["word/commentsExtended.xml"])
            for item in ext.iter(f"{{{W15}}}commentEx"):
                extended[item.get(f"{{{W15}}}paraId")] = {
                    "done": item.get(f"{{{W15}}}done"),
                    "parent": item.get(f"{{{W15}}}paraIdParent"),
                }
        para_owner: dict[str, str] = {}
        result: dict[str, dict] = {}
        for comment in root.findall(f"{{{W}}}comment"):
            cid = comment.get(f"{{{W}}}id")
            paragraphs = comment.findall(f"{{{W}}}p")
            info = {
                "author": comment.get(f"{{{W}}}author", ""),
                "initials": comment.get(f"{{{W}}}initials", ""),
                "date": comment.get(f"{{{W}}}date", ""),
                "paragraphs": paragraphs,
                "done": None,
                "parent": None,
            }
            for paragraph in paragraphs:
                para_id = paragraph.get(f"{{{KNOWN_NAMESPACES['w14']}}}paraId")
                if para_id:
                    para_owner[para_id] = cid
            if paragraphs:
                last = paragraphs[-1].get(f"{{{KNOWN_NAMESPACES['w14']}}}paraId")
                if last in extended:
                    info["done"] = extended[last]["done"]
                    info["parent_para"] = extended[last]["parent"]
            result[cid] = info
        for info in result.values():
            parent = info.pop("parent_para", None)
            if parent:
                info["parent"] = para_owner.get(parent)
        return result

    def _comment_id(self, raw: str) -> str:
        # Keeping Word's id lets a redline against the source match comments.
        if raw not in self.comment_ids:
            self.comment_ids[raw] = f"c{raw}" if raw.isdigit() else f"c{len(self.comment_ids) + 1}"
        return self.comment_ids[raw]

    # -- raw fallback --------------------------------------------------------
    def _raw(self, node: etree._Element, level: str) -> Atom:
        clone = copy.deepcopy(node)
        for element in clone.iter():
            for key in list(element.attrib):
                local = etree.QName(key).localname
                if local.startswith("rsid") or key in {f"{{{KNOWN_NAMESPACES['w14']}}}paraId", f"{{{KNOWN_NAMESPACES['w14']}}}textId"}:
                    del element.attrib[key]
        rels = []
        for element in clone.iter():
            for key, value in element.attrib.items():
                if etree.QName(key).namespace == R and value in self.rels:
                    rel_type, target, mode = self.rels[value]
                    short = rel_type.rsplit("/", 1)[-1]
                    if short == "image":
                        target = self._media_path(value)
                    rels.append(f"{value} {short} {target}{' external' if mode == 'External' else ''}")
        etree.cleanup_namespaces(clone)
        body = etree.tostring(clone, encoding=str)
        attrs = [("level", level)]
        if rels:
            attrs.append(("rels", ";".join(dict.fromkeys(rels))))
        self.report["raw"].append(_local(node))
        return Atom("ooxml", tuple(attrs), (), body)

    # -- paragraph analysis ----------------------------------------------------
    def _rpr(self, run: etree._Element) -> list[Token]:
        return clean_tokens(flatten(run.find(f"./{{{W}}}rPr")))

    def _pieces(self, container: etree._Element, wrappers: tuple, out: list[Piece]) -> None:
        for child in container:
            name = _local(child)
            if name in {"pPr", ""}:
                continue
            if name == "r":
                self._run_pieces(child, wrappers, out)
            elif name in {"ins", "del", "moveTo", "moveFrom"}:
                attrs = tuple(
                    (key, child.get(f"{{{W}}}{key}")) for key in ("author", "date") if child.get(f"{{{W}}}{key}")
                )
                kind = "ins" if name in {"ins", "moveTo"} else "del"
                self._pieces(child, wrappers + (("rev", kind, attrs),), out)
            elif name == "hyperlink":
                rel = child.get(f"{{{R}}}id")
                if rel and rel in self.rels:
                    href = self.rels[rel][1]
                elif child.get(f"{{{W}}}anchor"):
                    href = "#" + child.get(f"{{{W}}}anchor")
                else:
                    href = ""
                self._pieces(child, wrappers + (("link", href),), out)
            elif name in {"bookmarkStart", "bookmarkEnd"} and child.get(f"{{{W}}}id") in self.bookmarks:
                kind = "bookmark-start" if name == "bookmarkStart" else "bookmark-end"
                out.append(Piece("milestone", Atom(kind, (("name", self.bookmarks[child.get(f"{{{W}}}id")]),)), [], ()))
            elif name in {"bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd"}:
                self.report["dropped"][name] += 1
            elif name in {"commentRangeStart", "commentRangeEnd"}:
                cid = self._comment_id(child.get(f"{{{W}}}id"))
                kind = "comment-start" if name == "commentRangeStart" else "comment-end"
                out.append(Piece("milestone", Atom(kind, (("id", cid),)), [], ()))
            elif name in {"smartTag", "customXml"}:
                self._pieces(child, wrappers, out)
            else:
                out.append(Piece("milestone", self._raw(child, "node"), [], ()))

    def _run_pieces(self, run: etree._Element, wrappers: tuple, out: list[Piece]) -> None:
        rpr = self._rpr(run)
        for child in run:
            name = _local(child)
            if name in {"rPr", "", "lastRenderedPageBreak"}:
                continue
            if name in {"t", "delText"}:
                if child.text:
                    out.append(Piece("text", child.text, rpr, wrappers))
                continue
            atom: Atom | None = None
            if name == "tab":
                atom = Atom("tab")
            elif name == "br":
                attrs = tuple((key, child.get(f"{{{W}}}{key}")) for key in ("type", "clear") if child.get(f"{{{W}}}{key}"))
                atom = Atom("br", attrs)
            elif name == "cr":
                atom = Atom("cr")
            elif name == "sym":
                atom = Atom("sym", (("font", child.get(f"{{{W}}}font", "")), ("char", child.get(f"{{{W}}}char", ""))))
            elif name == "noBreakHyphen":
                atom = Atom("nbhyphen")
            elif name == "softHyphen":
                atom = Atom("shy")
            elif name == "fldChar" and len(child) == 0:
                kind = {"begin": "fld-begin", "separate": "fld-sep", "end": "fld-end"}[child.get(f"{{{W}}}fldCharType")]
                attrs = tuple((key, child.get(f"{{{W}}}{key}")) for key in ("dirty", "fldLock") if child.get(f"{{{W}}}{key}"))
                atom = Atom(kind, attrs)
            elif name in {"instrText", "delInstrText"}:
                atom = Atom("instr", (("code", child.text or ""),))
            elif name == "annotationRef":
                atom = Atom("annotation-ref")
            elif name == "commentReference":
                raw = child.get(f"{{{W}}}id")
                if raw in self.range_ends:
                    continue
                out.append(Piece("milestone", Atom("comment-ref", (("id", self._comment_id(raw)),)), [], ()))
                continue
            elif name == "drawing":
                described = drawing_mod.describe(child)
                if described is not None:
                    attrs, rel_id = described
                    attrs = [("alt", attrs[0][1]), ("src", self._media_path(rel_id))] + attrs[1:]
                    atom = Atom("image", tuple(attrs))
            if atom is None:
                atom = self._raw(child, "run")
            out.append(Piece("atom", atom, rpr, wrappers))

    def _paragraph(self, element: etree._Element, table_style: str | None = None) -> Para:
        ppr = clean_tokens(flatten(element.find(f"./{{{W}}}pPr")))
        style = next((value for key, value in ppr if key == "pStyle@val"), None)
        mark = [(key[4:], value) for key, value in ppr if key.startswith("rPr.")]
        mark = self._prune(style, table_style, [], mark, "a中")
        ppr = [(key, value) for key, value in ppr if not key.startswith("rPr.")] + [("rPr." + key, value) for key, value in mark]
        pieces: list[Piece] = []
        self._pieces(element, (), pieces)
        weights: Counter = Counter()
        first: dict[tuple, int] = {}
        for index, piece in enumerate(pieces):
            if piece.kind == "text":
                key = tuple(piece.rpr)
                weights[key] += len(piece.value)
                first.setdefault(key, index)
        base = None
        if weights:
            best = max(weights, key=lambda key: (weights[key], -first[key]))
            base = list(best)
        signature = tuple(
            item for item in ppr
            if top_name(item[0]) not in {"pPrChange", "sectPr", "rPr"}
        )
        level = None
        for key, value in ppr:
            if key == "outlineLvl@val":
                level = int(value)
        if level is None:
            level = self.styles.outline_level(style)
        if level is not None and level >= 9:
            level = None
        return Para(element, ppr, style, table_style, pieces, base, signature, level=level)

    # -- visibility --------------------------------------------------------------
    def _visible(self, style, table_style, tokens: list[Token], sample: str):
        key = (style, table_style, tuple(tokens), sample)
        if key not in self.visible_cache:
            self.visible_cache[key] = self.styles.visible(style, tokens, sample, table_style)
        return self.visible_cache[key]

    def _prune(self, style, table_style, base: list[Token], target: list[Token], text: str) -> list[Token]:
        """Smallest override of ``base`` that renders like ``target`` for ``text``."""
        sample = _sample_text(text) or "a"
        goal = self._visible(style, table_style, target, sample)
        delta = diff(base, target)
        groups: list[list[Token]] = []
        index: dict[str, int] = {}
        for token in delta:
            name = top_name(token[0])
            group = token[0].lstrip("-") if name == "rFonts" and "@" in token[0] else name
            if group not in index:
                index[group] = len(groups)
                groups.append([])
            groups[index[group]].append(token)
        kept = list(delta)
        for group in groups:
            trial = [token for token in kept if token not in group]
            if self._visible(style, table_style, merge(base, trial), sample) == goal:
                kept = trial
        return kept

    def _marks(self, para: Para, base: list[Token], rpr: list[Token], text: str) -> tuple:
        sample = _sample_text(text)
        key = (para.style, para.table_style, tuple(base), tuple(rpr), sample)
        cached = self.mark_cache.get(key)
        if cached is not None:
            return cached
        sample = sample or "a"
        goal = self._visible(para.style, para.table_style, rpr, sample)
        delta = self._prune(para.style, para.table_style, base, rpr, sample)
        shortcuts: list[tuple] = []
        merged = dict(merge(base, delta))
        candidates = []
        for mark, element in SHORTCUT_ELEMENT.items():
            if not any(top_name(token[0]) == element for token in delta):
                continue
            if mark in {"b", "i", "strike"}:
                active = (element in merged or f"{element}@val" in merged) and (merged.get(f"{element}@val") or "1") not in {"0", "false", "off"}
            elif mark in {"sup", "sub"}:
                active = merged.get("vertAlign@val") == {"sup": "superscript", "sub": "subscript"}[mark]
            elif mark == "u":
                active = merged.get("u@val") == "single" and not any(key.startswith("u@") and key != "u@val" for key in merged)
            else:
                active = merged.get("highlight@val") == "yellow"
            if active:
                candidates.append((mark, element))
        for mark, element in candidates:
            trial = [token for token in delta if top_name(token[0]) != element]
            marks = tuple(shortcuts) + ((mark,),)
            span = (("span", format_tokens(trial)),) if trial else ()
            if self._visible(para.style, para.table_style, run_tokens_for(base, span + marks), sample) == goal:
                delta = trial
                shortcuts.append((mark,))
        result = tuple(shortcuts)
        if delta:
            result = (("span", format_tokens(delta)),) + result
        self.mark_cache[key] = result
        return result

    # -- class registry --------------------------------------------------------
    def assign_classes(self, paragraphs: list[Para]) -> dict[str, dict]:
        groups: dict[tuple, list[Para]] = {}
        order: list[tuple] = []
        for para in paragraphs:
            key = (para.style, para.signature)
            if key not in groups:
                groups[key] = []
                order.append(key)
            groups[key].append(para)
        by_style: dict[str | None, list[tuple]] = {}
        for key in order:
            by_style.setdefault(key[0], []).append(key)
        classes: dict[str, dict] = {}
        used: set[str] = set()
        for style, keys in by_style.items():
            keys = sorted(keys, key=lambda item: (-len(groups[item]), order.index(item)))
            label = self._class_label(style)
            for rank, key in enumerate(keys):
                name = label if rank == 0 else f"{label}-{rank + 1}"
                while name in used:
                    name += "x"
                used.add(name)
                members = groups[key]
                marks = Counter(tuple(t for t in para.ppr if t[0].startswith("rPr.")) for para in members)
                mark = list(max(marks, key=lambda item: (marks[item], -[tuple(t for t in p.ppr if t[0].startswith("rPr.")) for p in members].index(item))))
                bases = Counter(tuple(para.base) for para in members if para.base is not None)
                base = list(max(bases, key=lambda item: bases[item])) if bases else []
                classes[name] = {"p": list(key[1]) + mark, "r": base}
                for para in members:
                    para.cls = name
                    para.override = diff(classes[name]["p"], para.ppr)
                    if para.base is not None:
                        para.r_override = self._prune(para.style, para.table_style, base, para.base, para.text)
        return classes

    def _class_label(self, style: str | None) -> str:
        if style is None:
            return "Normal"
        name = style if re.fullmatch(r"[A-Za-z][\w-]*", style) else self.styles.names.get(style, style)
        name = re.sub(r"[^\w-]+", "-", name).strip("-") or "Style"
        if not re.match(r"[A-Za-z]", name):
            name = "S" + name
        return name

    # -- rendering -------------------------------------------------------------
    def items(self, para: Para, base: list[Token], citations: bool) -> list:
        items: list = []
        for piece in para.pieces:
            if piece.kind == "milestone":
                items.append(piece.value)
                continue
            text = piece.value if piece.kind == "text" else ""
            marks = piece.wrappers + self._marks(para, base, piece.rpr, text)
            if piece.kind == "text":
                items.append(Text(piece.value, marks))
            else:
                items.append(Atom(piece.value.kind, piece.value.attrs, marks, piece.value.body))
        items = normalize(items)
        if citations and self.citation_keys:
            items = self._citations(items)
        return items

    def _citations(self, items: list) -> list:
        result: list = []
        previous = ""
        for item in items:
            if isinstance(item, Text) and ("sup",) in item.marks:
                numbers = parse_citation_numbers(item.text)
                if (
                    numbers
                    and previous[-1:] in _CITATION_PRECEDERS
                    and all(number in self.citation_keys for number in numbers)
                    and format_citation_numbers(numbers) == item.text
                ):
                    marks = tuple(mark for mark in item.marks if mark != ("sup",))
                    keys = ",".join(self.citation_keys[number] for number in numbers)
                    result.append(Atom("cite", (("keys", keys),), marks))
                    previous += item.text
                    continue
            if isinstance(item, Text):
                previous += item.text
            result.append(item)
        return result

    def attrs_for(self, para: Para, default: str | None, role: str | None = None) -> Attrs:
        attrs = Attrs()
        if para.cls != default:
            attrs.cls = para.cls
        attrs.p = list(para.override)
        attrs.r = list(para.r_override)
        if role:
            attrs.params["role"] = role
        return attrs

    def role(self, para: Para) -> str | None:
        if para.style is None:
            return None
        if para.text.strip().lower().startswith("keywords"):
            return "keywords"
        name = re.sub(r"[\s_-]+", "", (self.styles.names.get(para.style, para.style) + " " + para.style).lower())
        for keyword, role in _ROLE_KEYWORDS:
            if keyword in name:
                return role
        return None


# ----------------------------------------------------------------------------


def _paragraph_markdown(reader: _Reader, para: Para, classes: dict, default: str | None, headings: dict[int, str]) -> str:
    base = merge(classes[para.cls]["r"], para.r_override)
    items = reader.items(para, base, reader.options.citations)
    role = reader.role(para)
    if para.level is not None:
        heading_default = headings.get(para.level)
        attrs = reader.attrs_for(para, heading_default, role)
        line = "#" * (para.level + 1) + " " + write_inline(items, block_start=False)
        return line + ("\n" + format_attrs(attrs) if not attrs.empty() else "")
    attrs = reader.attrs_for(para, default, role)
    line = write_inline(items) if items else ""
    if not line:
        return format_attrs(attrs)
    return line + ("\n" + format_attrs(attrs) if not attrs.empty() else "")


def _cell_markdown(reader: _Reader, cell_paras: list[Para], classes: dict, default: str | None, tc: list[Token], tr: list[Token]) -> str:
    chunks = []
    for index, para in enumerate(cell_paras):
        base = merge(classes[para.cls]["r"], para.r_override)
        items = reader.items(para, base, False)
        text = write_inline(items, block_start=False) if items else ""
        attrs = reader.attrs_for(para, default)
        if index == 0:
            attrs.tc = tc
            attrs.tr = tr
        suffix = format_attrs(attrs) if not attrs.empty() else ""
        chunks.append((text + (" " if text and suffix else "") + suffix).strip() or "{:}")
    return " <p/> ".join(chunks)


def _block_text(block: etree._Element) -> str:
    return "".join(node.text or "" for node in block.iter(f"{{{W}}}t"))


def _surname_key(text: str, used: set[str]) -> str:
    plain = re.sub(r"^\[\d+\]\s*", "", text).strip()
    first = re.split(r"[;,]", plain, maxsplit=1)[0].strip()
    words = [word for word in re.split(r"\s+", first) if word]
    initials = re.compile(r"^(?:[A-Z]\.(?:-?[A-Z]\.)*|[A-Z]\.?)$")
    surname_words = [word for word in words if not initials.match(word)]
    surname = surname_words[-1] if surname_words and initials.match(words[0]) else (surname_words[0] if surname_words else "ref")
    surname = unicodedata.normalize("NFKD", surname).encode("ascii", "ignore").decode()
    surname = re.sub(r"[^a-z]", "", surname.lower()) or "ref"
    year = re.search(r"\b(19|20)\d{2}\b", plain)
    key = surname + (year.group(0) if year else "")
    candidate = key
    suffix = ord("a")
    while candidate in used:
        candidate = key + chr(suffix)
        suffix += 1
    used.add(candidate)
    return candidate


def export_docx(source: Path, bundle: Path, options: ExportOptions) -> dict:
    """Write ``source`` into the Markdown ``bundle`` as document ``options.doc_id``."""
    bundle.mkdir(parents=True, exist_ok=True)
    media_dir = bundle / "media" / options.doc_id
    if media_dir.exists():
        shutil.rmtree(media_dir)
    reader = _Reader(source, bundle, options)
    blocks = [child for child in reader.body if _local(child) != "sectPr"]
    paragraphs: list[Para] = []
    block_models: list[tuple[str, object, etree._Element]] = []
    for block in blocks:
        name = _local(block)
        if name == "p":
            para = reader._paragraph(block)
            paragraphs.append(para)
            block_models.append(("p", para, block))
        elif name == "tbl":
            model = _table_model(reader, block)
            if model is None:
                block_models.append(("raw", reader._raw(block, "block"), block))
            else:
                for row in model["rows"]:
                    for cell in row["cells"]:
                        paragraphs.extend(cell["paras"])
                block_models.append(("tbl", model, block))
        elif name in {"bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd"}:
            reader.report["dropped"][name] += 1
        else:
            block_models.append(("raw", reader._raw(block, "block"), block))
    classes = reader.assign_classes(paragraphs)

    # References: the paragraphs after a "References" heading.
    bib_style = options.bib_style or options.doc_id
    bib_path = bundle / options.bibliography
    bibliography = _load_bibliography(bib_path, bib_style)
    used_keys = {key for key in bibliography if not key.startswith("_")}
    references = _find_references(block_models)
    reference_block = None
    if references is not None:
        start, end, mode = references
        entries = [block_models[index][1] for index in range(start, end)]
        reference_block = (start, end, mode, entries)
        for number, para in enumerate(entries, 1):
            key = _surname_key(para.text, used_keys)
            reader.citation_keys[number] = key

    # Paragraph classes used as defaults.
    split = options.split or [{"file": f"{options.doc_id}.md", "start": None}]
    part_of_block: list[int] = []
    current = 0
    for index, (kind, model, element) in enumerate(block_models):
        if current + 1 < len(split):
            pattern = split[current + 1].get("start")
            if pattern and re.search(pattern, _block_text(element)):
                current += 1
        part_of_block.append(current)
    in_references = set(range(reference_block[0], reference_block[1])) if reference_block else set()
    part_defaults: dict[int, str | None] = {}
    for part in range(len(split)):
        counts = Counter(
            model.cls for index, ((kind, model, _), owner) in enumerate(zip(block_models, part_of_block))
            if owner == part and kind == "p" and model.level is None and model.text.strip()
            and index not in in_references
        )
        part_defaults[part] = max(counts, key=lambda name: counts[name]) if counts else None
    headings: dict[int, str] = {}
    heading_counts: dict[int, Counter] = {}
    for para in paragraphs:
        if para.level is not None:
            heading_counts.setdefault(para.level, Counter())[para.cls] += 1
    for level, counts in heading_counts.items():
        headings[level] = max(counts, key=lambda name: counts[name])

    # Bibliography entries.
    if reference_block is not None:
        start, end, mode, entries = reference_block
        ref_counts = Counter(para.cls for para in entries)
        ref_class = max(ref_counts, key=lambda name: ref_counts[name])
        bases = Counter(tuple(merge(classes[para.cls]["r"], para.r_override)) for para in entries)
        ref_base = list(max(bases, key=lambda item: bases[item]))
        directive = Attrs(cls=ref_class)
        directive.r = diff(classes[ref_class]["r"], ref_base)
        directive.params["style"] = bib_style
        rendered_entries = []
        label_ok = mode == "bracket"
        for number, para in enumerate(entries, 1):
            items = reader.items(para, ref_base, False)
            if mode == "bracket":
                if not (
                    len(items) >= 2 and isinstance(items[0], Text) and items[0].text == f"[{number}]"
                    and not items[0].marks and isinstance(items[1], Atom) and items[1].kind == "tab"
                    and not items[1].marks
                ):
                    label_ok = False
            rendered_entries.append((para, items))
        if label_ok:
            directive.params["label"] = "\\[{n}\\]<tab/>"
        for number, (para, items) in enumerate(rendered_entries, 1):
            if label_ok:
                items = items[2:]
            key = reader.citation_keys[number]
            entry = bibliography.setdefault(key, {})
            entry.setdefault("rendered", {})[bib_style] = write_inline(items, block_start=False)
            layout = Attrs()
            if para.cls != ref_class:
                layout.cls = para.cls
            layout.p = list(para.override)
            layout.r = diff(ref_base, merge(classes[para.cls]["r"], para.r_override))
            if not layout.empty():
                entry.setdefault("layout", {})[bib_style] = format_attrs(layout)
        reference_directive = "::: references " + format_attrs(directive) + "\n:::"
    _save_bibliography(bib_path, bibliography)

    # Parts.
    texts: list[list[str]] = [[] for _ in split]
    comment_part: dict[str, int] = {}
    metadata: dict[str, object] = {"title": None, "authors": [], "affiliations": [], "keywords": None}
    index = 0
    while index < len(block_models):
        kind, model, element = block_models[index]
        part = part_of_block[index]
        if reference_block is not None and index == reference_block[0]:
            texts[part].append(reference_directive)
            index = reference_block[1]
            continue
        if kind == "p":
            texts[part].append(_paragraph_markdown(reader, model, classes, part_defaults[part], headings))
            role = reader.role(model)
            if role == "title" and metadata["title"] is None:
                metadata["title"] = model.text.strip()
            elif role == "authors":
                names = (re.sub(r"[†‡*#§¶\d,]+$", "", name.strip()).strip()
                         for name in re.split(r",\s*|\s+and\s+", model.text))
                metadata["authors"].extend(name for name in names if name)
            elif role == "affiliation" and model.text.strip():
                metadata["affiliations"].append(model.text.strip())
            elif role == "keywords":
                metadata["keywords"] = model.text.split(":", 1)[-1].strip()
        elif kind == "tbl":
            texts[part].append(_table_markdown(reader, model, classes))
        else:
            head = Attrs()
            if model.get("rels"):
                head.params["rels"] = model.get("rels")
            fence = "::: ooxml" + (" " + format_attrs(head) if not head.empty() else "")
            texts[part].append(fence + "\n" + model.body + "\n:::")
        for atom_id in re.findall(r'<comment-(?:start|ref) id="(c\d+)"/>', texts[part][-1]):
            comment_part.setdefault(atom_id, part)
        index += 1

    # Comment bodies at the end of the part that anchors them.
    raw_by_md = {value: key for key, value in reader.comment_ids.items()}
    for md_id, part in sorted(comment_part.items(), key=lambda item: int(item[0][1:])):
        info = reader.comments.get(raw_by_md[md_id])
        if info is None:
            continue
        texts[part].append(_comment_markdown(reader, md_id, info))

    header = _document_header(reader, classes, headings)
    files = []
    for number, part in enumerate(split):
        lines = []
        if number == 0:
            lines.append(header)
        default = part_defaults[number]
        lines.append(f"<!-- docforge:part default={default} -->" if default else "<!-- docforge:part -->")
        content = "\n\n".join(lines + texts[number]) + "\n"
        (bundle / part["file"]).write_text(content, encoding="utf-8", newline="\n")
        files.append(part["file"])

    templates = bundle / "_docx"
    templates.mkdir(exist_ok=True)
    template = templates / f"{options.doc_id}.docx"
    if Path(source).resolve() != template.resolve():
        shutil.copyfile(source, template)
    manifest_path = bundle / "bundle.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {"schema": SCHEMA, "documents": []}
    documents = [doc for doc in manifest["documents"] if doc["id"] != options.doc_id]
    entry = {
        "id": options.doc_id,
        "template": f"_docx/{options.doc_id}.docx",
        "parts": files,
        "bibliography": {"path": options.bibliography, "style": bib_style} if reference_block else None,
        "metadata": {key: value for key, value in metadata.items() if value},
        "source": Path(source).name,
        "source_sha256": hashlib.sha256(Path(source).read_bytes()).hexdigest(),
    }
    documents.append(entry)
    order = {doc["id"]: index for index, doc in enumerate(manifest["documents"])}
    documents.sort(key=lambda doc: order.get(doc["id"], len(order)))
    manifest["documents"] = documents
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    report = {
        "document": options.doc_id,
        "parts": files,
        "paragraphs": len(paragraphs),
        "classes": len(classes),
        "references": (reference_block[1] - reference_block[0]) if reference_block else 0,
        "comments": len(comment_part),
        "raw": Counter(reader.report["raw"]),
        "dropped": dict(reader.report["dropped"]),
    }
    return report


def _load_bibliography(path: Path, style: str) -> dict:
    if path.exists():
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    else:
        payload = {"_schema": BIB_SCHEMA}
    for key in [key for key in payload if not key.startswith("_")]:
        entry = payload[key]
        if isinstance(entry, dict):
            entry.get("rendered", {}).pop(style, None)
            entry.get("layout", {}).pop(style, None)
            if "layout" in entry and not entry["layout"]:
                del entry["layout"]
            if "rendered" in entry and not entry["rendered"] and len(entry) == 1:
                del payload[key]
    payload.setdefault("_schema", BIB_SCHEMA)
    payload.setdefault(
        "_comment",
        "Entries render verbatim from 'rendered[style]' (inline docforge Markdown). "
        "Within one style, entry order is reference order.",
    )
    return payload


def _save_bibliography(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _find_references(block_models) -> tuple[int, int, str] | None:
    for index, (kind, model, element) in enumerate(block_models):
        if kind != "p" or not _REFERENCE_HEADING.match(model.text):
            continue
        start = index + 1
        numbered = []
        for later in range(start, len(block_models)):
            kind2, para, _ = block_models[later]
            if kind2 != "p":
                break
            num = next((value for key, value in para.ppr if key == "numPr.numId@val"), None)
            if num is None:
                break
            numbered.append(num)
        if numbered and len(set(numbered)) == 1:
            return start, start + len(numbered), "numbering"
        count = 0
        for later in range(start, len(block_models)):
            kind2, para, _ = block_models[later]
            if kind2 != "p" or not para.text.startswith(f"[{count + 1}]"):
                break
            count += 1
        if count:
            return start, start + count, "bracket"
    return None


def _table_model(reader: _Reader, table: etree._Element) -> dict | None:
    tbl_pr = clean_tokens(flatten(table.find(f"./{{{W}}}tblPr")))
    style = next((value for key, value in tbl_pr if key == "tblStyle@val"), None)
    grid = [col.get(f"{{{W}}}w", "0") for col in table.findall(f"./{{{W}}}tblGrid/{{{W}}}gridCol")]
    rows = []
    for child in table:
        name = _local(child)
        if name in {"tblPr", "tblGrid", "bookmarkStart", "bookmarkEnd"}:
            continue
        if name != "tr":
            return None
        cells = []
        tr = clean_tokens(flatten(child.find(f"./{{{W}}}trPr")))
        for cell in child:
            cname = _local(cell)
            if cname in {"trPr", "bookmarkStart", "bookmarkEnd"}:
                continue
            if cname != "tc":
                return None
            tc = clean_tokens(flatten(cell.find(f"./{{{W}}}tcPr")))
            paras = []
            for item in cell:
                iname = _local(item)
                if iname == "tcPr" or iname in {"bookmarkStart", "bookmarkEnd"}:
                    continue
                if iname != "p":
                    return None
                paras.append(reader._paragraph(item, style))
            cells.append({"tc": tc, "paras": paras})
        rows.append({"tr": tr, "cells": cells})
    return {"tblPr": tbl_pr, "grid": grid, "rows": rows}


def _table_markdown(reader: _Reader, model: dict, classes: dict) -> str:
    counts = Counter(para.cls for row in model["rows"] for cell in row["cells"] for para in cell["paras"])
    default = max(counts, key=lambda name: counts[name]) if counts else None
    head = Attrs(p=model["tblPr"])
    head.params["grid"] = ",".join(model["grid"])
    if default:
        head.params["default"] = default
    lines = ["::: table " + format_attrs(head)]
    for number, row in enumerate(model["rows"]):
        cells = [
            _cell_markdown(reader, cell["paras"], classes, default, cell["tc"], row["tr"] if index == 0 else [])
            for index, cell in enumerate(row["cells"])
        ]
        lines.append("| " + " | ".join(cells) + " |")
        if number == 0:
            lines.append("|" + "---|" * len(cells))
    lines.append(":::")
    return "\n".join(lines)


def _comment_markdown(reader: _Reader, md_id: str, info: dict) -> str:
    head = Attrs()
    head.params["id"] = md_id
    for key in ("author", "initials", "date"):
        if info.get(key):
            head.params[key] = info[key]
    if info.get("done") not in (None, "0"):
        head.params["done"] = info["done"]
    if info.get("parent"):
        head.params["parent"] = reader.comment_ids.get(info["parent"], info["parent"])
    lines = ["::: comment " + format_attrs(head)]
    for element in info["paragraphs"]:
        para = reader._paragraph(element)
        base = para.base or []
        para.cls = None
        para.override = list(para.ppr)
        para.r_override = reader._prune(para.style, None, [], base, para.text)
        base = merge([], para.r_override)
        items = reader.items(para, base, False)
        attrs = Attrs(p=list(para.ppr), r=list(para.r_override))
        line = write_inline(items) if items else ""
        lines.append((line + "\n" if line else "") + format_attrs(attrs))
        lines.append("")
    if lines[-1] == "":
        lines.pop()
    lines.append(":::")
    return "\n".join(lines)


def _document_header(reader: _Reader, classes: dict, headings: dict[int, str]) -> str:
    lines = [f"<!-- docforge:document id={reader.options.doc_id}"]
    sect = reader.body.find(f"./{{{W}}}sectPr")
    if sect is not None:
        tokens = clean_tokens(flatten(sect))
        attrs = [(f"@{key.split('}')[-1]}", value) for key, value in sect.attrib.items() if not etree.QName(key).localname.startswith("rsid")]
        lines.append("section: " + format_tokens(tokens + attrs))
    background = reader.root.find(f"./{{{W}}}background")
    if background is not None:
        tokens = [(f"@{etree.QName(key).localname}", value) for key, value in background.attrib.items()]
        lines.append("background: " + format_tokens(tokens + flatten(background)))
    for level, name in sorted(headings.items()):
        lines.append(f"heading{level + 1}: {name}")
    for name, spec in classes.items():
        lines.append(f"class {name} | p: {format_tokens(spec['p'])} | r: {format_tokens(spec['r'])}")
    lines.append("-->")
    return "\n".join(lines)


__all__ = ["ExportOptions", "export_docx"]
