"""Markdown bundle -> DOCX, on top of the document's source package."""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath

from lxml import etree

from ..docxdiff.package import Package
from . import drawing as drawing_mod
from .common import ATTR_LINE, Attrs, format_citation_numbers, parse_attrs, run_tokens_for
from .inline import MILESTONES, Atom, Text, parse as parse_inline
from .ooxml import (
    CONTENT_TYPES, KNOWN_NAMESPACES, PKG_REL, R, W, XML_NS, Token, build, merge, parse_tokens,
    qname,
)
from .styles import StyleSheet

W14 = KNOWN_NAMESPACES["w14"]
W15 = KNOWN_NAMESPACES["w15"]
WP = KNOWN_NAMESPACES["wp"]
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_TYPES = {
    "image": f"{REL_BASE}/image",
    "hyperlink": f"{REL_BASE}/hyperlink",
    "comments": f"{REL_BASE}/comments",
    "commentsExtended": "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
}
MEDIA_TYPES = {
    "png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg", "gif": "image/gif",
    "tif": "image/tiff", "tiff": "image/tiff", "bmp": "image/bmp", "emf": "image/x-emf",
    "wmf": "image/x-wmf", "svg": "image/svg+xml",
}
_REVISION_TAGS = {
    f"{{{W}}}{name}" for name in (
        "ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange", "sectPrChange",
        "tblPrChange", "trPrChange", "tcPrChange", "numberingChange", "cellIns", "cellDel",
        "cellMerge", "tblGridChange",
    )
}


# ----------------------------------------------------------------------------
# Markdown parsing


@dataclass
class Block:
    type: str
    text: str = ""
    attrs: Attrs = field(default_factory=Attrs)
    level: int | None = None
    rows: list = field(default_factory=list)
    paras: list = field(default_factory=list)
    body: str = ""


@dataclass
class Part:
    document: dict | None
    default: str | None
    blocks: list[Block]


def _split_unescaped(text: str, separator: str) -> list[str]:
    pieces: list[str] = []
    current: list[str] = []
    index = 0
    while index < len(text):
        if text.startswith("<ooxml", index):
            end = text.find("</ooxml>", index)
            if end >= 0:
                current.append(text[index:end + 8])
                index = end + 8
                continue
        char = text[index]
        if char == "\\" and index + 1 < len(text):
            current.append(text[index:index + 2])
            index += 2
            continue
        if text.startswith(separator, index):
            pieces.append("".join(current))
            current = []
            index += len(separator)
            continue
        current.append(char)
        index += 1
    pieces.append("".join(current))
    return pieces


def _trailing_attrs(text: str) -> tuple[str, Attrs]:
    text = text.strip()
    index = len(text)
    position = -1
    scan = 0
    while scan < len(text):
        if text[scan] == "\\":
            scan += 2
            continue
        if text.startswith("<ooxml", scan):
            end = text.find("</ooxml>", scan)
            scan = end + 8 if end >= 0 else len(text)
            continue
        if text.startswith("{:", scan):
            position = scan
        scan += 1
    if position >= 0 and text.endswith("}") and ATTR_LINE.match(text[position:index]):
        return text[:position].strip(), parse_attrs(text[position:])
    return text, Attrs()


def _document_header(body: str) -> dict:
    lines = body.strip().splitlines()
    header = {"id": None, "section": [], "background": None, "headings": {}, "classes": {}}
    match = re.search(r"id=(\S+)", lines[0])
    if match:
        header["id"] = match.group(1)
    for line in lines[1:]:
        line = line.strip()
        if not line:
            continue
        if line.startswith("section:"):
            header["section"] = parse_tokens(line[len("section:"):])
        elif line.startswith("background:"):
            header["background"] = parse_tokens(line[len("background:"):])
        elif re.match(r"heading\d+:", line):
            level, name = line.split(":", 1)
            header["headings"][int(level[7:]) - 1] = name.strip()
        elif line.startswith("class "):
            name, rest = line[6:].split("|", 1)
            spec = {"p": [], "r": []}
            for chunk in rest.split(" | "):
                chunk = chunk.strip().lstrip("|").strip()
                if chunk.startswith("p:"):
                    spec["p"] = parse_tokens(chunk[2:])
                elif chunk.startswith("r:"):
                    spec["r"] = parse_tokens(chunk[2:])
            header["classes"][name.strip()] = spec
    return header


def parse_part(text: str) -> Part:
    lines = text.splitlines()
    blocks: list[Block] = []
    document = None
    default = None
    index = 0
    while index < len(lines):
        line = lines[index].rstrip()
        if not line.strip():
            index += 1
            continue
        if line.lstrip().startswith("<!--"):
            chunk = [line]
            while "-->" not in chunk[-1] and index + 1 < len(lines):
                index += 1
                chunk.append(lines[index])
            index += 1
            body = "\n".join(chunk).strip()[4:-3].strip()
            if body.startswith("docforge:document"):
                document = _document_header(body)
            elif body.startswith("docforge:part"):
                match = re.search(r"default=(\S+)", body)
                default = match.group(1) if match else None
            continue
        if line.startswith(":::"):
            head = line[3:].strip()
            kind, _, rest = head.partition(" ")
            body_lines = []
            index += 1
            while index < len(lines) and lines[index].strip() != ":::":
                body_lines.append(lines[index])
                index += 1
            index += 1
            attrs = parse_attrs(rest) if rest.strip() else Attrs()
            if kind == "table":
                rows = []
                for row in body_lines:
                    row = row.strip()
                    if not row or re.fullmatch(r"\|(?:\s*:?-+:?\s*\|)+", row):
                        continue
                    cells = _split_unescaped(row, "|")[1:-1]
                    parsed = []
                    for cell in cells:
                        paragraphs = [_trailing_attrs(item) for item in _split_unescaped(cell, "<p/>")]
                        parsed.append(paragraphs)
                    rows.append(parsed)
                blocks.append(Block("table", attrs=attrs, rows=rows))
            elif kind == "comment":
                blocks.append(Block("comment", attrs=attrs, paras=_paragraph_chunks(body_lines)))
            elif kind == "references":
                blocks.append(Block("references", attrs=attrs))
            elif kind == "ooxml":
                blocks.append(Block("ooxml", attrs=attrs, body="\n".join(body_lines).strip()))
            else:
                raise ValueError(f"Unknown block ::: {kind}")
            continue
        heading = re.match(r"^(#{1,9}) (.*)$", line)
        if heading:
            attrs = Attrs()
            if index + 1 < len(lines) and ATTR_LINE.match(lines[index + 1].strip()):
                attrs = parse_attrs(lines[index + 1])
                index += 1
            blocks.append(Block("p", heading.group(2).strip(), attrs, level=len(heading.group(1)) - 1))
            index += 1
            continue
        chunk = []
        while index < len(lines) and lines[index].strip():
            if chunk and (lines[index].startswith(":::") or lines[index].lstrip().startswith("<!--")):
                break
            chunk.append(lines[index].strip())
            index += 1
        attrs = Attrs()
        if chunk and ATTR_LINE.match(chunk[-1]):
            attrs = parse_attrs(chunk.pop())
        blocks.append(Block("p", " ".join(chunk), attrs))
    return Part(document, default, blocks)


def _paragraph_chunks(lines: list[str]) -> list[tuple[str, Attrs]]:
    result = []
    chunk: list[str] = []
    for line in lines + [""]:
        if line.strip():
            chunk.append(line.strip())
            continue
        if chunk:
            attrs = Attrs()
            if ATTR_LINE.match(chunk[-1]):
                attrs = parse_attrs(chunk.pop())
            result.append((" ".join(chunk), attrs))
            chunk = []
    return result


# ----------------------------------------------------------------------------
# DOCX assembly


class _Writer:
    def __init__(self, bundle: Path, document: dict) -> None:
        self.bundle = bundle
        self.document = document
        self.package = Package.load(bundle / document["template"])
        self.root = self.package.xml("word/document.xml")
        self.nsmap = dict(self.root.nsmap)
        for prefix, uri in KNOWN_NAMESPACES.items():
            if prefix != "xml":
                self.nsmap.setdefault(prefix, uri)
        styles = self.package.parts.get("word/styles.xml")
        self.styles = StyleSheet(etree.fromstring(styles) if styles else None)
        texts = [(bundle / name).read_text(encoding="utf-8") for name in document["parts"]]
        self.parts = [parse_part(text) for text in texts]
        self._numbered_comments = {int(value) for text in texts for value in re.findall(r'id="?c(\d+)', text)}
        header = next((part.document for part in self.parts if part.document), None)
        if header is None:
            raise ValueError("The first part must carry the <!-- docforge:document --> header")
        self.header = header
        self.classes = header["classes"]
        self.rels_name = "word/_rels/document.xml.rels"
        self.rels_root = etree.fromstring(self.package.parts[self.rels_name])
        self.media_rels: dict[str, str] = {}
        self.comment_numbers: dict[str, int] = {}
        self.bookmark_numbers: dict[str, int] = {}
        self.comment_blocks: dict[str, Block] = {}
        self.bib_order: dict[str, int] = {}
        self.bib_entries: list[tuple[str, dict]] = []
        bib = document.get("bibliography")
        if bib:
            payload = json.loads((bundle / bib["path"]).read_text(encoding="utf-8-sig"))
            style = bib["style"]
            for key, entry in payload.items():
                if key.startswith("_") or not isinstance(entry, dict):
                    continue
                if style in entry.get("rendered", {}):
                    self.bib_entries.append((key, entry))
                    self.bib_order[key] = len(self.bib_entries)
            self.bib_style = style
        self.comment_style = next(
            (sid for sid, name in self.styles.names.items() if name.lower() == "annotation reference"), None
        )
        self._next_rel = 1

    # -- relationships -------------------------------------------------------
    def _rel_id(self) -> str:
        existing = {rel.get("Id") for rel in self.rels_root}
        while f"rIdm{self._next_rel}" in existing:
            self._next_rel += 1
        return f"rIdm{self._next_rel}"

    def _add_rel(self, rel_type: str, target: str, external: bool = False) -> str:
        rel_id = self._rel_id()
        rel = etree.SubElement(self.rels_root, f"{{{PKG_REL}}}Relationship", Id=rel_id, Type=rel_type, Target=target)
        if external:
            rel.set("TargetMode", "External")
        return rel_id

    def _media_rel(self, source: str) -> str:
        if source in self.media_rels:
            return self.media_rels[source]
        payload = (self.bundle / source).read_bytes()
        digest = hashlib.sha1(payload).hexdigest()
        previous = self._template_images.pop(digest, None)
        if previous is not None:
            # Unchanged pictures keep their relationship id, so a redline
            # against the source sees identical drawing markup.
            rel_id, target = previous
            etree.SubElement(self.rels_root, f"{{{PKG_REL}}}Relationship", Id=rel_id, Type=REL_TYPES["image"], Target=target)
            self.package.parts[re.sub(r"[^/]+/\.\./", "", str(PurePosixPath("word") / target))] = payload
            self._written_media.add(str(PurePosixPath("word") / target))
            self.media_rels[source] = rel_id
            return rel_id
        name = PurePosixPath(source).name
        part = f"word/media/{name}"
        if part in self.package.parts and self.package.parts[part] != payload:
            stem, suffix = name.rsplit(".", 1)
            part = f"word/media/{stem}-{hashlib.sha1(payload).hexdigest()[:8]}.{suffix}"
        self.package.parts[part] = payload
        self._written_media.add(part)
        self._ensure_default(PurePosixPath(part).suffix[1:].lower())
        rel_id = self._add_rel(REL_TYPES["image"], part[len("word/"):])
        self.media_rels[source] = rel_id
        return rel_id

    def _ensure_default(self, extension: str) -> None:
        root = etree.fromstring(self.package.parts["[Content_Types].xml"])
        for item in root.findall(f"{{{CONTENT_TYPES}}}Default"):
            if item.get("Extension", "").lower() == extension:
                return
        etree.SubElement(root, f"{{{CONTENT_TYPES}}}Default", Extension=extension, ContentType=MEDIA_TYPES.get(extension, "application/octet-stream"))
        self.package.parts["[Content_Types].xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _raw(self, atom_body: str, rels: str | None) -> etree._Element:
        declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in self.nsmap.items() if prefix)
        wrapper = etree.fromstring(f"<wrapper {declarations}>{atom_body}</wrapper>")
        node = wrapper[0]
        mapping = {}
        for entry in (rels or "").split(";"):
            fields = entry.split()
            if len(fields) < 3:
                continue
            old, short, target = fields[:3]
            external = len(fields) > 3 and fields[3] == "external"
            if short == "image":
                mapping[old] = self._media_rel(target)
            else:
                rel_type = REL_TYPES.get(short, f"{REL_BASE}/{short}")
                mapping[old] = self._add_rel(rel_type, target, external)
        for element in node.iter():
            for key, value in list(element.attrib.items()):
                if etree.QName(key).namespace == R and value in mapping:
                    element.set(key, mapping[value])
        return node

    # -- paragraphs ------------------------------------------------------------
    def _class(self, name: str | None) -> dict:
        if name is None:
            return {"p": [], "r": []}
        if name not in self.classes:
            raise ValueError(f"Unknown paragraph class: {name}")
        return self.classes[name]

    def paragraph(
        self, text: str, attrs: Attrs, default: str | None, *, level: int | None = None,
        table_style: str | None = None, base_r: list[Token] | None = None,
    ) -> etree._Element:
        cls_name = attrs.cls or (self.header["headings"].get(level) if level is not None else default)
        spec = self._class(cls_name)
        p_tokens = merge(spec["p"], attrs.p)
        r_base = merge(base_r if base_r is not None else spec["r"], attrs.r)
        paragraph = etree.Element(f"{{{W}}}p", nsmap=self.nsmap)
        if p_tokens:
            paragraph.append(build(f"{{{W}}}pPr", p_tokens, self.nsmap))
        self._fill(paragraph, parse_inline(text) if text else [], r_base)
        return paragraph

    def _fill(self, paragraph: etree._Element, items: list, base: list[Token]) -> None:
        groups: list = []
        for item in items:
            if isinstance(item, Atom) and (item.kind in MILESTONES or (item.kind == "ooxml" and item.get("level") == "node")):
                groups.append(("milestone", item))
                continue
            marks = item.marks
            if isinstance(item, Atom) and item.kind == "cite":
                marks = tuple(marks) + (("sup",),)
            link = next((mark for mark in marks if mark[0] == "link"), None)
            rev = next((mark for mark in marks if mark[0] == "rev"), None)
            rest = tuple(mark for mark in marks if mark[0] not in {"link", "rev"})
            groups.append(((link, rev, rest), item))
        index = 0
        while index < len(groups):
            key, item = groups[index]
            if key == "milestone":
                self._milestone(paragraph, item)
                index += 1
                continue
            link = key[0]
            container = paragraph
            if link is not None:
                container = etree.SubElement(paragraph, f"{{{W}}}hyperlink")
                href = link[1]
                if href.startswith("#"):
                    container.set(f"{{{W}}}anchor", href[1:])
                else:
                    container.set(f"{{{R}}}id", self._add_rel(REL_TYPES["hyperlink"], href, True))
            while index < len(groups) and groups[index][0] != "milestone" and groups[index][0][0] == link:
                rev = groups[index][0][1]
                target = container
                if rev is not None:
                    target = etree.SubElement(container, f"{{{W}}}{rev[1]}")
                    for attr, value in rev[2]:
                        target.set(f"{{{W}}}{attr}", value)
                while (
                    index < len(groups) and groups[index][0] != "milestone"
                    and groups[index][0][0] == link and groups[index][0][1] == rev
                ):
                    rest = groups[index][0][2]
                    run = etree.SubElement(target, f"{{{W}}}r")
                    tokens = run_tokens_for(base, rest)
                    if tokens:
                        run.append(build(f"{{{W}}}rPr", tokens, self.nsmap))
                    deleted = rev is not None and rev[1] == "del"
                    while (
                        index < len(groups) and groups[index][0] != "milestone"
                        and groups[index][0] == (link, rev, rest)
                    ):
                        self._content(run, groups[index][1], deleted)
                        index += 1

    def _content(self, run: etree._Element, item, deleted: bool) -> None:
        if isinstance(item, Text):
            node = etree.SubElement(run, f"{{{W}}}{'delText' if deleted else 't'}")
            node.text = item.text
            if item.text != item.text.strip() or "  " in item.text:
                node.set(f"{{{XML_NS}}}space", "preserve")
            return
        kind = item.kind
        if kind == "tab":
            etree.SubElement(run, f"{{{W}}}tab")
        elif kind == "br":
            node = etree.SubElement(run, f"{{{W}}}br")
            for key, value in item.attrs:
                node.set(f"{{{W}}}{key}", value)
        elif kind == "cr":
            etree.SubElement(run, f"{{{W}}}cr")
        elif kind == "sym":
            etree.SubElement(run, f"{{{W}}}sym", {f"{{{W}}}font": item.get("font", ""), f"{{{W}}}char": item.get("char", "")})
        elif kind == "nbhyphen":
            etree.SubElement(run, f"{{{W}}}noBreakHyphen")
        elif kind == "shy":
            etree.SubElement(run, f"{{{W}}}softHyphen")
        elif kind in {"fld-begin", "fld-sep", "fld-end"}:
            node = etree.SubElement(run, f"{{{W}}}fldChar")
            node.set(f"{{{W}}}fldCharType", {"fld-begin": "begin", "fld-sep": "separate", "fld-end": "end"}[kind])
            for key, value in item.attrs:
                node.set(f"{{{W}}}{key}", value)
        elif kind == "instr":
            node = etree.SubElement(run, f"{{{W}}}{'delInstrText' if deleted else 'instrText'}")
            node.text = item.get("code", "")
            node.set(f"{{{XML_NS}}}space", "preserve")
        elif kind == "annotation-ref":
            etree.SubElement(run, f"{{{W}}}annotationRef")
        elif kind == "image":
            attrs = dict(item.attrs)
            rel_id = self._media_rel(attrs.pop("src"))
            run.append(drawing_mod.create(attrs, rel_id, 1))
        elif kind == "cite":
            keys = [key for key in (item.get("keys") or "").split(",") if key]
            missing = [key for key in keys if key not in self.bib_order]
            if missing:
                raise ValueError(f"Citation keys missing from the bibliography: {missing}")
            node = etree.SubElement(run, f"{{{W}}}t")
            node.text = format_citation_numbers([self.bib_order[key] for key in keys])
        elif kind == "ooxml":
            run.append(self._raw(item.body or "", item.get("rels")))
        else:
            raise ValueError(f"Unsupported inline element <{kind}/>")

    def _comment_number(self, md_id: str) -> int:
        if md_id not in self.comment_numbers:
            match = re.fullmatch(r"c(\d+)", md_id or "")
            if match:
                self.comment_numbers[md_id] = int(match.group(1))
            else:
                used = self._numbered_comments | set(self.comment_numbers.values())
                self.comment_numbers[md_id] = max(used, default=-1) + 1
        return self.comment_numbers[md_id]

    def _milestone(self, paragraph: etree._Element, atom: Atom) -> None:
        if atom.kind == "ooxml":
            paragraph.append(self._raw(atom.body or "", atom.get("rels")))
            return
        if atom.kind in {"bookmark-start", "bookmark-end"}:
            name = atom.get("name") or ""
            number = str(self.bookmark_numbers.setdefault(name, len(self.bookmark_numbers)))
            if atom.kind == "bookmark-start":
                etree.SubElement(paragraph, f"{{{W}}}bookmarkStart", {f"{{{W}}}id": number, f"{{{W}}}name": name})
            else:
                etree.SubElement(paragraph, f"{{{W}}}bookmarkEnd", {f"{{{W}}}id": number})
            return
        number = str(self._comment_number(atom.get("id")))
        if atom.kind == "comment-start":
            etree.SubElement(paragraph, f"{{{W}}}commentRangeStart", {f"{{{W}}}id": number})
            return
        if atom.kind == "comment-end":
            etree.SubElement(paragraph, f"{{{W}}}commentRangeEnd", {f"{{{W}}}id": number})
        run = etree.SubElement(paragraph, f"{{{W}}}r")
        if self.comment_style:
            rpr = etree.SubElement(run, f"{{{W}}}rPr")
            etree.SubElement(rpr, f"{{{W}}}rStyle", {f"{{{W}}}val": self.comment_style})
        etree.SubElement(run, f"{{{W}}}commentReference", {f"{{{W}}}id": number})

    # -- blocks ----------------------------------------------------------------
    def table(self, block: Block) -> etree._Element:
        attrs = block.attrs
        table = etree.Element(f"{{{W}}}tbl", nsmap=self.nsmap)
        table.append(build(f"{{{W}}}tblPr", attrs.p, self.nsmap))
        grid = etree.SubElement(table, f"{{{W}}}tblGrid")
        for width in [item for item in attrs.params.get("grid", "").split(",") if item]:
            etree.SubElement(grid, f"{{{W}}}gridCol", {f"{{{W}}}w": width})
        style = next((value for key, value in attrs.p if key == "tblStyle@val"), None)
        default = attrs.params.get("default")
        for row in block.rows:
            tr = etree.SubElement(table, f"{{{W}}}tr")
            first = row[0][0][1] if row and row[0] else Attrs()
            if first.tr:
                tr.append(build(f"{{{W}}}trPr", first.tr, self.nsmap))
            for cell in row:
                tc = etree.SubElement(tr, f"{{{W}}}tc")
                tc_tokens = cell[0][1].tc if cell else []
                if tc_tokens:
                    tc.append(build(f"{{{W}}}tcPr", tc_tokens, self.nsmap))
                for text, para_attrs in cell:
                    tc.append(self.paragraph(text, para_attrs, default, table_style=style))
        return table

    def references(self, block: Block) -> list[etree._Element]:
        attrs = block.attrs
        label = attrs.params.get("label", "")
        spec = self._class(attrs.cls)
        directive_base = merge(spec["r"], attrs.r)
        result = []
        for number, (key, entry) in enumerate(self.bib_entries, 1):
            layout = parse_attrs(entry.get("layout", {}).get(self.bib_style, "{:}"))
            text = label.replace("{n}", str(number)) + entry["rendered"][self.bib_style]
            para_attrs = Attrs(cls=layout.cls or attrs.cls, p=layout.p, r=layout.r)
            result.append(self.paragraph(text, para_attrs, None, base_r=directive_base))
        return result

    def comment_parts(self) -> None:
        name = "word/comments.xml"
        if not self.comment_blocks and name not in self.package.parts:
            return
        if name in self.package.parts:
            old = etree.fromstring(self.package.parts[name])
            root = etree.Element(old.tag, nsmap=old.nsmap)
            for key, value in old.attrib.items():
                root.set(key, value)
        else:
            root = etree.Element(f"{{{W}}}comments", nsmap={"w": W, "w14": W14, "w15": W15})
        ext_root = etree.Element(f"{{{W15}}}commentsEx", nsmap={"w15": W15, "mc": KNOWN_NAMESPACES["mc"]})
        ext_root.set(f"{{{KNOWN_NAMESPACES['mc']}}}Ignorable", "w15")
        last_para: dict[str, str] = {}
        ordered = sorted(self.comment_blocks.items(), key=lambda item: self._comment_number(item[0]))
        for md_id, block in ordered:
            params = block.attrs.params
            comment = etree.SubElement(root, f"{{{W}}}comment")
            comment.set(f"{{{W}}}id", str(self._comment_number(md_id)))
            for key in ("author", "date", "initials"):
                if params.get(key):
                    comment.set(f"{{{W}}}{key}", params[key])
            for index, (text, attrs) in enumerate(block.paras):
                paragraph = self.paragraph(text, attrs, None)
                digest = hashlib.sha1(f"{md_id}:{index}".encode()).hexdigest()[:8].upper()
                para_id = f"{int(digest, 16) & 0x7FFFFFFF:08X}"
                paragraph.set(f"{{{W14}}}paraId", para_id)
                paragraph.set(f"{{{W14}}}textId", "77777777")
                comment.append(paragraph)
                last_para[md_id] = para_id
        for md_id, block in ordered:
            if md_id not in last_para:
                continue
            item = etree.SubElement(ext_root, f"{{{W15}}}commentEx")
            item.set(f"{{{W15}}}paraId", last_para[md_id])
            parent = block.attrs.params.get("parent")
            if parent and parent in last_para:
                item.set(f"{{{W15}}}paraIdParent", last_para[parent])
            item.set(f"{{{W15}}}done", block.attrs.params.get("done", "0"))
        types = etree.fromstring(self.package.parts["[Content_Types].xml"])
        overrides = {item.get("PartName"): item for item in types.findall(f"{{{CONTENT_TYPES}}}Override")}
        # Durable ids and extensible data describe comments that were regenerated.
        for stale in ("word/commentsIds.xml", "word/commentsExtensible.xml"):
            self.package.parts.pop(stale, None)
            if "/" + stale in overrides:
                types.remove(overrides["/" + stale])
            for rel in list(self.rels_root):
                if rel.get("Target") == stale[len("word/"):]:
                    self.rels_root.remove(rel)
        for part, rel_key, content_type in (
            ("word/comments.xml", "comments", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
            ("word/commentsExtended.xml", "commentsExtended", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"),
        ):
            if "/" + part not in overrides:
                etree.SubElement(types, f"{{{CONTENT_TYPES}}}Override", PartName="/" + part, ContentType=content_type)
            if not any(rel.get("Target") == part[len("word/"):] for rel in self.rels_root):
                self._add_rel(REL_TYPES[rel_key], part[len("word/"):])
        self.package.parts["[Content_Types].xml"] = etree.tostring(types, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.package.set_xml("word/comments.xml", root)
        self.package.set_xml("word/commentsExtended.xml", ext_root)

    def build(self) -> Package:
        # Template images and hyperlinks belong to the old body.
        self._template_images: dict[str, tuple[str, str]] = {}
        for rel in list(self.rels_root):
            if rel.get("Type") == REL_TYPES["image"] and rel.get("TargetMode") != "External":
                part = re.sub(r"[^/]+/\.\./", "", str(PurePosixPath("word") / rel.get("Target")))
                payload = self.package.parts.get(part)
                if payload is not None:
                    self._template_images.setdefault(hashlib.sha1(payload).hexdigest(), (rel.get("Id"), rel.get("Target")))
            if rel.get("Type") in {REL_TYPES["image"], REL_TYPES["hyperlink"]}:
                self.rels_root.remove(rel)
        self._written_media: set[str] = set()
        body = self.root.find(f"./{{{W}}}body")
        for child in list(body):
            body.remove(child)
        for part in self.parts:
            for block in part.blocks:
                if block.type == "comment":
                    self.comment_blocks[block.attrs.params["id"]] = block
        for part in self.parts:
            for block in part.blocks:
                if block.type == "p":
                    body.append(self.paragraph(block.text, block.attrs, part.default, level=block.level))
                elif block.type == "table":
                    body.append(self.table(block))
                elif block.type == "references":
                    for paragraph in self.references(block):
                        body.append(paragraph)
                elif block.type == "ooxml":
                    body.append(self._raw(block.body, block.attrs.params.get("rels")))
        section = self.header["section"]
        root_attrs = [(key[1:], value) for key, value in section if key.startswith("@")]
        sect = build(f"{{{W}}}sectPr", [item for item in section if not item[0].startswith("@")], self.nsmap)
        for key, value in root_attrs:
            sect.set(qname(key, self.nsmap), value or "")
        body.append(sect)
        old_background = self.root.find(f"./{{{W}}}background")
        if old_background is not None:
            self.root.remove(old_background)
        if self.header["background"] is not None:
            tokens = self.header["background"]
            background = build(f"{{{W}}}background", [item for item in tokens if not item[0].startswith("@")], self.nsmap)
            for key, value in tokens:
                if key.startswith("@"):
                    background.set(qname(key[1:], self.nsmap), value or "")
            self.root.insert(0, background)
        self._renumber()
        self.comment_parts()
        self.package.set_xml("word/document.xml", self.root)
        self.package.parts[self.rels_name] = etree.tostring(self.rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self._drop_unused_media()
        return self.package

    def _renumber(self) -> None:
        number = 1
        for element in self.root.iter():
            if element.tag in _REVISION_TAGS:
                element.set(f"{{{W}}}id", str(number))
                number += 1
        shape = 1
        for element in self.root.iter(f"{{{WP}}}docPr"):
            element.set("id", str(shape))
            shape += 1

    def _drop_unused_media(self) -> None:
        referenced: set[str] = set()
        for name, payload in self.package.parts.items():
            if not name.endswith(".rels"):
                continue
            base = PurePosixPath(name).parent.parent
            for rel in etree.fromstring(payload):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                path = target[1:] if target.startswith("/") else str(base / target)
                referenced.add(re.sub(r"[^/]+/\.\./", "", path))
        for name in [name for name in self.package.parts if name.startswith("word/media/")]:
            if name not in referenced:
                del self.package.parts[name]


def build_docx(bundle: Path, doc_id: str, output: Path, *, overwrite: bool = False) -> Path:
    """Compile document ``doc_id`` of ``bundle`` into ``output``."""
    manifest = json.loads((bundle / "bundle.json").read_text(encoding="utf-8"))
    document = next((doc for doc in manifest["documents"] if doc["id"] == doc_id), None)
    if document is None:
        raise ValueError(f"Document {doc_id!r} is not in {bundle / 'bundle.json'}")
    package = _Writer(bundle, document).build()
    package.write(output, overwrite=overwrite)
    return output


__all__ = ["build_docx", "parse_part"]
