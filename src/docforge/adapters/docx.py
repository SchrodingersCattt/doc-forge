"""DOCX semantic projection, package audit, and exact-reuse metadata.

Rendering a DOCX from scratch stays with the legacy Markdown and TeX
converters. This module reads a package into SemanticDocument and checks
relationships without rebuilding the file.
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from pathlib import Path, PurePosixPath

from lxml import etree

from ..pipeline.models import Diagnostic, ValidationReport
from ..semantic import BlockNode, DocumentMetadata, InlineNode, Provenance, ResourceDescriptor, SemanticDocument, text_node
from .protocol import FormatCapabilities, ReadOptions, ReadResult, SourceContext, finding


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W, "r": R, "pr": REL, "a": A, "wp": WP, "v": V, "m": M, "mc": MC}
_REL_ATTRS = frozenset({"id", "embed", "link"})
_SKIP_BODY = frozenset({
    "sectPr", "bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd",
    "commentRangeStart", "commentRangeEnd", "moveFromRangeStart", "moveFromRangeEnd",
    "moveToRangeStart", "moveToRangeEnd",
})
_SKIP_RUN = frozenset({
    "rPr", "lastRenderedPageBreak", "fldChar", "instrText", "separator",
    "continuationSeparator", "pgNum", "ruby", "ptab", "proofErr",
})
_CORE_PARTS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/styles.xml",
    "word/settings.xml",
    "word/fontTable.xml",
    "word/numbering.xml",
    "word/_rels/document.xml.rels",
}
_MEDIA_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
}

_CAPABILITIES = FormatCapabilities(
    format_id="docx",
    node_kinds=("paragraph", "heading", "table", "figure"),
    citations=False,
    media=True,
    annotations=False,
    preservation=("EXACT_REUSE", "SOURCE_PATCH", "REBUILD"),
    notes=("Semantic DOCX reading drops unsupported native constructs into diagnostics and SourceContext.",),
)


class DocxReader:
    format_id = "docx"

    def probe(self, source: Path) -> FormatCapabilities:
        from ..output import detect_format

        detect_format(source)
        return _CAPABILITIES

    def read(self, source: Path, *, options: ReadOptions) -> ReadResult:
        from ..output import detect_format

        source = Path(source)
        detect_format(source)
        payload = source.read_bytes()
        digest = hashlib.sha256(payload).hexdigest()
        document, diagnostics, resources, features, opaque = project_docx(source)
        return ReadResult(
            document=document,
            source=SourceContext(
                "docx", digest, source.name,
                opaque_parts=tuple(opaque),
                retained_features=tuple(features),
            ),
            diagnostics=tuple(diagnostics),
            capabilities=_CAPABILITIES,
            origin_format="docx",
            origin_name=source.name,
            content_hash=document.semantic_hash(),
        )


def project_docx(path: Path) -> tuple[SemanticDocument, list[Diagnostic], list[ResourceDescriptor], list[str], list[str]]:
    path = Path(path)
    with zipfile.ZipFile(path) as package:
        projector = _Projector(package, path)
        document, diagnostics, resources, features, opaque = projector.project()
    return document, diagnostics, resources, features, opaque


class _Projector:
    """Walk ``w:body`` in document order, including nested tables and drawings."""

    def __init__(self, package: zipfile.ZipFile, path: Path) -> None:
        self.package = package
        self.path = path
        self.names = set(package.namelist())
        self.provenance = Provenance(path.name)
        self.styles = _style_names(package, self.names)
        self.part = "word/document.xml"
        self.rels = _read_relationships(package.read(self.part_rels)) if self.part_rels in self.names else {}
        self.document_rels = dict(self.rels)
        self.diagnostics: list[Diagnostic] = []
        self.resources: list[ResourceDescriptor] = []
        self.resource_ids: dict[str, str] = {}
        self.referenced: set[str] = set()
        self.features: list[str] = []
        self.opaque = _opaque_parts(self.names)
        self.title = ""
        self._counter = 0
        self._noted: set[str] = set()

    @property
    def part_rels(self) -> str:
        return rels_name_for(self.part)

    def project(self) -> tuple[SemanticDocument, list[Diagnostic], list[ResourceDescriptor], list[str], list[str]]:
        self._prepare_media()
        document_xml = self.package.read("word/document.xml")
        self._note_package_features(document_xml)
        root = etree.fromstring(document_xml)
        body = root.find(f"{{{W}}}body")
        blocks = self._container(body if body is not None else root)
        blocks.extend(self._stories())
        self._note_unreferenced_media()
        if self.resources and "media" not in self.features:
            self.features.append("media")
        document = SemanticDocument(
            metadata=DocumentMetadata(title=self.title),
            blocks=tuple(blocks),
            resources=tuple(self.resources),
        )
        return document, self.diagnostics, self.resources, self.features, self.opaque

    def _prepare_media(self) -> None:
        for name in sorted(self.names):
            if name.startswith("word/media/") and not name.endswith("/"):
                self._ensure_resource(name)

    def _note_package_features(self, document_xml: bytes) -> None:
        if b"w:comment" in document_xml or "word/comments.xml" in self.names:
            self.features.append("comments")
            self._note("docx.comments", "comments are retained in the semantic projection", loss=False)
        if b"w:ins" in document_xml or b"w:del" in document_xml:
            self.features.append("revisions")
            self._note("docx.revisions", "tracked revisions are retained in the semantic projection", loss=False)

    def _note_unreferenced_media(self) -> None:
        missing = [part for part in self.resource_ids if part not in self.referenced]
        if missing:
            self._note("docx.unreferenced_media", "a media part is not referenced by a projected figure", loss=True)

    def _stories(self) -> list[BlockNode]:
        blocks: list[BlockNode] = []
        suffixes = ("/header", "/footer", "/footnotes", "/endnotes", "/comments")
        for target, rel_type, external in self.document_rels.values():
            if external or not rel_type.endswith(suffixes):
                continue
            part = resolve_opc_target(self.part, target)
            if not part or part not in self.names:
                continue
            root = etree.fromstring(self.package.read(part))
            previous = (self.part, self.rels)
            self.part = part
            self.rels = _read_relationships(self.package.read(self.part_rels)) if self.part_rels in self.names else {}
            if rel_type.endswith("/comments"):
                for comment in root.findall(f"{{{W}}}comment"):
                    blocks.extend(self._container(comment))
            elif rel_type.endswith(("/footnotes", "/endnotes")):
                for note in root:
                    if note.get(f"{{{W}}}type") in {"separator", "continuationSeparator", "continuationNotice"}:
                        continue
                    blocks.extend(self._container(note))
            else:
                blocks.extend(self._container(root))
            self.part, self.rels = previous
        if blocks:
            self._note("docx.stories", "headers, footers, notes, and comments are projected after the body", loss=False)
        return blocks

    def _container(self, element: etree._Element | None) -> list[BlockNode]:
        if element is None:
            return []
        blocks: list[BlockNode] = []
        for child in element:
            blocks.extend(self._dispatch(child))
        return blocks

    def _dispatch(self, element: etree._Element) -> list[BlockNode]:
        name = etree.QName(element)
        local = name.localname
        if local == "AlternateContent":
            choice = element.find(f"{{{MC}}}Choice")
            fallback = element.find(f"{{{MC}}}Fallback")
            return self._container(choice if choice is not None else fallback)
        if name.namespace == M or local in {"oMath", "oMathPara"}:
            return [self._math(element)]
        if local == "p":
            return self._paragraph(element)
        if local == "tbl":
            return [self._table(element)]
        if local == "sdt":
            content = element.find(f"{{{W}}}sdtContent")
            self._note("docx.content_control", "content control chrome is not modeled; its content is projected", loss=False)
            return self._container(content)
        if local in _SKIP_BODY:
            return []
        if local in {"drawing", "pict", "object"}:
            return self._drawing(element)
        return self._unknown(element)

    def _paragraph(self, element: etree._Element) -> list[BlockNode]:
        style_id = _paragraph_style_id(element)
        produced: list[BlockNode] = []
        inlines: list[InlineNode] = []

        def flush() -> None:
            if not any(node.kind != "text" or node.text for node in inlines):
                inlines.clear()
                return
            produced.append(self._text_block(tuple(inlines), style_id))
            inlines.clear()

        self._consume_inlines(element, inlines, produced, flush)
        flush()
        return produced

    def _consume_inlines(self, element: etree._Element, inlines: list[InlineNode], produced: list[BlockNode], flush) -> None:
        for child in element:
            name = etree.QName(child)
            local = name.localname
            if name.namespace == M or local in {"oMath", "oMathPara"}:
                flush()
                produced.append(self._math(child))
                continue
            if local == "hyperlink":
                self._hyperlink(child, inlines, produced, flush)
                continue
            if local in {"ins", "moveTo", "smartTag", "fldSimple"}:
                self._consume_inlines(child, inlines, produced, flush)
                continue
            if local in {"del", "moveFrom"}:
                self._deleted(child, inlines, produced, flush)
                continue
            if local == "sdt":
                content = child.find(f"{{{W}}}sdtContent")
                if content is not None:
                    self._consume_inlines(content, inlines, produced, flush)
                continue
            if local == "r":
                self._run(child, inlines, produced, flush)
                continue
            if local in _SKIP_BODY or local == "pPr":
                continue
            if local in {"drawing", "pict", "object"} or _contains_drawing(child):
                flush()
                produced.extend(self._drawing(child))
                continue
            text = "".join(child.itertext()).strip()
            if text or local not in {"bookmarkStart", "bookmarkEnd"}:
                self._note("docx.unmodeled", f"unmodeled element {local}", loss=True)
                if text:
                    inlines.append(text_node(text))

    def _run(self, run: etree._Element, inlines: list[InlineNode], produced: list[BlockNode], flush) -> None:
        for child in run:
            local = etree.QName(child).localname
            if local == "t" and child.text:
                inlines.append(_format_run(run, child.text))
            elif local == "delText" and child.text:
                inlines.append(InlineNode("strike", children=(_format_run(run, child.text),)))
            elif local == "tab":
                inlines.append(text_node("\t"))
            elif local in {"br", "cr"}:
                inlines.append(InlineNode("break"))
            elif local == "noBreakHyphen":
                inlines.append(text_node("-"))
            elif local in {"drawing", "pict", "object"} or _contains_drawing(child):
                flush()
                produced.extend(self._drawing(child))
            elif local in {"footnoteReference", "endnoteReference", "commentReference"}:
                self._note(f"docx.{local}", f"{local} targets are projected with their story", loss=False)
            elif local == "instrText":
                self._note("docx.fields", "field codes are not modeled; displayed field results are kept", loss=False)
            elif local in _SKIP_RUN or local == "rPr":
                continue
            elif "".join(child.itertext()).strip():
                self._note("docx.unmodeled", f"unmodeled element {local}", loss=True)
                inlines.append(text_node("".join(child.itertext())))

    def _hyperlink(self, element: etree._Element, inlines: list[InlineNode], produced: list[BlockNode], flush) -> None:
        rel_id = element.get(f"{{{R}}}id")
        anchor = element.get(f"{{{W}}}anchor") or ""
        href = f"#{anchor}" if anchor else ""
        if rel_id and rel_id in self.rels:
            target, _rel_type, external = self.rels[rel_id]
            href = target if external or "://" in target else href or target
        children: list[InlineNode] = []
        nested: list[BlockNode] = []

        def nested_flush() -> None:
            return None

        self._consume_inlines(element, children, nested, nested_flush)
        if children:
            inlines.append(InlineNode("link", attrs=(("href", href),), children=tuple(children)))
        if nested:
            flush()
            produced.extend(nested)

    def _deleted(self, element: etree._Element, inlines: list[InlineNode], produced: list[BlockNode], flush) -> None:
        for node in element.iter(f"{{{W}}}delText"):
            if node.text:
                inlines.append(InlineNode("strike", children=(text_node(node.text),)))
        for drawing in element.iter(f"{{{W}}}drawing"):
            flush()
            produced.extend(self._drawing(drawing))
        for pict in element.iter(f"{{{W}}}pict"):
            flush()
            produced.extend(self._drawing(pict))

    def _table(self, element: etree._Element) -> BlockNode:
        rows = []
        for row in element.findall(f"{{{W}}}tr"):
            cells = []
            for cell in row.findall(f"{{{W}}}tc"):
                children: list[BlockNode] = []
                for child in cell:
                    if etree.QName(child).localname == "tcPr":
                        continue
                    children.extend(self._dispatch(child))
                attrs: list[tuple[str, str]] = []
                grid = cell.find(f"{{{W}}}tcPr/{{{W}}}gridSpan")
                if grid is not None and str(grid.get(f"{{{W}}}val") or "").isdigit():
                    attrs.append(("colspan", str(grid.get(f"{{{W}}}val"))))
                cells.append(BlockNode(self._next_id(), "table_cell", children=tuple(children), attrs=tuple(attrs), provenance=self.provenance))
            rows.append(BlockNode(self._next_id(), "table_row", children=tuple(cells), provenance=self.provenance))
        return BlockNode(self._next_id(), "table", children=tuple(rows), provenance=self.provenance)

    def _drawing(self, element: etree._Element) -> list[BlockNode]:
        contents = _textbox_contents(element)
        if contents:
            self._note("docx.textboxes", "text boxes are not a semantic node; their text is projected in place", loss=True)
            if "textboxes" not in self.features:
                self.features.append("textboxes")
            blocks: list[BlockNode] = []
            for content in contents:
                blocks.extend(self._container(content))
            return blocks
        descr = _drawing_descr(element)
        rel_ids = _image_rel_ids(element)
        if not rel_ids:
            self._note("docx.unmodeled_drawing", "drawing has no projected image or text", loss=True)
            text = "".join(element.itertext()).strip()
            return [BlockNode(self._next_id(), "opaque", attrs=(("element", "drawing"), ("text", text)), provenance=self.provenance)]
        return [self._figure(rel_id, descr) for rel_id in rel_ids]

    def _figure(self, rel_id: str, descr: str) -> BlockNode:
        relationship = self.rels.get(rel_id)
        if relationship is None:
            self._note("docx.dangling_drawing", "drawing references a missing relationship", loss=True)
            return BlockNode(self._next_id(), "opaque", attrs=(("element", "drawing"), ("rel", rel_id)), provenance=self.provenance)
        target, _rel_type, external = relationship
        if external or "://" in target:
            self._note("docx.external_image", "linked image is outside the package", loss=True)
            return BlockNode(self._next_id(), "figure", attrs=(("resource", ""), ("path", ""), ("href", target)), provenance=self.provenance)
        part = resolve_opc_target(self.part, target)
        if not part or part not in self.names:
            self._note("docx.missing_media", "drawing target is missing from the package", loss=True)
            return BlockNode(self._next_id(), "opaque", attrs=(("element", "drawing"), ("target", target)), provenance=self.provenance)
        resource_id = self._ensure_resource(part)
        self.referenced.add(part)
        children = ()
        if descr:
            children = (BlockNode(self._next_id(), "caption", inlines=(text_node(descr),), provenance=self.provenance),)
        return BlockNode(
            self._next_id(), "figure", children=children,
            attrs=(("resource", resource_id), ("path", PurePosixPath(part).name)),
            provenance=self.provenance,
        )

    def _math(self, element: etree._Element) -> BlockNode:
        self._note("docx.omml", "OMML is kept as opaque text because it has no TeX projection", loss=True)
        text = "".join(node.text or "" for node in element.xpath(".//m:t", namespaces=NS))
        return BlockNode(self._next_id(), "opaque", attrs=(("element", "omml"), ("text", text)), provenance=self.provenance)

    def _unknown(self, element: etree._Element) -> list[BlockNode]:
        local = etree.QName(element).localname
        inner: list[BlockNode] = []
        for child in element:
            inner.extend(self._dispatch(child))
        self._note("docx.unmodeled", f"unmodeled element {local}", loss=True)
        if inner:
            return inner
        text = "".join(element.itertext()).strip()
        return [BlockNode(self._next_id(), "opaque", attrs=(("element", local), ("text", text)), provenance=self.provenance)]

    def _text_block(self, inlines: tuple[InlineNode, ...], style_id: str | None) -> BlockNode:
        plain = "".join(node.plain() for node in inlines)
        if _is_title(style_id, self.styles) and not self.title:
            self.title = plain
        level = _heading_level(style_id, self.styles)
        if level:
            return BlockNode(self._next_id(), "heading", inlines=inlines, attrs=(("level", level),), provenance=self.provenance)
        style_name = self.styles.get(style_id or "", style_id or "")
        attrs = (("style", style_name),) if style_name else ()
        return BlockNode(self._next_id(), "paragraph", inlines=inlines, attrs=attrs, provenance=self.provenance)

    def _ensure_resource(self, part: str) -> str:
        existing = self.resource_ids.get(part)
        if existing:
            return existing
        payload = self.package.read(part) if part in self.names else b""
        resource_id = "media-" + part.replace("/", "-")
        suffix = PurePosixPath(part).suffix.casefold()
        self.resource_ids[part] = resource_id
        self.resources.append(ResourceDescriptor(
            resource_id,
            _MEDIA_TYPES.get(suffix, "application/octet-stream"),
            hashlib.sha256(payload).hexdigest() if payload else "",
            PurePosixPath(part).name,
        ))
        return resource_id

    def _note(self, code: str, message: str, *, loss: bool) -> None:
        if code in self._noted:
            return
        self._noted.add(code)
        if loss:
            self.diagnostics.append(finding("warning", code, message, category="loss"))
        else:
            self.diagnostics.append(finding("info", code, message))

    def _next_id(self) -> str:
        self._counter += 1
        return f"b{self._counter}"


def audit_docx(path: Path) -> ValidationReport:
    """Check package structure and OPC relationship targets. Does not render the file."""

    from ..output import detect_format

    path = Path(path)
    try:
        detect_format(path)
    except Exception as exc:  # format errors are already user-facing
        return ValidationReport("fail", (finding("error", "docx.invalid", str(exc)),))
    diagnostics: list[Diagnostic] = []
    with zipfile.ZipFile(path) as package:
        corrupt = package.testzip()
        if corrupt is not None:
            return ValidationReport("fail", (finding("error", "docx.corrupt", "DOCX member failed CRC"),))
        names = set(package.namelist())
        if "word/document.xml" not in names:
            return ValidationReport("fail", (finding("error", "docx.invalid", "DOCX is missing word/document.xml"),))
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in names:
            return ValidationReport("fail", (finding("error", "docx.relationships", "document relationships part is missing"),))
        for name in sorted(item for item in names if item.endswith(".rels")):
            owner = relationship_owner(name)
            if owner is None:
                diagnostics.append(finding("error", "docx.relationships", "relationship part has no owner", owner=Path(name).name))
                continue
            relationships = _read_relationships(package.read(name))
            diagnostics.extend(_missing_targets(relationships, owner, names))
            if owner and owner in names and owner.endswith(".xml"):
                diagnostics.extend(_dangling_references(package.read(owner), relationships, owner))
    status = "fail" if any(item.severity == "error" for item in diagnostics) else "pass"
    return ValidationReport(status, tuple(diagnostics), {"relationships": "checked"})


def resolve_opc_target(source_part: str, target: str) -> str | None:
    """Resolve an OPC relationship target against the part that owns it.

    ``None`` means the target is external or only a fragment. An empty string
    means the relative path escaped the package.
    """

    raw = target.replace("\\", "/").strip()
    if not raw or raw.startswith("#") or "://" in raw:
        return None
    body = raw.split("#", 1)[0]
    if not body:
        return None
    absolute = body.startswith("/")
    if absolute:
        stack: list[str] = []
    else:
        parent = PurePosixPath(source_part).parent
        stack = [] if str(parent) in {"", "."} else list(parent.parts)
    for piece in body.lstrip("/").split("/"):
        if piece in {"", "."}:
            continue
        if piece == "..":
            if not stack:
                return ""
            stack.pop()
            continue
        stack.append(piece)
    return "/".join(stack)


def relationship_owner(rels_name: str) -> str | None:
    """Return the part that owns a ``.rels`` item, or ``\"\"`` for the package root."""

    path = PurePosixPath(rels_name)
    if path.name == ".rels" and path.parent.name == "_rels":
        return ""
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        return None
    owner_name = path.name[: -len(".rels")]
    parent = path.parent.parent
    if str(parent) in {"", "."}:
        return owner_name
    return f"{parent.as_posix()}/{owner_name}"


def rels_name_for(part: str) -> str:
    path = PurePosixPath(part)
    return f"{path.parent.as_posix()}/_rels/{path.name}.rels".removeprefix("./")


def paragraph_text(path: Path) -> tuple[str, ...]:
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    texts = []
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        value = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
        texts.append(value)
    return tuple(texts)


def _read_relationships(payload: bytes) -> dict[str, tuple[str, str, bool]]:
    root = etree.fromstring(payload)
    relationships: dict[str, tuple[str, str, bool]] = {}
    for item in root.findall(f"{{{REL}}}Relationship"):
        rel_id = item.get("Id")
        if not rel_id:
            continue
        relationships[rel_id] = (
            item.get("Target") or "",
            item.get("Type") or "",
            (item.get("TargetMode") or "").casefold() == "external",
        )
    return relationships


def _missing_targets(relationships: dict[str, tuple[str, str, bool]], owner: str, names: set[str]) -> list[Diagnostic]:
    findings: list[Diagnostic] = []
    for target, rel_type, external in relationships.values():
        if external:
            continue
        resolved = resolve_opc_target(owner, target)
        if resolved is None:
            continue
        if resolved not in names:
            findings.append(finding(
                "error", _relationship_code(rel_type),
                "relationship target is missing from the package",
                owner=Path(owner).name if owner else "[package]",
            ))
    return findings


def _dangling_references(payload: bytes, relationships: dict[str, tuple[str, str, bool]], owner: str) -> list[Diagnostic]:
    root = etree.fromstring(payload)
    external = {rel_id for rel_id, (_target, _rel_type, is_external) in relationships.items() if is_external}
    findings: list[Diagnostic] = []
    for node in root.iter():
        local = etree.QName(node).localname
        for attr, value in node.attrib.items():
            attribute = etree.QName(attr)
            if attribute.namespace != R or attribute.localname not in _REL_ATTRS or not value:
                continue
            if value in external:
                continue
            if value not in relationships:
                findings.append(finding(
                    "error",
                    "docx.header_footer" if local in {"headerReference", "footerReference"} else "docx.dangling_relationship",
                    "element references a missing relationship",
                    owner=Path(owner).name,
                ))
    return findings


def _relationship_code(rel_type: str, local: str = "") -> str:
    if local in {"headerReference", "footerReference"} or rel_type.endswith("/header") or rel_type.endswith("/footer"):
        return "docx.header_footer"
    return "docx.missing_target"


def _style_names(package: zipfile.ZipFile, names: set[str]) -> dict[str, str]:
    if "word/styles.xml" not in names:
        return {}
    root = etree.fromstring(package.read("word/styles.xml"))
    mapping: dict[str, str] = {}
    for style in root.findall(f"{{{W}}}style"):
        style_id = style.get(f"{{{W}}}styleId")
        if not style_id:
            continue
        name = style.find(f"{{{W}}}name")
        mapping[style_id] = name.get(f"{{{W}}}val") if name is not None and name.get(f"{{{W}}}val") else style_id
    return mapping


def _paragraph_style_id(paragraph: etree._Element) -> str | None:
    properties = paragraph.find(f"{{{W}}}pPr")
    if properties is None:
        return None
    marker = properties.find(f"{{{W}}}pStyle")
    if marker is None:
        return None
    return marker.get(f"{{{W}}}val")


def _heading_level(style_id: str | None, styles: dict[str, str]) -> str | None:
    if not style_id:
        return None
    for candidate in (styles.get(style_id, ""), style_id):
        match = re.search(r"heading\s*(\d+)", candidate, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    return None


def _is_title(style_id: str | None, styles: dict[str, str]) -> bool:
    if not style_id:
        return False
    return styles.get(style_id, "").casefold() == "title" or style_id.casefold() == "title"


def _format_run(run: etree._Element, text: str) -> InlineNode:
    node: InlineNode = text_node(text)
    properties = run.find(f"{{{W}}}rPr")
    if properties is None:
        return node
    vertical = properties.find(f"{{{W}}}vertAlign")
    if vertical is not None:
        value = vertical.get(f"{{{W}}}val")
        if value == "subscript":
            return InlineNode("subscript", children=(node,))
        if value == "superscript":
            return InlineNode("superscript", children=(node,))
    bold = _toggle(properties, "b") or _toggle(properties, "bCs")
    italic = _toggle(properties, "i") or _toggle(properties, "iCs")
    if bold and italic:
        node = InlineNode("strong", children=(InlineNode("emph", children=(node,)),))
    elif bold:
        node = InlineNode("strong", children=(node,))
    elif italic:
        node = InlineNode("emph", children=(node,))
    underline = properties.find(f"{{{W}}}u")
    if underline is not None and underline.get(f"{{{W}}}val", "single") not in {"0", "false", "off", "none"}:
        node = InlineNode("underline", children=(node,))
    return node


def _toggle(properties: etree._Element, tag: str) -> bool:
    element = properties.find(f"{{{W}}}{tag}")
    if element is None:
        return False
    return element.get(f"{{{W}}}val", "true") not in {"0", "false", "off", "none"}


def _contains_drawing(element: etree._Element) -> bool:
    return (
        element.find(f".//{{{W}}}drawing") is not None
        or element.find(f".//{{{W}}}pict") is not None
        or element.find(f".//{{{W}}}txbxContent") is not None
    )


def _textbox_contents(element: etree._Element) -> list[etree._Element]:
    choices = element.xpath(".//mc:Choice//w:txbxContent", namespaces=NS)
    if choices:
        return list(choices)
    fallbacks = element.xpath(".//mc:Fallback//w:txbxContent", namespaces=NS)
    if fallbacks:
        return list(fallbacks)
    return list(element.xpath(".//w:txbxContent", namespaces=NS))


def _drawing_descr(element: etree._Element) -> str:
    for node in element.xpath(".//wp:docPr", namespaces=NS):
        descr = node.get("descr") or node.get("title") or ""
        if descr:
            return descr
    return ""


def _image_rel_ids(element: etree._Element) -> list[str]:
    found: list[str] = []
    for node in element.xpath(".//a:blip", namespaces=NS):
        rel_id = node.get(f"{{{R}}}embed") or node.get(f"{{{R}}}link")
        if rel_id:
            found.append(rel_id)
    for node in element.xpath(".//v:imagedata", namespaces=NS):
        rel_id = node.get(f"{{{R}}}id")
        if rel_id:
            found.append(rel_id)
    return found


def _opaque_parts(names: set[str]) -> list[str]:
    opaque = []
    for name in sorted(names):
        if name.startswith("word/media/") or name.startswith("word/theme/"):
            continue
        if name in _CORE_PARTS:
            continue
        opaque.append(name)
    return opaque


__all__ = [
    "DocxReader", "audit_docx", "paragraph_text", "project_docx",
    "relationship_owner", "resolve_opc_target",
]
