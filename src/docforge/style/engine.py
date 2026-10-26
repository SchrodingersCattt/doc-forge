"""DOCX StyleEngine.

Restyle edits style definitions and conflicting direct formatting in the
source package. It does not rebuild the document, so unrelated parts stay
byte-identical inside the new package.
"""

from __future__ import annotations

import os
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from ..adapters.protocol import ConversionError
from .profile import StyleProfile, load_profile


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
_PAGE_TWIPS = {"A4": ("11906", "16838"), "LETTER": ("12240", "15840")}
_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}


def restyle_docx(source: Path, output: Path, profile: StyleProfile | Path | str, *, force: bool = False, preserve_content: bool = True) -> dict[str, object]:
    """Apply a style profile to a DOCX and publish only if content and package checks pass."""

    source, output = Path(source), Path(output)
    if source.resolve() == output.resolve():
        raise ConversionError("restyle output must be a different file", code="PATH_COLLISION")
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    loaded = load_profile(profile)
    if loaded.template:
        raise ConversionError(
            "StyleProfile.template is not applied by restyle; remove it until template application is implemented",
            code="STYLE_TEMPLATE",
        )
    with zipfile.ZipFile(source) as package:
        names = package.namelist()
        original = {name: package.read(name) for name in names}
    if "word/styles.xml" not in original or "word/document.xml" not in original:
        raise ConversionError("DOCX is missing styles or document part", code="DOCX_INVALID")
    styles = etree.fromstring(original["word/styles.xml"])
    document = etree.fromstring(original["word/document.xml"])
    before = _content_stream(document)
    applied: list[str] = []
    unused: list[str] = []
    for role_name, role in loaded.roles.items():
        style = _find_style(styles, role.word_style or role_name)
        if style is None:
            if loaded.policy.missing_style == "error":
                raise ConversionError(f"style role has no Word style: {role_name}", code="STYLE_MISSING")
            style = _create_style(styles, role.word_style or role_name)
        _apply_style(style, role)
        style_id = style.get(_q("styleId")) or ""
        matched = 0
        for paragraph in document.xpath(".//w:p", namespaces=NS):
            if role.select:
                chosen = _paragraph_matches_select(paragraph, styles, role.select)
                if chosen and style_id:
                    _set_paragraph_style(paragraph, style_id)
            else:
                chosen = _paragraph_style(paragraph, styles) == style_id
            if not chosen:
                continue
            matched += 1
            _apply_direct(paragraph, role, loaded.policy.direct_formatting)
        if matched:
            applied.append(role_name)
        else:
            unused.append(role_name)
    _apply_page(document, loaded)
    after = _content_stream(document)
    if preserve_content and before != after:
        raise ConversionError("restyle changed document text", code="CONTENT_CHANGED")
    replacements = {
        "word/styles.xml": etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True),
        "word/document.xml": etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True),
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{output.stem}-", suffix=".docx", dir=output.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(temporary, "w") as outgoing:
            for info in incoming.infolist():
                payload = replacements.get(info.filename, incoming.read(info.filename))
                outgoing.writestr(info, payload)
        if preserve_content and paragraph_texts(source) != paragraph_texts(temporary):
            raise ConversionError("restyle changed document text", code="CONTENT_CHANGED")
        from ..adapters.docx import audit_docx

        validation = audit_docx(temporary)
        if not validation.ok:
            raise ConversionError("restyle audit failed", code="AUDIT_FAILED", diagnostics=validation.diagnostics)
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
    return {
        "execution_mode": "SOURCE_PATCH",
        "roles": applied,
        "unused": unused,
        "content_unchanged": before == after,
    }


def resolved_paragraph_styles(path: Path) -> tuple[dict[str, object], ...]:
    """Return the effective paragraph style after docDefaults, basedOn, and direct runs."""

    with zipfile.ZipFile(path) as package:
        styles = etree.fromstring(package.read("word/styles.xml"))
        document = etree.fromstring(package.read("word/document.xml"))
    defaults = _defaults(styles)
    by_id = {style.get(_q("styleId")): style for style in styles.findall(_q("style"))}
    rows = []
    for paragraph in document.xpath(".//w:p", namespaces=NS):
        style_id = _paragraph_style(paragraph, styles) or "Normal"
        effective = dict(defaults)
        for item in reversed(_chain(by_id, style_id)):
            effective.update(_declared(item))
        direct = _direct_run(paragraph)
        rows.append({
            "style_id": style_id,
            "size_pt": direct.get("size_pt", effective.get("size_pt")),
            "bold": direct.get("bold", effective.get("bold")),
            "italic": direct.get("italic", effective.get("italic")),
            "latin": direct.get("latin", effective.get("latin")),
            "east_asia": direct.get("east_asia", effective.get("east_asia")),
            "complex": direct.get("complex", effective.get("complex")),
            "direct_italic": "italic" in direct,
            "direct_bold": "bold" in direct,
            "direct_size": "size_pt" in direct,
            "subscript": direct.get("subscript", False),
            "text": "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)),
        })
    return tuple(rows)


def paragraph_texts(path: Path) -> tuple[str, ...]:
    with zipfile.ZipFile(path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    return tuple(_content_stream(document))


def _content_stream(document: etree._Element) -> list[str]:
    texts = []
    for paragraph in document.xpath(".//w:p", namespaces=NS):
        texts.append("".join(node.text or "" for node in paragraph.xpath(".//w:t | .//w:delText", namespaces=NS)))
    return texts


def _q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _find_style(styles: etree._Element, name: str) -> etree._Element | None:
    folded = name.casefold()
    for style in styles.findall(_q("style")):
        style_id = (style.get(_q("styleId")) or "").casefold()
        name_node = style.find(_q("name"))
        style_name = (name_node.get(_q("val")) if name_node is not None else "").casefold()
        if folded in {style_id, style_name}:
            return style
    return None


def _create_style(styles: etree._Element, name: str) -> etree._Element:
    style_id = "".join(ch for ch in name if ch.isalnum()) or "Role"
    existing = {item.get(_q("styleId")) for item in styles.findall(_q("style"))}
    candidate = style_id
    suffix = 2
    while candidate in existing:
        candidate = f"{style_id}{suffix}"
        suffix += 1
    style = etree.SubElement(styles, _q("style"))
    style.set(_q("type"), "paragraph")
    style.set(_q("styleId"), candidate)
    label = etree.SubElement(style, _q("name"))
    label.set(_q("val"), name)
    return style


def _apply_style(style: etree._Element, role) -> None:
    if role.font.explicit():
        properties = style.find(_q("rPr"))
        if properties is None:
            properties = etree.SubElement(style, _q("rPr"))
        _apply_font(properties, role.font)
    if role.paragraph.explicit():
        properties = style.find(_q("pPr"))
        if properties is None:
            properties = etree.SubElement(style, _q("pPr"))
        _apply_paragraph(properties, role.paragraph)


def _apply_font(properties: etree._Element, font) -> None:
    if any((font.latin, font.east_asia, font.complex)):
        fonts = _child(properties, "rFonts")
        if font.latin:
            fonts.set(_q("ascii"), font.latin)
            fonts.set(_q("hAnsi"), font.latin)
        if font.east_asia:
            fonts.set(_q("eastAsia"), font.east_asia)
        if font.complex:
            fonts.set(_q("cs"), font.complex)
    if font.size_pt is not None:
        half = str(int(round(font.size_pt * 2)))
        _child(properties, "sz").set(_q("val"), half)
        _child(properties, "szCs").set(_q("val"), half)
    _toggle(properties, "b", font.bold)
    _toggle(properties, "i", font.italic)
    if font.color:
        _child(properties, "color").set(_q("val"), font.color.removeprefix("#"))


def _apply_paragraph(properties: etree._Element, paragraph) -> None:
    if paragraph.alignment:
        _child(properties, "jc").set(_q("val"), _ALIGN[paragraph.alignment])
    spacing = None
    if paragraph.before_pt is not None or paragraph.after_pt is not None or paragraph.line_spacing is not None:
        spacing = _child(properties, "spacing")
    if paragraph.before_pt is not None and spacing is not None:
        spacing.set(_q("before"), str(int(round(paragraph.before_pt * 20))))
    if paragraph.after_pt is not None and spacing is not None:
        spacing.set(_q("after"), str(int(round(paragraph.after_pt * 20))))
    if paragraph.line_spacing is not None and spacing is not None:
        spacing.set(_q("line"), str(int(round(paragraph.line_spacing * 240))))
        spacing.set(_q("lineRule"), "auto")
    if paragraph.first_line_pt is not None:
        _child(properties, "ind").set(_q("firstLine"), str(int(round(paragraph.first_line_pt * 20))))
    if paragraph.keep_with_next:
        _child(properties, "keepNext")
    if paragraph.page_break_before:
        _child(properties, "pageBreakBefore")


def _apply_direct(paragraph: etree._Element, role, policy: str) -> None:
    if policy == "preserve":
        return
    for run in paragraph.findall(_q("r")):
        properties = run.find(_q("rPr"))
        if properties is None:
            continue
        if policy == "clear_all":
            for child in list(properties):
                if etree.QName(child).localname not in {"vertAlign", "rStyle"}:
                    properties.remove(child)
            continue
        font = role.font
        if font.size_pt is not None:
            _drop(properties, "sz")
            _drop(properties, "szCs")
        if font.bold is not None:
            _drop(properties, "b")
            _drop(properties, "bCs")
        if font.italic is not None:
            _drop(properties, "i")
            _drop(properties, "iCs")
        if font.color is not None:
            _drop(properties, "color")
        fonts = properties.find(_q("rFonts"))
        if fonts is not None:
            if font.latin:
                for slot in ("ascii", "hAnsi", "asciiTheme", "hAnsiTheme"):
                    fonts.attrib.pop(_q(slot), None)
            if font.east_asia:
                fonts.attrib.pop(_q("eastAsia"), None)
                fonts.attrib.pop(_q("eastAsiaTheme"), None)
            if font.complex:
                fonts.attrib.pop(_q("cs"), None)
                fonts.attrib.pop(_q("csTheme"), None)
            if not fonts.attrib:
                properties.remove(fonts)


def _apply_page(document: etree._Element, profile: StyleProfile) -> None:
    page = profile.page
    if not page.size and not page.margins_mm:
        return
    if page.size and page.size.upper() not in _PAGE_TWIPS:
        raise ConversionError(f"unsupported page size: {page.size}", code="STYLE_PAGE")
    for section in document.xpath(".//w:sectPr", namespaces=NS):
        if page.size:
            width, height = _PAGE_TWIPS[page.size.upper()]
            size = _child(section, "pgSz")
            size.set(_q("w"), width)
            size.set(_q("h"), height)
        if page.margins_mm:
            margins = _child(section, "pgMar")
            for side, value in page.margins_mm.items():
                if side not in {"top", "bottom", "left", "right"}:
                    raise ConversionError(f"unsupported margin side: {side}", code="STYLE_PAGE")
                margins.set(_q(side), str(int(round(float(value) * 56.6929133858))))


def _paragraph_matches_select(paragraph: etree._Element, styles: etree._Element, select: str) -> bool:
    wanted = select.casefold()
    current_id = _paragraph_style(paragraph, styles)
    if current_id is None:
        return wanted == "normal"
    current = _find_style(styles, current_id)
    if current is None:
        return current_id.casefold() == wanted
    return wanted in _style_keys(current)


def _style_keys(style: etree._Element) -> set[str]:
    keys = set()
    style_id = style.get(_q("styleId"))
    if style_id:
        keys.add(style_id.casefold())
    name = style.find(_q("name"))
    if name is not None and name.get(_q("val")):
        keys.add((name.get(_q("val")) or "").casefold())
    return keys


def _set_paragraph_style(paragraph: etree._Element, style_id: str) -> None:
    properties = paragraph.find(_q("pPr"))
    if properties is None:
        properties = etree.Element(_q("pPr"))
        paragraph.insert(0, properties)
    marker = properties.find(_q("pStyle"))
    if marker is None:
        marker = etree.Element(_q("pStyle"))
        properties.insert(0, marker)
    marker.set(_q("val"), style_id)


def _paragraph_style(paragraph: etree._Element, styles: etree._Element) -> str | None:
    properties = paragraph.find(_q("pPr"))
    if properties is None:
        normal = _find_style(styles, "Normal")
        return normal.get(_q("styleId")) if normal is not None else None
    marker = properties.find(_q("pStyle"))
    if marker is None:
        normal = _find_style(styles, "Normal")
        return normal.get(_q("styleId")) if normal is not None else None
    return marker.get(_q("val"))


def _child(parent: etree._Element, tag: str) -> etree._Element:
    child = parent.find(_q(tag))
    if child is None:
        child = etree.SubElement(parent, _q(tag))
    return child


def _drop(parent: etree._Element, tag: str) -> None:
    child = parent.find(_q(tag))
    if child is not None:
        parent.remove(child)


def _toggle(parent: etree._Element, tag: str, value: bool | None) -> None:
    if value is None:
        return
    child = _child(parent, tag)
    if value:
        child.attrib.pop(_q("val"), None)
    else:
        child.set(_q("val"), "0")


def _defaults(styles: etree._Element) -> dict[str, object]:
    node = styles.find(_q("docDefaults"))
    if node is None:
        return {}
    run = node.find(f"{_q('rPrDefault')}/{_q('rPr')}")
    return _declared_properties(run) if run is not None else {}


def _chain(by_id: dict[str, etree._Element], style_id: str | None) -> list[etree._Element]:
    chain = []
    seen: set[str] = set()
    current = style_id
    while current and current not in seen and current in by_id:
        seen.add(current)
        chain.append(by_id[current])
        based = by_id[current].find(_q("basedOn"))
        current = based.get(_q("val")) if based is not None else None
    return chain


def _declared(style: etree._Element) -> dict[str, object]:
    return _declared_properties(style.find(_q("rPr")))


def _declared_properties(properties: etree._Element | None) -> dict[str, object]:
    if properties is None:
        return {}
    values: dict[str, object] = {}
    fonts = properties.find(_q("rFonts"))
    if fonts is not None:
        if fonts.get(_q("ascii")):
            values["latin"] = fonts.get(_q("ascii"))
        if fonts.get(_q("eastAsia")):
            values["east_asia"] = fonts.get(_q("eastAsia"))
        if fonts.get(_q("cs")):
            values["complex"] = fonts.get(_q("cs"))
    size = properties.find(_q("sz"))
    if size is not None and size.get(_q("val")):
        values["size_pt"] = int(size.get(_q("val"))) / 2
    bold = properties.find(_q("b"))
    if bold is not None:
        values["bold"] = bold.get(_q("val"), "true") not in {"0", "false", "off"}
    italic = properties.find(_q("i"))
    if italic is not None:
        values["italic"] = italic.get(_q("val"), "true") not in {"0", "false", "off"}
    return values


def _direct_run(paragraph: etree._Element) -> dict[str, object]:
    """Merge direct run properties. Emphasis and scripts count if any run has them."""

    values: dict[str, object] = {}
    for run in paragraph.findall(_q("r")):
        declared = _declared_properties(run.find(_q("rPr")))
        for key, value in declared.items():
            if key in {"bold", "italic"}:
                values[key] = bool(values.get(key)) or bool(value)
            elif key not in values:
                values[key] = value
        vertical = run.find(f"{_q('rPr')}/{_q('vertAlign')}")
        if vertical is not None and vertical.get(_q("val")) == "subscript":
            values["subscript"] = True
    return values


__all__ = ["paragraph_texts", "resolved_paragraph_styles", "restyle_docx"]
