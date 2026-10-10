"""Inline pictures as ``![alt](media/x.png){cx=.. cy=..}`` attributes."""

from __future__ import annotations

from lxml import etree

from .ooxml import KNOWN_NAMESPACES, R

WP = KNOWN_NAMESPACES["wp"]
A = KNOWN_NAMESPACES["a"]
PIC = KNOWN_NAMESPACES["pic"]
A14 = KNOWN_NAMESPACES["a14"]
WP14 = KNOWN_NAMESPACES["wp14"]
_LOCAL_DPI = "{28A0092B-C50C-407E-A947-70E740481C1C}"
_SHADOW = "{53640926-AAD7-44D8-BBD7-CCE9431645EC}"
_IGNORED_ATTRS = {f"{{{WP14}}}anchorId", f"{{{WP14}}}editId", f"{{{R}}}embed"}
_DIST = ("distT", "distB", "distL", "distR")


def _edges(node: etree._Element | None, names: tuple[str, ...]) -> str | None:
    if node is None:
        return None
    values = [node.get(name, "0") for name in names]
    return ",".join(values)


def _shape_extras(sp_pr: etree._Element | None) -> str | None:
    """Comma list of the extra ``pic:spPr`` flags Word writes after the geometry."""
    if sp_pr is None:
        return None
    flags = []
    if sp_pr.get("bwMode"):
        flags.append("bw-" + sp_pr.get("bwMode"))
    if sp_pr.find(f"./{{{A}}}noFill") is not None:
        flags.append("nofill")
    if sp_pr.find(f"./{{{A}}}ln/{{{A}}}noFill") is not None:
        flags.append("noline")
    if sp_pr.find(f"./{{{A}}}extLst/{{{A}}}ext[@uri='{_SHADOW}']") is not None:
        flags.append("noshadow")
    geom = sp_pr.find(f"./{{{A}}}prstGeom")
    if geom is not None and geom.find(f"./{{{A}}}avLst") is None:
        flags.append("noav")
    return ",".join(flags)


def describe(drawing: etree._Element) -> tuple[list[tuple[str, str]], str] | None:
    """Return ``(attrs, rel_id)`` when ``drawing`` is a plain inline picture."""
    inline = drawing.find(f"./{{{WP}}}inline")
    if inline is None or len(drawing) != 1:
        return None
    blip = inline.find(f".//{{{A}}}blip")
    extent = inline.find(f"./{{{WP}}}extent")
    doc_pr = inline.find(f"./{{{WP}}}docPr")
    pic_pr = inline.find(f".//{{{PIC}}}cNvPr")
    if blip is None or extent is None or doc_pr is None or blip.get(f"{{{R}}}embed") is None:
        return None
    attrs: list[tuple[str, str]] = [("alt", doc_pr.get("descr", ""))]
    attrs.append(("cx", extent.get("cx", "0")))
    attrs.append(("cy", extent.get("cy", "0")))
    effect = _edges(inline.find(f"./{{{WP}}}effectExtent"), ("l", "t", "r", "b"))
    if effect is None:
        attrs.append(("effect", "none"))
    elif effect != "0,0,0,0":
        attrs.append(("effect", effect))
    if not any(inline.get(name) is not None for name in _DIST):
        attrs.append(("dist", "none"))
    else:
        dist = ",".join(inline.get(name, "0") for name in _DIST)
        if dist != "0,0,0,0":
            attrs.append(("dist", dist))
    attrs.append(("name", doc_pr.get("name", "")))
    if doc_pr.get("title"):
        attrs.append(("title", doc_pr.get("title")))
    if pic_pr is not None and pic_pr.get("name", "") != doc_pr.get("name", ""):
        attrs.append(("picname", pic_pr.get("name", "")))
    if pic_pr is not None and pic_pr.get("id", "0") != "0":
        attrs.append(("picid", "1"))
    if blip.get("cstate"):
        attrs.append(("cstate", blip.get("cstate")))
    dpi = blip.find(f"./{{{A}}}extLst/{{{A}}}ext[@uri='{_LOCAL_DPI}']/{{{A14}}}useLocalDpi")
    if dpi is not None:
        attrs.append(("localdpi", dpi.get("val", "1")))
    fill = inline.find(f".//{{{PIC}}}blipFill")
    if fill is not None and fill.get("rotWithShape"):
        attrs.append(("rotate", fill.get("rotWithShape")))
    crop = _edges(inline.find(f".//{{{PIC}}}blipFill/{{{A}}}srcRect"), ("l", "t", "r", "b"))
    if crop is not None:
        attrs.append(("crop", crop))
    lock = inline.find(f"./{{{WP}}}cNvGraphicFramePr/{{{A}}}graphicFrameLocks")
    if lock is None or lock.get("noChangeAspect") != "1":
        attrs.append(("lock", "0"))
    pic_lock = inline.find(f".//{{{PIC}}}cNvPicPr/{{{A}}}picLocks")
    if pic_lock is None:
        attrs.append(("piclock", "0"))
    else:
        names = sorted(etree.QName(key).localname for key, value in pic_lock.attrib.items() if value == "1")
        if names != ["noChangeAspect"]:
            attrs.append(("piclock", ",".join(names) or "none"))
    extras = _shape_extras(inline.find(f".//{{{PIC}}}spPr"))
    if extras:
        attrs.append(("shape", extras))
    rel_id = blip.get(f"{{{R}}}embed")
    rebuilt = create(dict(attrs), "rIdX", 1)
    if _signature(rebuilt) != _signature(drawing):
        return None
    return attrs, rel_id


def _signature(node: etree._Element) -> tuple:
    local = etree.QName(node).localname
    attrs = tuple(sorted(
        (key, value) for key, value in node.attrib.items()
        if key not in _IGNORED_ATTRS and not (local == "docPr" and key == "id")
        and not (local == "cNvPr" and key == "id" and value != "0")
    ))
    return (node.tag, attrs, (node.text or "").strip(), tuple(_signature(child) for child in node if isinstance(child.tag, str)))


def create(attrs: dict[str, str], rel_id: str, shape_id: int) -> etree._Element:
    """Build ``w:drawing`` for an inline picture from :func:`describe` attrs."""
    nsmap = {key: KNOWN_NAMESPACES[key] for key in ("w", "wp", "a", "pic", "r", "a14")}
    w = KNOWN_NAMESPACES["w"]
    drawing = etree.Element(f"{{{w}}}drawing", nsmap=nsmap)
    inline = etree.SubElement(drawing, f"{{{WP}}}inline")
    if attrs.get("dist") != "none":
        for name, value in zip(_DIST, (attrs.get("dist") or "0,0,0,0").split(",")):
            inline.set(name, value)
    cx, cy = attrs.get("cx", "0"), attrs.get("cy", "0")
    etree.SubElement(inline, f"{{{WP}}}extent", cx=cx, cy=cy)
    effect = attrs.get("effect") or "0,0,0,0"
    if effect != "none":
        edges = effect.split(",")
        etree.SubElement(inline, f"{{{WP}}}effectExtent", l=edges[0], t=edges[1], r=edges[2], b=edges[3])
    doc_pr = etree.SubElement(inline, f"{{{WP}}}docPr", id=str(shape_id), name=attrs.get("name", f"Picture {shape_id}"))
    if attrs.get("alt"):
        doc_pr.set("descr", attrs["alt"])
    if attrs.get("title"):
        doc_pr.set("title", attrs["title"])
    frame = etree.SubElement(inline, f"{{{WP}}}cNvGraphicFramePr")
    if attrs.get("lock", "1") != "0":
        etree.SubElement(frame, f"{{{A}}}graphicFrameLocks", noChangeAspect="1")
    graphic = etree.SubElement(inline, f"{{{A}}}graphic")
    data = etree.SubElement(graphic, f"{{{A}}}graphicData", uri="http://schemas.openxmlformats.org/drawingml/2006/picture")
    pic = etree.SubElement(data, f"{{{PIC}}}pic")
    nv = etree.SubElement(pic, f"{{{PIC}}}nvPicPr")
    pic_id = str(shape_id) if attrs.get("picid") == "1" else "0"
    etree.SubElement(nv, f"{{{PIC}}}cNvPr", id=pic_id, name=attrs.get("picname", attrs.get("name", f"Picture {shape_id}")))
    cnv = etree.SubElement(nv, f"{{{PIC}}}cNvPicPr")
    piclock = attrs.get("piclock", "noChangeAspect")
    if piclock != "0":
        locks = etree.SubElement(cnv, f"{{{A}}}picLocks")
        for name in piclock.split(","):
            if name and name != "none":
                locks.set(name, "1")
    fill = etree.SubElement(pic, f"{{{PIC}}}blipFill")
    if attrs.get("rotate"):
        fill.set("rotWithShape", attrs["rotate"])
    blip = etree.SubElement(fill, f"{{{A}}}blip")
    blip.set(f"{{{R}}}embed", rel_id)
    if attrs.get("cstate"):
        blip.set("cstate", attrs["cstate"])
    if "localdpi" in attrs:
        ext_list = etree.SubElement(blip, f"{{{A}}}extLst")
        ext = etree.SubElement(ext_list, f"{{{A}}}ext", uri=_LOCAL_DPI)
        etree.SubElement(ext, f"{{{A14}}}useLocalDpi", val=attrs["localdpi"])
    if attrs.get("crop"):
        edges = attrs["crop"].split(",")
        rect = etree.SubElement(fill, f"{{{A}}}srcRect")
        for name, value in zip(("l", "t", "r", "b"), edges):
            if value != "0":
                rect.set(name, value)
    stretch = etree.SubElement(fill, f"{{{A}}}stretch")
    etree.SubElement(stretch, f"{{{A}}}fillRect")
    flags = set((attrs.get("shape") or "").split(","))
    sp = etree.SubElement(pic, f"{{{PIC}}}spPr")
    for flag in flags:
        if flag.startswith("bw-"):
            sp.set("bwMode", flag[3:])
    xfrm = etree.SubElement(sp, f"{{{A}}}xfrm")
    etree.SubElement(xfrm, f"{{{A}}}off", x="0", y="0")
    etree.SubElement(xfrm, f"{{{A}}}ext", cx=cx, cy=cy)
    geom = etree.SubElement(sp, f"{{{A}}}prstGeom", prst="rect")
    if "noav" not in flags:
        etree.SubElement(geom, f"{{{A}}}avLst")
    if "nofill" in flags:
        etree.SubElement(sp, f"{{{A}}}noFill")
    if "noline" in flags:
        line = etree.SubElement(sp, f"{{{A}}}ln")
        etree.SubElement(line, f"{{{A}}}noFill")
    if "noshadow" in flags:
        ext_list = etree.SubElement(sp, f"{{{A}}}extLst")
        ext = etree.SubElement(ext_list, f"{{{A}}}ext", uri=_SHADOW)
        etree.SubElement(ext, f"{{{A14}}}shadowObscured")
    return drawing


__all__ = ["create", "describe"]
