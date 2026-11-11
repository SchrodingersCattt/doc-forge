"""Write an editable PowerPoint package from title and bullet records.

The deck is an Office Open XML zip. Each slide stores its title in one
text box and its bullets as paragraphs in a second text box. Both boxes
have explicit positions and black text, so a reader can show them without
a theme or a consumer template. Slide text is escaped and never becomes a
package part name.
"""

from __future__ import annotations

import errno
import io
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any, Mapping, Sequence
from xml.sax.saxutils import escape

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_SLIDE_CX = 12_192_000
_SLIDE_CY = 6_858_000
_TITLE_BOX = (457_200, 365_760, 11_277_600, 1_143_000)
_BODY_BOX = (457_200, 1_600_200, 11_277_600, 4_800_600)

# XML 1.0 Char: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF].
_XML_DISALLOWED = re.compile(r"[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]")

_CONTENT_TYPE = {
    "ppt/presentation.xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "ppt/presProps.xml": "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
    "ppt/viewProps.xml": "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
    "ppt/tableStyles.xml": "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
    "ppt/slideMasters/slideMaster1.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "ppt/slideLayouts/slideLayout1.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "ppt/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
    "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
    "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
}
_SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"


def write_pptx(
    path: str | Path,
    slides: Sequence[Mapping[str, Any]],
    *,
    force: bool = False,
) -> None:
    """Write *slides* to a ``.pptx`` at *path*.

    Each slide mapping needs ``title`` (a non-empty string) and ``bullets``
    (a list of strings). Characters XML 1.0 cannot store are refused before
    the destination is opened. An existing file is left unchanged unless
    *force* is true. A directory is never replaced. The parent directory
    must already exist.
    """

    records = _normalize(slides)
    destination = Path(path)
    _require_output_destination(destination)
    if destination.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {destination}")
    _publish(destination, _package(records))


def load_slides(path: str | Path) -> list[Any]:
    """Read a JSON list of ``{"title", "bullets"}`` objects."""

    source = Path(path)
    try:
        text = source.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("slides file must be UTF-8 JSON") from exc
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"slides file is not valid JSON ({exc.msg})") from exc
    if not isinstance(payload, list):
        raise ValueError("slides file must be a JSON list")
    return payload


def _normalize(slides: Sequence[Mapping[str, Any]]) -> list[tuple[str, tuple[str, ...]]]:
    if isinstance(slides, (str, bytes)) or not isinstance(slides, Sequence):
        raise ValueError("slides must be a sequence of mappings")
    records: list[tuple[str, tuple[str, ...]]] = []
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, Mapping):
            raise ValueError(f"slide {index} must be a mapping")
        if "title" not in slide:
            raise ValueError(f"slide {index} is missing a title")
        title = slide["title"]
        if not isinstance(title, str):
            raise ValueError(f"slide {index} title must be a string")
        _xml_text(title)
        if title.strip() == "":
            raise ValueError(f"slide {index} title must not be empty")
        if "bullets" not in slide:
            raise ValueError(f"slide {index} is missing bullets")
        bullets = slide["bullets"]
        if isinstance(bullets, (str, bytes)) or not isinstance(bullets, Sequence):
            raise ValueError(f"slide {index} bullets must be a list of strings")
        if not all(isinstance(item, str) for item in bullets):
            raise ValueError(f"slide {index} bullets must be a list of strings")
        for bullet in bullets:
            _xml_text(bullet)
        records.append((title, tuple(bullets)))
    return records


def _xml_text(value: str) -> str:
    if _XML_DISALLOWED.search(value):
        raise ValueError("slide text contains a character XML cannot store")
    return escape(value, {'"': "&quot;", "'": "&apos;"})


def _output_error(path: Path, reason: str) -> ValueError:
    """Filesystem failure for *path* that does not name a temporary file."""

    return ValueError(f"Cannot write {path}: {reason}")


def _require_output_destination(path: Path) -> None:
    """Refuse a missing parent or a directory at the requested path."""

    parent = path.parent
    try:
        parent_is_dir = parent.is_dir()
        parent_exists = parent.exists()
        destination_is_dir = path.is_dir()
    except OSError as exc:
        raise _output_error(path, exc.strerror or "unable to write output") from exc
    if not parent_is_dir:
        reason = "Not a directory" if parent_exists else "No such file or directory"
        raise _output_error(path, reason)
    if destination_is_dir:
        raise _output_error(path, "Is a directory")


def _part_name(name: str) -> str:
    """Return a zip member name that stays inside the package."""

    if not name or "\\" in name or ":" in name:
        raise ValueError(f"package part is not inside the deck: {name}")
    path = PurePosixPath(name)
    if path.is_absolute() or any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"package part is not inside the deck: {name}")
    return "/".join(path.parts)


def _relationship_target(target: str) -> str:
    """Return a package-absolute relationship target with no parent segments."""

    if not target.startswith("/") or "\\" in target or ":" in target:
        raise ValueError(f"relationship target escapes the deck: {target}")
    path = PurePosixPath(target)
    if any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"relationship target escapes the deck: {target}")
    return "/" + "/".join(part for part in path.parts if part != "/")


def _publish(path: Path, parts: Mapping[str, str]) -> None:
    """Publish a complete package without reusing a predictable temp name.

    The zip is built in memory first, so an encoding failure does not create
    a file. ``mkstemp`` creates a new inode in the destination directory.
    Bytes are written to that file descriptor, not reopened by name, and
    ``os.replace`` publishes the same inode. An existing ``.{name}.partial``
    sibling, including a symlink, is not opened.
    """

    payload = _zip_bytes(parts)
    temporary: Path | None = None
    descriptor: int | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{path.name}.",
            suffix=".partial",
            dir=path.parent,
        )
        temporary = Path(temporary_name)
        _write_fd(descriptor, payload)
        os.fchmod(descriptor, _creation_mode())
        expected = os.fstat(descriptor)
        current = os.lstat(temporary)
        if (current.st_ino, current.st_dev) != (expected.st_ino, expected.st_dev):
            raise OSError(errno.EIO, "temporary output changed before publish")
        os.close(descriptor)
        descriptor = None
        os.replace(temporary, path)
    except OSError as exc:
        raise _output_error(path, exc.strerror or "unable to write output") from exc
    finally:
        if descriptor is not None:
            try:
                os.close(descriptor)
            except OSError:
                pass
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass


def _creation_mode() -> int:
    """Match a normal ``open`` so ``mkstemp`` does not leave the deck owner-only."""

    current = os.umask(0)
    try:
        return 0o666 & ~current
    finally:
        os.umask(current)


def _write_fd(descriptor: int, payload: bytes) -> None:
    view = memoryview(payload)
    while view:
        written = os.write(descriptor, view)
        if written <= 0:
            raise OSError(errno.EIO, "unable to write output")
        view = view[written:]
    os.fsync(descriptor)


def _zip_bytes(parts: Mapping[str, str]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(_part_name(name), payload.encode("utf-8"))
    return buffer.getvalue()


def _package(slides: Sequence[tuple[str, tuple[str, ...]]]) -> dict[str, str]:
    parts: dict[str, str] = {
        "ppt/presentation.xml": _presentation(len(slides)),
        "ppt/_rels/presentation.xml.rels": _presentation_rels(len(slides)),
        "ppt/presProps.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr xmlns:p="{_P}"/>',
        "ppt/viewProps.xml": _VIEW_PROPS,
        "ppt/tableStyles.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:tblStyleLst xmlns:a="{_A}"/>',
        "ppt/slideMasters/slideMaster1.xml": _SLIDE_MASTER,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": _rels((
            ("rId1", f"{_R}/slideLayout", "/ppt/slideLayouts/slideLayout1.xml"),
            ("rId2", f"{_R}/theme", "/ppt/theme/theme1.xml"),
        )),
        "ppt/slideLayouts/slideLayout1.xml": _SLIDE_LAYOUT,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": _rels((
            ("rId1", f"{_R}/slideMaster", "/ppt/slideMasters/slideMaster1.xml"),
        )),
        "ppt/theme/theme1.xml": _THEME,
        "docProps/core.xml": _CORE,
        "docProps/app.xml": _app(len(slides)),
        "_rels/.rels": _rels((
            ("rId1", f"{_R}/officeDocument", "/ppt/presentation.xml"),
            ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "/docProps/core.xml"),
            ("rId3", f"{_R}/extended-properties", "/docProps/app.xml"),
        )),
    }
    content_types = dict(_CONTENT_TYPE)
    for index, (title, bullets) in enumerate(slides, start=1):
        slide_name = f"ppt/slides/slide{index}.xml"
        parts[slide_name] = _slide_xml(title, bullets)
        parts[f"ppt/slides/_rels/slide{index}.xml.rels"] = _rels((
            ("rId1", f"{_R}/slideLayout", "/ppt/slideLayouts/slideLayout1.xml"),
        ))
        content_types[slide_name] = _SLIDE_CONTENT_TYPE
    parts["[Content_Types].xml"] = _content_types(content_types)
    return parts


def _presentation(count: int) -> str:
    slide_ids = "".join(
        f'<p:sldId id="{255 + index}" r:id="rId{index + 4}"/>'
        for index in range(1, count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p="{_P}" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f"<p:sldIdLst>{slide_ids}</p:sldIdLst>"
        f'<p:sldSz cx="{_SLIDE_CX}" cy="{_SLIDE_CY}" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        "</p:presentation>"
    )


def _presentation_rels(count: int) -> str:
    items = [
        ("rId1", f"{_R}/slideMaster", "/ppt/slideMasters/slideMaster1.xml"),
        ("rId2", f"{_R}/presProps", "/ppt/presProps.xml"),
        ("rId3", f"{_R}/viewProps", "/ppt/viewProps.xml"),
        ("rId4", f"{_R}/tableStyles", "/ppt/tableStyles.xml"),
    ]
    items.extend(
        (f"rId{index + 4}", f"{_R}/slide", f"/ppt/slides/slide{index}.xml")
        for index in range(1, count + 1)
    )
    return _rels(items)


def _rels(items: Sequence[tuple[str, str, str]]) -> str:
    body = "".join(
        f'<Relationship Id="{rid}" Type="{rel_type}" Target="{_relationship_target(target)}"/>'
        for rid, rel_type, target in items
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL}">{body}</Relationships>'
    )


def _content_types(parts: Mapping[str, str]) -> str:
    overrides = "".join(
        f'<Override PartName="/{_part_name(name)}" ContentType="{content_type}"/>'
        for name, content_type in parts.items()
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{_CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{overrides}</Types>"
    )


def _slide_xml(title: str, bullets: Sequence[str]) -> str:
    shapes = [_shape(2, "Title", _TITLE_BOX, _run(title, size=3200, bold=True))]
    if bullets:
        paragraphs = "".join(_bullet(bullet) for bullet in bullets)
        shapes.append(_shape(3, "Body", _BODY_BOX, paragraphs))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p="{_P}">'
        "<p:cSld><p:bg><p:bgPr>"
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/>'
        "</p:bgPr></p:bg><p:spTree>"
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        "<p:grpSpPr/>"
        f"{''.join(shapes)}</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


def _shape(shape_id: int, name: str, box: tuple[int, int, int, int], paragraphs: str) -> str:
    x, y, cx, cy = box
    return (
        "<p:sp><p:nvSpPr>"
        f'<p:cNvPr id="{shape_id}" name="{name}"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>'
        f'<a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
        "</a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        '<p:txBody><a:bodyPr wrap="square" anchor="t"/>'
        f"<a:lstStyle/>{paragraphs}</p:txBody></p:sp>"
    )


def _run(text: str, *, size: int, bold: bool = False) -> str:
    bold_attr = ' b="1"' if bold else ""
    return (
        f'<a:p><a:r><a:rPr lang="en-US" sz="{size}"{bold_attr} dirty="0">'
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        '<a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/>'
        f'</a:rPr><a:t xml:space="preserve">{_xml_text(text)}</a:t></a:r></a:p>'
    )


def _bullet(text: str) -> str:
    return (
        '<a:p><a:pPr marL="342900" indent="-342900">'
        '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>'
        '<a:r><a:rPr lang="en-US" sz="2000" dirty="0">'
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        '<a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/>'
        f'</a:rPr><a:t xml:space="preserve">{_xml_text(text)}</a:t></a:r></a:p>'
    )


def _app(count: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>docforge</Application>"
        "<PresentationFormat>On-screen Show (16:9)</PresentationFormat>"
        f"<Slides>{count}</Slides></Properties>"
    )


_VIEW_PROPS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="{_A}" xmlns:p="{_P}" lastView="sldView">
  <p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>
  <p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale>
    <a:sx n="100" d="100"/><a:sy n="100" d="100"/>
  </p:scale><p:origin x="0" y="0"/></p:cViewPr></p:cSldViewPr></p:slideViewPr>
</p:viewPr>
"""

_SLIDE_MASTER = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p="{_P}">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>
"""

_SLIDE_LAYOUT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p="{_P}" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"""

_CORE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>docforge</dc:creator>
  <cp:lastModifiedBy>docforge</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>
"""

_THEME = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="{_A}" name="Plain">
  <a:themeElements>
    <a:clrScheme name="Plain">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F4E79"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D08E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Plain">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Plain">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
"""


__all__ = ["load_slides", "write_pptx"]
