"""DOCX-to-Markdown export built on the installed Pandoc command.

The exporter keeps Pandoc as the single conversion engine and only performs
small, deterministic post-processing that is needed for portable media links
and optional section files.
"""

from __future__ import annotations

import html
import hashlib
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from PIL import Image
from lxml import etree, html as lxml_html

from .._pandoc import run_pandoc
from ..output import validate_output_path

IMAGE_TAG_RE = re.compile(r"<img\b(?P<attrs>[^>]*?)(?:/?)>", re.IGNORECASE)
IMAGE_ATTR_RE = re.compile(r"\b(?P<name>src|alt)\s*=\s*(['\"])(?P<value>.*?)\2", re.IGNORECASE)
MARKDOWN_IMAGE_RE = re.compile(r"(?P<start>!\[[^\]]*\]\()(?P<src>[^)\s]+)(?P<end>\))")
INLINE_IMAGE_CAPTION_RE = re.compile(
    r"(?m)^(?P<image>!\[[^\]]*\]\([^\n)]+\))(?P<caption>\s*(?:\*{0,3})?(?:Fig\.|Supplementary Fig\.|Table\b))"
)
BOLD_HEADING_RE = re.compile(r"^\*\*(?P<text>.+?)\*\*\s*:?[ \t]*$")
ATX_HEADING_RE = re.compile(r"^(?P<hashes>#{1,6})\s+(?P<text>.+?)\s*$")
SUBSCRIPT_RE = re.compile(r"<sub>(?P<value>[^<\n]*?)</sub>", re.IGNORECASE)
SUBSCRIPT_CLOSING_DELIMITER_RE = re.compile(r"(\\\]|[)\]])")
HTML_TABLE_RE = re.compile(r"<table\b[^>]*>.*?</table>", re.IGNORECASE | re.DOTALL)


@dataclass(frozen=True)
class MarkdownExportResult:
    """Paths and provenance for one Markdown export."""

    output: Path | None
    split_dir: Path | None
    command: tuple[str, ...]
    media: tuple[Path, ...]
    sections: tuple[Path, ...]


def _image_source(value: str, media_root: Path) -> Path | None:
    value = html.unescape(value).strip().strip("'\"")
    if not value or value.startswith(("http://", "https://", "data:")):
        return None
    candidate = Path(value.replace("/", os.sep).replace("\\", os.sep))
    if candidate.exists() and candidate.is_file():
        return candidate
    by_name = media_root / candidate.name
    if by_name.exists() and by_name.is_file():
        return by_name
    for path in media_root.rglob(candidate.name):
        if path.is_file():
            return path
    return None


def _convert_tiff(path: Path) -> Path:
    if path.suffix.lower() not in {".tif", ".tiff"}:
        return path
    target = path.with_suffix(".png")
    if not target.exists():
        with Image.open(path) as image:
            image.save(target, format="PNG")
    return target


def _relative_media(value: str, *, document_dir: Path, media_root: Path) -> tuple[str, Path | None]:
    source = _image_source(value, media_root)
    if source is None:
        return value, None
    rendered = _convert_tiff(source)
    relative = os.path.relpath(rendered, document_dir).replace(os.sep, "/")
    return relative, rendered


def normalize_media_links(text: str, *, document_dir: Path, media_root: Path) -> tuple[str, tuple[Path, ...]]:
    """Make Pandoc's absolute image links portable and convert HTML images.

    Pandoc emits HTML ``img`` tags when it needs to retain image dimensions.
    The standard Markdown image form is easier to consume by docforge's
    Markdown renderer, so dimensions are intentionally represented by the
    media asset rather than embedded HTML attributes.
    """

    media: set[Path] = set()

    def replace_html(match: re.Match[str]) -> str:
        attrs = match.group("attrs")
        parsed = {item.group("name").lower(): item.group("value") for item in IMAGE_ATTR_RE.finditer(attrs)}
        if "src" not in parsed:
            return match.group(0)
        relative, source = _relative_media(parsed["src"], document_dir=document_dir, media_root=media_root)
        if source is None:
            return match.group(0)
        media.add(source)
        alt = parsed.get("alt") or source.stem
        return f"![{alt}]({relative})"

    text = IMAGE_TAG_RE.sub(replace_html, text)

    def replace_markdown(match: re.Match[str]) -> str:
        relative, source = _relative_media(match.group("src"), document_dir=document_dir, media_root=media_root)
        if source is None:
            return match.group(0)
        media.add(source)
        return f"{match.group('start')}{relative}{match.group('end')}"

    text = MARKDOWN_IMAGE_RE.sub(replace_markdown, text)
    # Pandoc can place a caption immediately after an inline image when the
    # original Word paragraph contained both.  Split that pair into the
    # block forms consumed by the Markdown parser; otherwise the image is
    # mistaken for an ordinary paragraph and the following figure shifts.
    text = INLINE_IMAGE_CAPTION_RE.sub(
        lambda match: f"{match.group('image')}\n\n{match.group('caption').lstrip()}",
        text,
    )
    return text, tuple(sorted(media))


def normalize_script_boundaries(text: str) -> str:
    """Move baseline formula delimiters out of Pandoc subscript spans.

    Word often stores a closing ``)`` or ``]`` in the same subscript run as
    the preceding digit. Pandoc preserves that run boundary literally, for
    example ``N<sub>3)</sub>``. Delimiters are baseline punctuation in chemical
    formulae, so canonicalize them before Markdown is written.
    """

    def normalize_subscript(match: re.Match[str]) -> str:
        value = match.group("value")
        if not SUBSCRIPT_CLOSING_DELIMITER_RE.search(value):
            return match.group(0)
        pieces = SUBSCRIPT_CLOSING_DELIMITER_RE.split(value)
        return "".join(
            piece if SUBSCRIPT_CLOSING_DELIMITER_RE.fullmatch(piece) else f"<sub>{piece}</sub>"
            for piece in pieces
            if piece
        )

    return SUBSCRIPT_RE.sub(normalize_subscript, text)


def _table_cell_markdown(cell) -> str:
    """Serialize one Pandoc HTML table cell for the Markdown block parser.

    Pandoc keeps tables containing row/column spans as raw HTML.  That is a
    valid GFM document, but it cannot be consumed by the small, auditable
    block parser used by ``md2docx``.  Keep semantic inline markup while
    removing layout-only HTML so the resulting pipe table is portable.
    """

    def render(node) -> str:
        pieces: list[str] = []
        if node.text:
            pieces.append(node.text)
        for child in node:
            tag = etree.QName(child).localname.lower()
            value = render(child)
            if tag in {"em", "i"}:
                value = f"*{value}*"
            elif tag in {"strong", "b"}:
                value = f"**{value}**"
            elif tag in {"sup", "sub"}:
                value = f"<{tag}>{value}</{tag}>"
            elif tag == "br":
                value = "<br>"
            elif tag == "p":
                value = f"{value}<br>"
            # ``p`` and other layout tags contribute their text only.
            pieces.append(value)
            if child.tail:
                pieces.append(child.tail)
        return "".join(pieces)

    value = render(cell)
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    value = re.sub(r"(?:<br>\s*)+$", "", value, flags=re.IGNORECASE)
    # A literal pipe would change the table shape.  Markdown backslash
    # escaping is understood by the parser and survives Word export.
    return value.replace("|", r"\|")


def _html_table_to_pipe(fragment: str) -> str:
    """Convert one raw Pandoc HTML table to a rectangular pipe table."""

    root = lxml_html.fragment_fromstring(fragment, create_parent=True)
    table = root.find(".//table")
    if table is None:
        return fragment

    rows: list[list[str]] = []
    active: dict[int, tuple[str, int]] = {}
    for tr in table.xpath(".//tr"):
        values: list[str] = []
        col = 0

        def carry() -> None:
            nonlocal col
            while col in active:
                value, remaining = active[col]
                values.append(value)
                if remaining <= 1:
                    del active[col]
                else:
                    active[col] = (value, remaining - 1)
                col += 1

        carry()
        for cell in tr.xpath("./th|./td"):
            carry()
            value = _table_cell_markdown(cell)
            colspan = max(1, int(cell.get("colspan", "1")))
            rowspan = max(1, int(cell.get("rowspan", "1")))
            for offset in range(colspan):
                values.append(value)
                if rowspan > 1:
                    active[col + offset] = (value, rowspan - 1)
            col += colspan
        carry()
        rows.append(values)

    if not rows:
        return ""
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = [
        "| " + " | ".join(rows[0]) + " |",
        "| " + " | ".join("---" for _ in range(width)) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(lines)


def normalize_html_tables(text: str) -> str:
    """Convert Pandoc's raw HTML tables to parser-compatible pipe tables."""

    return HTML_TABLE_RE.sub(lambda match: _html_table_to_pipe(match.group(0)), text)


def normalize_inline_html(text: str) -> str:
    """Map harmless HTML emphasis/layout tags to Markdown equivalents."""

    text = re.sub(r"<\s*(?:em|i)\s*>", "*", text, flags=re.IGNORECASE)
    text = re.sub(r"<\s*/\s*(?:em|i)\s*>", "*", text, flags=re.IGNORECASE)
    text = re.sub(r"<\s*(?:strong|b)\s*>", "**", text, flags=re.IGNORECASE)
    text = re.sub(r"<\s*/\s*(?:strong|b)\s*>", "**", text, flags=re.IGNORECASE)
    # Paragraph tags occur inside table cells after Pandoc export.  They are
    # layout markers, not manuscript text.
    text = re.sub(r"<\s*/?\s*p\s*>", " ", text, flags=re.IGNORECASE)
    return text


def _heading_text(line: str) -> str | None:
    stripped = line.strip()
    match = ATX_HEADING_RE.match(stripped)
    if match:
        return match.group("text").strip().rstrip(":").strip()
    match = BOLD_HEADING_RE.match(stripped)
    if match:
        return html.unescape(match.group("text")).strip().rstrip(":").strip()
    return None


def _normalise_heading_line(line: str, level: int) -> str:
    stripped = line.strip()
    match = _heading_text(stripped)
    if match is None:
        return line
    return f"{'#' * level} {match}"


def _normalise_section_headings(lines: list[str], *, first_level: int = 1) -> list[str]:
    result: list[str] = []
    first = True
    for line in lines:
        heading = _heading_text(line)
        if heading is None:
            result.append(line)
            continue
        level = first_level if first else min(first_level + 1, 6)
        result.append(_normalise_heading_line(line, level))
        first = False
    return result


def _plain_inline(value: str) -> str:
    value = value.replace(r"\*", "*")
    value = re.sub(r"<\/?(?:sup|sub)>", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\*\*|__|(?<!\w)\*(?!\s)|(?<!\s)\*(?!\w)", "", value)
    if value.startswith("\\"):
        value = value[1:]
    return html.unescape(value).strip()


def metadata_markdown(front_matter: str) -> str:
    """Turn the pre-Abstract paragraphs into template-compatible metadata."""

    blocks = [block.strip() for block in re.split(r"\n\s*\n", front_matter) if block.strip()]
    if not blocks:
        return "# METADATA\n"
    title = _plain_inline(blocks[0])
    author = _plain_inline(blocks[1]) if len(blocks) > 1 else ""
    affiliations: list[str] = []
    contacts: list[str] = []
    for block in blocks[2:]:
        if re.search(r"correspondence", block, re.IGNORECASE):
            contacts.append(_plain_inline(block))
        else:
            affiliations.append(_plain_inline(block))
    parts = ["# TITLE", "", title, "", "# AUTHOR", "", author]
    if affiliations:
        parts.extend(["", "# AFFILIATION", "", "\n\n".join(affiliations)])
    if contacts:
        parts.extend(["", "# EMAILS", "", "\n".join(contacts)])
    return "\n".join(parts).rstrip() + "\n"


def _ref_value(value: Any) -> tuple[str, int]:
    if isinstance(value, str):
        return value, 1
    if not isinstance(value, Mapping) or "heading" not in value:
        raise ValueError("section map heading references must be a string or {heading, occurrence}")
    occurrence = int(value.get("occurrence", 1))
    if occurrence < 1:
        raise ValueError("section map occurrence must be >= 1")
    return str(value["heading"]), occurrence


def load_section_map(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    sections = payload.get("sections") if isinstance(payload, Mapping) else None
    if not isinstance(sections, list) or not sections:
        raise ValueError("section map must contain a non-empty 'sections' list")
    result: list[dict[str, Any]] = []
    for item in sections:
        if not isinstance(item, Mapping) or not item.get("file"):
            raise ValueError("each section map entry requires a file")
        result.append(dict(item))
    return result


def split_markdown_sections(text: str, section_map: list[dict[str, Any]]) -> list[tuple[str, str]]:
    """Split Markdown using ordered heading references from a section map."""

    lines = text.splitlines()
    markers: list[tuple[str, int]] = []
    for index, line in enumerate(lines):
        heading = _heading_text(line)
        if heading:
            markers.append((heading, index))

    def resolve(reference: Any, *, default: int | None = None) -> int | None:
        if reference is None:
            return default
        heading, occurrence = _ref_value(reference)
        wanted = heading.rstrip(":").strip().casefold()
        seen = 0
        for value, index in markers:
            if value.rstrip(":").strip().casefold() != wanted:
                continue
            seen += 1
            if seen == occurrence:
                return index
        raise ValueError(f"section heading not found: {heading!r} occurrence {occurrence}")

    result: list[tuple[str, str]] = []
    previous_end = 0
    for item in section_map:
        start = resolve(item.get("start"), default=0)
        end = resolve(item.get("end"), default=len(lines))
        assert start is not None and end is not None
        if start < previous_end or end <= start:
            raise ValueError(f"section map ranges overlap or are empty: {item['file']}")
        chunk = lines[start:end]
        if item.get("kind") == "metadata" or item["file"].startswith("00_"):
            content = metadata_markdown("\n".join(lines[:end]))
        else:
            content = "\n".join(_normalise_section_headings(chunk)).strip() + "\n"
        result.append((str(item["file"]), content))
        previous_end = end
    return result


def write_conversion_manifest(path: Path, payload: Mapping[str, Any]) -> Path:
    path.write_text(json.dumps(dict(payload), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _roundtrip_source_path(split_dir: Path) -> Path:
    """Return the portable source snapshot path used by a Markdown bundle."""

    return split_dir / "roundtrip" / "source.docx"


def reuse_unchanged_roundtrip_source(
    manifest_path: Path,
    output: Path,
    *,
    force: bool = False,
) -> bool:
    """Reuse the source DOCX when a Markdown bundle is byte-for-byte unchanged.

    A Markdown renderer cannot encode every OOXML detail (for example custom
    run properties, theme colors, relationship ids, or section controls).  A
    conversion manifest therefore carries a portable source snapshot and the
    SHA-256 values of every editable source.  When all values still match, the
    exact source package is copied to the requested output.  If anything
    changed, ``False`` is returned and the caller can continue through the
    normal Markdown renderer.
    """

    manifest_path = manifest_path.resolve()
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    if payload.get("schema") != "docforge.docx2md.v1":
        raise ValueError("roundtrip manifest must use schema docforge.docx2md.v1")
    root = manifest_path.parent
    source_name = payload.get("source_copy")
    source_sha = payload.get("source_copy_sha256")
    if not isinstance(source_name, str) or not isinstance(source_sha, str):
        raise ValueError("roundtrip manifest does not contain a source DOCX snapshot")
    if Path(source_name).is_absolute():
        raise ValueError("roundtrip source snapshot must use a bundle-relative path")
    source = (root / source_name).resolve()
    try:
        source.relative_to(root.resolve())
    except ValueError as exc:
        raise ValueError("roundtrip source snapshot escapes the bundle") from exc
    if not source.is_file() or _sha256(source) != source_sha:
        raise ValueError("roundtrip source snapshot is missing or its SHA-256 does not match")

    for name, expected in dict(payload.get("section_sha256", {})).items():
        if Path(str(name)).is_absolute():
            raise ValueError("roundtrip section path must be bundle-relative")
        path = (root / str(name)).resolve()
        try:
            path.relative_to(root.resolve())
        except ValueError as exc:
            raise ValueError("roundtrip section path escapes the bundle") from exc
        if not path.is_file() or _sha256(path) != expected:
            return False
    for name, expected in dict(payload.get("media_sha256", {})).items():
        if Path(str(name)).is_absolute():
            raise ValueError("roundtrip media path must be bundle-relative")
        path = (root / str(name)).resolve()
        try:
            path.relative_to(root.resolve())
        except ValueError as exc:
            raise ValueError("roundtrip media path escapes the bundle") from exc
        if not path.is_file() or _sha256(path) != expected:
            return False

    validate_output_path(output)
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, output)
    return True


def docx_to_markdown(
    input_path: Path,
    *,
    output: Path | None = None,
    split_dir: Path | None = None,
    section_map_path: Path | None = None,
    media_dir: Path | None = None,
    track_changes: str = "accept",
    force: bool = False,
) -> MarkdownExportResult:
    """Convert a DOCX to GFM, optionally emitting mapped section files."""

    validate_output_path(output)
    validate_output_path(split_dir, label="split output directory")
    if not input_path.exists():
        raise FileNotFoundError(input_path)
    if output is None and split_dir is None:
        raise ValueError("docx2md requires --output, --split-dir, or both")
    if split_dir is not None and section_map_path is None:
        raise ValueError("--split-dir requires --section-map")
    if output is not None and output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    if split_dir is not None:
        split_dir.mkdir(parents=True, exist_ok=True)
    destination_dir = output.parent if output is not None else split_dir
    assert destination_dir is not None
    destination_dir.mkdir(parents=True, exist_ok=True)
    extraction_root = media_dir or destination_dir
    extraction_root.mkdir(parents=True, exist_ok=True)
    media_root = extraction_root / "media"

    with tempfile.TemporaryDirectory(prefix="docforge-docx2md-") as temporary:
        rendered_path = output or Path(temporary) / "document.md"
        command = run_pandoc(
            input_path,
            rendered_path,
            output_format="gfm",
            extract_media_root=extraction_root,
            track_changes=track_changes,
        )
        markdown, media = normalize_media_links(
            rendered_path.read_text(encoding="utf-8"),
            document_dir=destination_dir,
            media_root=media_root,
        )
        markdown = normalize_script_boundaries(markdown)
        markdown = normalize_html_tables(markdown)
        markdown = normalize_inline_html(markdown)
        # Word occasionally contains an empty bold marker paragraph.  Pandoc
        # serializes that artifact as ``**\\**``; it is not manuscript text.
        markdown = re.sub(r"(?m)^\*\*\\\*\*\s*\r?\n?", "", markdown)
        if output is not None:
            output.write_text(markdown, encoding="utf-8")
        sections: list[Path] = []
        if split_dir is not None:
            source_copy = _roundtrip_source_path(split_dir)
            source_copy.parent.mkdir(parents=True, exist_ok=True)
            if source_copy.exists() and not force:
                raise FileExistsError(f"Output exists; pass --force to overwrite: {source_copy}")
            shutil.copyfile(input_path, source_copy)
            section_map = load_section_map(section_map_path)  # type: ignore[arg-type]
            for filename, content in split_markdown_sections(markdown, section_map):
                path = (split_dir / filename).resolve()
                try:
                    path.relative_to(split_dir.resolve())
                except ValueError as exc:
                    raise ValueError(f"section path escapes split directory: {filename}") from exc
                if path.exists() and not force:
                    raise FileExistsError(f"Output exists; pass --force to overwrite: {path}")
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(content, encoding="utf-8")
                sections.append(path)
            manifest = {
                "schema": "docforge.docx2md.v1",
                "source": str(input_path),
                "source_sha256": _sha256(input_path),
                "source_copy": os.path.relpath(source_copy, split_dir).replace(os.sep, "/"),
                "source_copy_sha256": _sha256(source_copy),
                "track_changes": track_changes,
                "command": list(command),
                "sections": [str(path.name) for path in sections],
                "section_sha256": {str(path.name): _sha256(path) for path in sections},
                "media": [
                    os.path.relpath(path, split_dir).replace(os.sep, "/")
                    for path in media
                    if path.exists()
                ],
                "media_sha256": {
                    os.path.relpath(path, split_dir).replace(os.sep, "/"): _sha256(path)
                    for path in media
                    if path.exists()
                },
            }
            manifest_path = split_dir / "manifest.json"
            if manifest_path.exists() and not force:
                raise FileExistsError(f"Output exists; pass --force to overwrite: {manifest_path}")
            write_conversion_manifest(manifest_path, manifest)
    return MarkdownExportResult(output, split_dir, command, media, tuple(sections))


__all__ = [
    "MarkdownExportResult",
    "docx_to_markdown",
    "load_section_map",
    "metadata_markdown",
    "normalize_media_links",
    "normalize_script_boundaries",
    "normalize_html_tables",
    "normalize_inline_html",
    "split_markdown_sections",
    "write_conversion_manifest",
    "reuse_unchanged_roundtrip_source",
]
