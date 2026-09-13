"""Template-aware Markdown to DOCX assembly.

This module is project agnostic. A caller supplies Markdown paths, a DOCX
template, optional metadata, and an ordered JSON bibliography.
"""

from __future__ import annotations

import copy
import hashlib
import json
import re
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from pathlib import PurePosixPath
from typing import Iterable, Mapping, Sequence

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from PIL import Image

from .blocks import Block
from .launcher import (
    accept_docx_revisions,
    add_inline,
    convert_unicode_scripts_in_docx,
    normalize_document_typography,
    parse_markdown,
    remap_image_relationships,
    remove_docx_comments,
    render_blocks_to_doc,
    verify_clean_review_state,
    verify_image_relationships,
    verify_navigation_headings,
)

__all__ = [
    "AssemblyResult",
    "ManuscriptMetadata",
    "assemble_markdown_template",
    "discover_template_styles",
    "_format_bibliography_record",
    "load_bibliography",
    "parse_metadata",
    "sha256_file",
    "verify_template_output",
    "write_assembly_sidecars",
]

CITATION_RE = re.compile(r"\\citep?\{([^{}]+)\}")
HEADING_RE = re.compile(r"^#\s+(.+?)\s*$")
FORBIDDEN_TOKENS = (
    "[placeholder text]",
    "[title]",
    "[authors]",
    "[abstract]",
    "[section]",
    "[subsection]",
    "[subsubsection]",
    "[citation field]",
    "[acknowledgment]",
)
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass(frozen=True)
class ManuscriptMetadata:
    title: str = ""
    authors: str = ""
    affiliations: str = ""
    contacts: tuple[str, ...] = ()
    acknowledgement: str = ""
    author_contributions: str = ""
    code_availability: str = ""


@dataclass(frozen=True)
class AssemblyResult:
    output: Path
    title: str
    citation_map: Mapping[str, str | int]
    used_citations: tuple[str, ...]
    reference_keys: tuple[str, ...]
    style_map: Mapping[str, str]
    template_sections_before: int
    template_sections_after: int
    section_geometry: tuple[tuple[int, int, int, int, int, int], ...]
    body_columns_before: int
    body_columns_after: int
    column_layout: str
    verification: Mapping[str, object]
    font_family: str | None = None
    east_asia_font: str | None = None
    style_profile: str = "template"
    line_numbers: str = "template"
    contacts: tuple[str, ...] = ()
    include_title: bool = True
    strip_level_one_headings: bool = False
    heading_before: float | None = None
    section_sources: tuple[int, ...] = ()
    figures: tuple[Mapping[str, object], ...] = ()
    tables: tuple[Mapping[str, object], ...] = ()
    numbering_prefix: str = ""
    bibliography_scope: str = "all"
    include_metadata_back_matter: bool = True


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _metadata_key(label: str) -> str | None:
    label = re.sub(r"[\s:_-]+", " ", label.strip().lower()).strip()
    return {
        "title": "title",
        "author": "authors",
        "authors": "authors",
        "affiliation": "affiliations",
        "affiliations": "affiliations",
        "acknowledgement": "acknowledgement",
        "acknowledgment": "acknowledgement",
        "acknowledgements": "acknowledgement",
        "acknowledgments": "acknowledgement",
        "author contribution": "author_contributions",
        "author contributions": "author_contributions",
        "author contributions statement": "author_contributions",
        "email": "contacts",
        "emails": "contacts",
        "e mail": "contacts",
        "corresponding email": "contacts",
        "corresponding emails": "contacts",
        "code availability": "code_availability",
    }.get(label)


def parse_metadata(path: Path) -> ManuscriptMetadata:
    sections: dict[str, list[str]] = {}
    current: str | None = None
    for line in path.read_text(encoding="utf-8-sig").replace("\r\n", "\n").splitlines():
        match = HEADING_RE.match(line.strip())
        if match:
            current = _metadata_key(match.group(1))
            continue
        if current is not None:
            sections.setdefault(current, []).append(line.strip())
    raw_values = {
        key: [part for part in parts if part]
        for key, parts in sections.items()
    }
    author_lines = raw_values.get("authors", [])
    contact_lines = list(raw_values.get("contacts", []))
    retained_authors = []
    for line in author_lines:
        if "@" in line and line.lstrip("*").strip():
            contact_lines.append(line)
        else:
            retained_authors.append(line)
    if len(contact_lines) > 2:
        raise ValueError("Metadata supports at most two contact email lines")
    return ManuscriptMetadata(
        title=" ".join(raw_values.get("title", [])).strip(),
        authors=" ".join(retained_authors).strip(),
        affiliations=" ".join(raw_values.get("affiliations", [])).strip(),
        contacts=tuple(contact_lines),
        acknowledgement=" ".join(raw_values.get("acknowledgement", [])).strip(),
        author_contributions=" ".join(raw_values.get("author_contributions", [])).strip(),
        code_availability=" ".join(raw_values.get("code_availability", [])).strip(),
    )


def _is_filled_metadata(value: str) -> bool:
    """Treat unresolved optional metadata markers as absent during assembly."""
    return value.strip().upper() not in {"", "TODO", "TBD", "TBA"}


def _format_bibliography_record(record: Mapping[str, object], key: str) -> str:
    required = {"authors", "year", "journal", "doi"}
    missing = sorted(required - set(record))
    if missing:
        raise ValueError(f"Bibliography record {key!r} is missing: {', '.join(missing)}")
    authors = record["authors"]
    if not isinstance(authors, list) or not authors or not all(isinstance(item, str) and item.strip() for item in authors):
        raise ValueError(f"Bibliography record {key!r} authors must be a non-empty string array")
    year = record["year"]
    if not isinstance(year, int) or isinstance(year, bool) or year < 0:
        raise ValueError(f"Bibliography record {key!r} year must be a non-negative integer")
    journal = record["journal"]
    doi = record["doi"]
    if not isinstance(journal, str) or not journal.strip():
        raise ValueError(f"Bibliography record {key!r} journal must be a non-empty string")
    if not isinstance(doi, str) or not doi.strip():
        raise ValueError(f"Bibliography record {key!r} doi must be a non-empty string")
    author_text = ", ".join(item.strip() for item in authors[:-1])
    if author_text:
        author_text += ", and " + authors[-1].strip()
    else:
        author_text = authors[0].strip()
    parts = [f"{author_text}."]
    title = record.get("title")
    if title is not None:
        if not isinstance(title, str) or not title.strip():
            raise ValueError(f"Bibliography record {key!r} title must be a non-empty string")
        parts.append(title.strip() + ".")
    volume = record.get("volume")
    issue = record.get("issue")
    locator = record.get("locator")
    if volume is not None and not isinstance(volume, (str, int)):
        raise ValueError(f"Bibliography record {key!r} volume must be a string or integer")
    if issue is not None and not isinstance(issue, (str, int)):
        raise ValueError(f"Bibliography record {key!r} issue must be a string or integer")
    if locator is not None and not isinstance(locator, str):
        raise ValueError(f"Bibliography record {key!r} locator must be a string")
    journal_part = f"*{journal.strip()}*"
    if isinstance(year, int):
        journal_part += f", **{year}**"
    if volume is not None:
        journal_part += f", *{str(volume).strip()}*"
    if issue is not None:
        journal_part += f" ({str(issue).strip()})"
    if locator:
        journal_part += f", {locator.strip()}"
    parts.append(journal_part + ".")
    parts.append(f"DOI: {doi.strip().removeprefix('https://doi.org/').removeprefix('doi:').strip()}")
    return " ".join(parts)


def load_bibliography(path: Path) -> dict[str, str]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError(f"Bibliography must be a JSON object: {path}")
    result: dict[str, str] = {}
    for key, value in payload.items():
        if str(key).startswith("_"):
            continue
        if not isinstance(key, str):
            raise ValueError(f"Bibliography entry key must be a string: {key!r}")
        if isinstance(value, str):
            if not value.strip():
                raise ValueError(f"Bibliography entry must be non-empty: {key!r}")
            result[key] = value.strip()
        elif isinstance(value, dict):
            result[key] = _format_bibliography_record(value, key)
        else:
            raise ValueError(f"Bibliography entry must be a string or object: {key!r}")
    return result


def _block_texts(block: Block) -> Iterable[str]:
    if block.text:
        yield block.text
    for row in block.rows:
        yield from row


def _citation_keys(blocks: Iterable[Block]) -> tuple[str, ...]:
    found: list[str] = []
    seen: set[str] = set()
    for block in blocks:
        for text in _block_texts(block):
            for match in CITATION_RE.finditer(text):
                for raw_key in match.group(1).split(","):
                    key = raw_key.strip()
                    if key and key not in seen:
                        seen.add(key)
                        found.append(key)
    return tuple(found)


def _load_citation_base(path: Path) -> dict[str, int]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict) or not isinstance(payload.get("citation_map"), dict):
        raise ValueError(f"Citation base must be a docforge manifest with citation_map: {path}")
    result: dict[str, int] = {}
    for key, value in payload["citation_map"].items():
        if not isinstance(key, str) or not isinstance(value, int) or value < 1:
            raise ValueError(f"Citation base contains an invalid mapping: {key!r} -> {value!r}")
        result[key] = value
    return result


def _citation_plan(
    cited: Sequence[str], bibliography: Mapping[str, str], base: Mapping[str, int] | None
) -> tuple[tuple[str, ...], dict[str, str | int]]:
    """Reuse main numeric labels in SI and prefix SI-only labels with S."""
    if base is None:
        used = tuple(cited)
        return used, {key: index for index, key in enumerate(used, start=1)}
    shared = [key for key in cited if key in base]
    si_only = [key for key in cited if key not in base]
    mapping: dict[str, str | int] = {key: base[key] for key in shared}
    mapping.update({key: f"S{index}" for index, key in enumerate(si_only, start=1)})
    return tuple(shared + si_only), mapping


def _replace_citations(text: str, mapping: Mapping[str, str | int], *, superscript: bool = False) -> str:
    def replace(match: re.Match[str]) -> str:
        keys = [key.strip() for key in match.group(1).split(",") if key.strip()]
        missing = [key for key in keys if key not in mapping]
        if missing:
            raise ValueError(f"Unknown citation key(s): {', '.join(missing)}")
        labels = [mapping[key] for key in keys]
        labels.sort(key=lambda value: (isinstance(value, str), int(str(value).removeprefix("S"))))
        numbers = ",".join(str(value) for value in labels)
        return ("\ue000" + numbers + "\ue001") if superscript else "[" + numbers + "]"

    pattern = r"\s*" + CITATION_RE.pattern if superscript else CITATION_RE.pattern
    return re.sub(pattern, replace, text)


def _replace_block_citations(
    blocks: Iterable[Block], mapping: Mapping[str, str | int], *, superscript: bool = False
) -> tuple[Block, ...]:
    return tuple(
        Block(
            kind=block.kind,
            text=_replace_citations(block.text, mapping, superscript=superscript),
            level=block.level,
            rows=tuple(
                tuple(_replace_citations(value, mapping, superscript=superscript) for value in row)
                for row in block.rows
            ),
            language=block.language,
            path=block.path,
            options=block.options,
        )
        for block in blocks
    )


def _extract_abstract(blocks: Iterable[Block]) -> tuple[str, tuple[Block, ...]]:
    abstract: list[str] = []
    body: list[Block] = []
    collecting = False
    found = False
    for block in blocks:
        is_abstract = (
            block.kind == "heading"
            and block.level == 1
            and block.text.strip().lower() == "abstract"
        )
        if is_abstract and not found:
            collecting = True
            found = True
            continue
        if collecting and block.kind == "heading" and block.level == 1:
            collecting = False
        if collecting:
            if block.text:
                abstract.append(block.text)
            elif block.rows:
                abstract.extend(" ".join(row) for row in block.rows)
        else:
            body.append(block)
    return " ".join(part.strip() for part in abstract if part.strip()), tuple(body)


def _norm_style(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _find_style(document: DocumentType, candidates: Sequence[str]) -> str:
    styles = list(document.styles)
    by_id = {style.style_id: style for style in styles}
    by_name = {style.name: style for style in styles}
    by_norm = {_norm_style(style.name): style for style in styles}
    for candidate in candidates:
        if candidate in by_id:
            return by_id[candidate].style_id
        if candidate in by_name:
            return by_name[candidate].style_id
        style = by_norm.get(_norm_style(candidate))
        if style is not None:
            return style.style_id
    return by_name["Normal"].style_id


def discover_template_styles(document: DocumentType) -> dict[str, str]:
    candidates = {
        "title": ("BBAuthorName", "BB_Author_Name", "BATitle", "BA_Title", "Title"),
        "authors": ("BBAuthorName", "BB_Author_Name", "Author", "Normal"),
        "affiliations": (
            "FACorrespondingAuthorFootnote",
            "FA_Corresponding_Author_Footnote",
            "BCAuthorAddress",
            "BC_Author_Address",
            "Normal",
        ),
        "abstract_title": ("BDAbstractTitle", "BD_Abstract_Title", "Heading 1", "1"),
        "abstract": ("BDAbstract", "BD_Abstract", "Normal"),
        "body": ("TAMainText1", "TAMainText", "TA_Main_Text", "Normal"),
        "heading_1": ("1", "Heading 1"),
        "heading_2": ("2", "Heading 2"),
        "heading_3": ("3", "Heading 3"),
        "caption": ("a4", "Caption", "VA_Figure_Caption"),
        "references_heading": (
            "TFReferencesSection",
            "TF_References_Section",
            "EndNoteBibliographyTitle",
            "Heading 1",
        ),
        "reference": ("EndNoteBibliography", "EndNote Bibliography", "Normal"),
    }
    return {role: _find_style(document, options) for role, options in candidates.items()}


def _resolve_style_profile(
    document: DocumentType, styles: Mapping[str, str], profile: str
) -> dict[str, str]:
    """Resolve an optional JSON semantic-style override against the template."""
    if profile == "template":
        return dict(styles)
    path = Path(profile)
    if not path.is_file():
        raise FileNotFoundError(f"Style profile JSON does not exist: {path}")
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError(f"Style profile must be a JSON object: {path}")
    result = dict(styles)
    known = set(result)
    unknown = sorted(str(key) for key in payload if key not in known)
    if unknown:
        raise ValueError(f"Style profile contains unknown roles: {', '.join(unknown)}")
    available = {style.style_id for style in document.styles} | {style.name for style in document.styles}
    for role, value in payload.items():
        if not isinstance(value, str) or value not in available:
            raise ValueError(f"Style profile names an unavailable style: {role}={value!r}")
        result[role] = document.styles[value].style_id if value in {style.name for style in document.styles} else value
    return result


def _style(document: DocumentType, style_id: str):
    for style in document.styles:
        if style.style_id == style_id:
            return style
    raise ValueError(f"Template style is unavailable: {style_id}")


def _strip_fonts(element) -> None:
    for properties in element.iter(qn("w:rPr")):
        for tag in (qn("w:rFonts"), qn("w:sz"), qn("w:szCs"), qn("w:kern")):
            for child in list(properties.findall(tag)):
                properties.remove(child)


@dataclass(frozen=True)
class _TemplateRegion:
    index: int
    nodes: tuple
    section: object
    figure: bool = False
    image_node: object | None = None
    caption_node: object | None = None


@dataclass(frozen=True)
class _TemplatePrototypes:
    paragraphs: Mapping[str, object]
    figure_regions: tuple[_TemplateRegion, ...]


def _text_of(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def _is_section_break(element) -> bool:
    return element.tag == qn("w:p") and element.find(".//" + qn("w:sectPr")) is not None


def _clean_paragraph_prototype(element):
    """Keep a prototype's paragraph properties but remove example content."""
    result = copy.deepcopy(element)
    properties = result.find(qn("w:pPr"))
    if properties is not None:
        # A paragraph-level rPr controls the mark, not the generated text, and
        # in the reference file contains example-only highlighting.
        mark = properties.find(qn("w:rPr"))
        if mark is not None:
            properties.remove(mark)
        for child in list(properties):
            if child.tag == qn("w:sectPr"):
                properties.remove(child)
    for child in list(result):
        if child.tag != qn("w:pPr"):
            result.remove(child)
    return result


def _run_prototype(element, *, prefer_long: bool = False):
    runs = [run for run in element.findall(".//" + qn("w:r")) if _text_of(run).strip()]
    if not runs:
        runs = list(element.findall(".//" + qn("w:r")))
    if not runs:
        return None
    if prefer_long:
        return copy.deepcopy(max(runs, key=lambda run: len(_text_of(run))))
    return copy.deepcopy(runs[0])


def _sanitize_run_properties(run_properties, *, keep_bold: bool | None = None):
    if run_properties is None:
        return None
    result = copy.deepcopy(run_properties)
    for tag in (qn("w:highlight"), qn("w:shd"), qn("w:rStyle"), qn("w:rPrChange")):
        for child in list(result.findall(tag)):
            result.remove(child)
    if keep_bold is False:
        for child in list(result.findall(qn("w:b"))) + list(result.findall(qn("w:bCs"))):
            result.remove(child)
    return result


def _copy_positive_markers(source_rpr, target_rpr):
    """Carry Markdown emphasis while retaining template font metrics."""
    if source_rpr is None:
        return target_rpr
    if target_rpr is None:
        target_rpr = OxmlElement("w:rPr")
    for tag in ("b", "bCs", "i", "iCs", "u", "strike", "vertAlign"):
        for marker in source_rpr.findall(qn("w:" + tag)):
            value = marker.get(qn("w:val"))
            if value is None or value.lower() not in {"0", "false", "off", "none"}:
                existing = target_rpr.find(qn("w:" + tag))
                if existing is not None:
                    target_rpr.remove(existing)
                target_rpr.append(copy.deepcopy(marker))
    return target_rpr


def _override_run_fonts(element, font_family: str | None, east_asia_font: str | None) -> None:
    """Override generated text fonts only when the caller explicitly requests it."""
    if not font_family and not east_asia_font:
        return
    latin = font_family or east_asia_font
    east_asia = east_asia_font or font_family
    for run in element.iter(qn("w:r")):
        properties = run.find(qn("w:rPr"))
        if properties is None:
            properties = OxmlElement("w:rPr")
            run.insert(0, properties)
        fonts = properties.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            properties.insert(0, fonts)
        for slot in ("ascii", "hAnsi", "cs"):
            if latin:
                fonts.set(qn(f"w:{slot}"), latin)
        if east_asia:
            fonts.set(qn("w:eastAsia"), east_asia)
        for slot in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn(f"w:{slot}"), None)


def _override_heading_before(element, style_ids: set[str], points: float | None) -> None:
    """Set explicit spacing before generated headings when requested."""
    if points is None:
        return
    before = str(max(0, round(points * 20)))
    for paragraph in element.iter(qn("w:p")):
        properties = paragraph.find(qn("w:pPr"))
        style = properties.find(qn("w:pStyle")) if properties is not None else None
        if style is None or style.get(qn("w:val")) not in style_ids:
            continue
        spacing = properties.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            properties.append(spacing)
        spacing.set(qn("w:before"), before)


def _new_paragraph(
    document: DocumentType,
    style_id: str,
    text: str,
    *,
    prototype=None,
    bold_default: bool = False,
    uppercase: bool = False,
    font_family: str | None = None,
    east_asia_font: str | None = None,
):
    """Render text with a template paragraph and run-format prototype."""
    temporary = Document()
    paragraph = temporary.add_paragraph()
    add_inline(paragraph, text.upper() if uppercase else text, bold_default=bold_default)
    element = copy.deepcopy(paragraph._p)
    if prototype is not None:
        original_properties = prototype.find(qn("w:pPr"))
        properties = element.find(qn("w:pPr"))
        if properties is not None:
            element.remove(properties)
        if original_properties is not None:
            element.insert(0, _clean_paragraph_prototype(prototype).find(qn("w:pPr")))
    properties = element.find(qn("w:pPr"))
    if properties is None:
        properties = OxmlElement("w:pPr")
        element.insert(0, properties)
    style = properties.find(qn("w:pStyle"))
    if style is None:
        style = OxmlElement("w:pStyle")
        properties.insert(0, style)
    style.set(qn("w:val"), style_id)
    # A generated heading must expose the same outline level as its prototype.
    heading_match = re.fullmatch(r"(?:Heading ?)?([1-9])", style_id)
    if heading_match is not None:
        heading_level = int(heading_match.group(1))
        outline = properties.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            properties.append(outline)
        outline.set(qn("w:val"), str(heading_level - 1))

    base_run = _run_prototype(prototype, prefer_long=True) if prototype is not None else None
    base_rpr = _sanitize_run_properties(
        base_run.find(qn("w:rPr")) if base_run is not None else None,
        keep_bold=None,
    )
    for run in element.findall(".//" + qn("w:r")):
        source_rpr = run.find(qn("w:rPr"))
        if source_rpr is None and base_rpr is None:
            continue
        if run.find(qn("w:rPr")) is not None:
            run.remove(run.find(qn("w:rPr")))
        if base_rpr is not None:
            run.insert(0, copy.deepcopy(base_rpr))
        target_rpr = run.find(qn("w:rPr"))
        merged_rpr = _copy_positive_markers(source_rpr, target_rpr)
        if target_rpr is None and merged_rpr is not None and len(merged_rpr):
            run.insert(0, merged_rpr)
    _override_run_fonts(element, font_family, east_asia_font)
    return element


def _new_metadata_paragraph(document: DocumentType, style_id: str, text: str, *, prototype=None):
    """Render metadata literally so correspondence asterisks cannot become emphasis."""
    temporary = Document()
    paragraph = temporary.add_paragraph()
    paragraph.add_run(text)
    element = copy.deepcopy(paragraph._p)
    if prototype is not None:
        properties = element.find(qn("w:pPr"))
        if properties is not None:
            element.remove(properties)
        original_properties = _clean_paragraph_prototype(prototype).find(qn("w:pPr"))
        if original_properties is not None:
            element.insert(0, original_properties)
    properties = element.find(qn("w:pPr"))
    if properties is None:
        properties = OxmlElement("w:pPr")
        element.insert(0, properties)
    style = properties.find(qn("w:pStyle"))
    if prototype is None:
        if style is None:
            style = OxmlElement("w:pStyle")
            properties.insert(0, style)
        style.set(qn("w:val"), style_id)
    base_run = _run_prototype(prototype, prefer_long=True) if prototype is not None else None
    base_rpr = _sanitize_run_properties(base_run.find(qn("w:rPr")) if base_run is not None else None)
    for run in element.findall(".//" + qn("w:r")):
        existing = run.find(qn("w:rPr"))
        if existing is not None:
            run.remove(existing)
        if base_rpr is not None:
            run.insert(0, copy.deepcopy(base_rpr))
    return element


def _remap_styles(element, source: DocumentType, target: DocumentType) -> None:
    source_styles = {style.style_id: style for style in source.styles}
    target_styles = {style.name: style for style in target.styles}
    for tag in (qn("w:pStyle"), qn("w:rStyle"), qn("w:tblStyle")):
        for node in element.iter(tag):
            source_style = source_styles.get(node.get(qn("w:val"), ""))
            target_style = target_styles.get(source_style.name) if source_style else None
            if target_style is not None:
                node.set(qn("w:val"), target_style.style_id)


def _add_body_style(element, body_style_id: str) -> None:
    for paragraph in element.iter(qn("w:p")):
        properties = paragraph.find(qn("w:pPr"))
        if properties is None:
            properties = OxmlElement("w:pPr")
            paragraph.insert(0, properties)
        if properties.find(qn("w:pStyle")) is None:
            style = OxmlElement("w:pStyle")
            style.set(qn("w:val"), body_style_id)
            properties.insert(0, style)


def _body_column_width_twips(document: DocumentType) -> int:
    """Return the usable width of the section receiving generated body nodes."""
    section = document.sections[1] if len(document.sections) > 1 else document.sections[0]
    page = int(section.page_width / 635)
    margins = int((section.left_margin + section.right_margin) / 635)
    columns = section._sectPr.find(qn("w:cols"))
    count = int(columns.get(qn("w:num"), "1")) if columns is not None else 1
    gap = int(columns.get(qn("w:space"), "0")) if columns is not None else 0
    return max(1, (page - margins - gap * max(count - 1, 0)) // max(count, 1))


def _fit_tables(element, width_twips: int) -> None:
    """Scale generated table grids to the receiving section's column width."""
    for table in element.iter(qn("w:tbl")):
        grid = table.find(qn("w:tblGrid"))
        if grid is None:
            continue
        columns = grid.findall(qn("w:gridCol"))
        old = [max(1, int(column.get(qn("w:w"), "1"))) for column in columns]
        if not old:
            continue
        total = sum(old)
        scaled = [max(240, round(width_twips * value / total)) for value in old]
        scaled[-1] += width_twips - sum(scaled)
        for column, value in zip(columns, scaled):
            column.set(qn("w:w"), str(value))

        properties = table.find(qn("w:tblPr"))
        if properties is None:
            properties = OxmlElement("w:tblPr")
            table.insert(0, properties)
        table_width = properties.find(qn("w:tblW"))
        if table_width is None:
            table_width = OxmlElement("w:tblW")
            properties.insert(0, table_width)
        table_width.set(qn("w:w"), str(width_twips))
        table_width.set(qn("w:type"), "dxa")
        layout = properties.find(qn("w:tblLayout"))
        if layout is None:
            layout = OxmlElement("w:tblLayout")
            properties.append(layout)
        layout.set(qn("w:type"), "fixed")
        rows = table.findall(qn("w:tr"))
        for row_index, row in enumerate(rows):
            row_properties = row.find(qn("w:trPr"))
            if row_properties is None:
                row_properties = OxmlElement("w:trPr")
                row.insert(0, row_properties)
            cant_split = row_properties.find(qn("w:cantSplit"))
            if cant_split is None:
                cant_split = OxmlElement("w:cantSplit")
                row_properties.append(cant_split)
            cells = row.findall(qn("w:tc"))
            for cell, value in zip(cells, scaled):
                cell_properties = cell.find(qn("w:tcPr"))
                if cell_properties is None:
                    cell_properties = OxmlElement("w:tcPr")
                    cell.insert(0, cell_properties)
                cell_width = cell_properties.find(qn("w:tcW"))
                if cell_width is None:
                    cell_width = OxmlElement("w:tcW")
                    cell_properties.insert(0, cell_width)
                cell_width.set(qn("w:w"), str(value))
                cell_width.set(qn("w:type"), "dxa")
                if row_index == 0:
                    for paragraph in cell.findall(qn("w:p")):
                        ppr = paragraph.find(qn("w:pPr"))
                        if ppr is None:
                            ppr = OxmlElement("w:pPr")
                            paragraph.insert(0, ppr)
                        if ppr.find(qn("w:keepNext")) is None:
                            ppr.append(OxmlElement("w:keepNext"))
        # Small audit tables are easier to read as a unit.  Keeping their
        # rows together prevents a continuation page without its header;
        # larger tables are deliberately allowed to flow across pages.
        if len(rows) <= 7:
            for row in rows[:-1]:
                for paragraph in row.iter(qn("w:p")):
                    ppr = paragraph.find(qn("w:pPr"))
                    if ppr is None:
                        ppr = OxmlElement("w:pPr")
                        paragraph.insert(0, ppr)
                    if ppr.find(qn("w:keepNext")) is None:
                        ppr.append(OxmlElement("w:keepNext"))


def _fit_equation_tabs(element, width_twips: int) -> None:
    """Place generated right-aligned equation numbers at the active column edge."""
    for tab in element.iter(qn("w:tab")):
        if tab.get(qn("w:val")) == "right":
            tab.set(qn("w:pos"), str(max(1, width_twips)))


def _clone_body(blocks: Iterable[Block], target: DocumentType, styles: Mapping[str, str]) -> list:
    source = render_blocks_to_doc(blocks)
    result: list = []
    body_width = _body_column_width_twips(target)
    for child in source._element.body.iterchildren():
        if child.tag == qn("w:sectPr"):
            continue
        clone = copy.deepcopy(child)
        remap_image_relationships(clone, source, target)
        _remap_styles(clone, source, target)
        _add_body_style(clone, styles["body"])
        _fit_tables(clone, body_width)
        result.append(clone)
    return result


def _template_regions(document: DocumentType) -> tuple[_TemplateRegion, ...]:
    body = document._element.body
    regions: list[_TemplateRegion] = []
    pending: list = []
    index = 0
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            regions.append(_TemplateRegion(index, tuple(pending), copy.deepcopy(child)))
            pending = []
            index += 1
            continue
        if _is_section_break(child):
            properties = child.find(qn("w:pPr"))
            section = properties.find(qn("w:sectPr")) if properties is not None else None
            if section is None:
                pending.append(copy.deepcopy(child))
                continue
            nodes = tuple(pending)
            regions.append(_TemplateRegion(index, nodes, copy.deepcopy(section)))
            pending = []
            index += 1
            continue
        pending.append(copy.deepcopy(child))
    if pending:
        terminal = body.sectPr
        if terminal is None:
            raise ValueError("Template has no section properties")
        regions.append(_TemplateRegion(index, tuple(pending), copy.deepcopy(terminal)))
    if not regions:
        raise ValueError("Template has no section properties")

    marked: list[_TemplateRegion] = []
    for region in regions:
        drawing = next((node for node in region.nodes if next(node.iter(qn("w:drawing")), None) is not None), None)
        caption = next(
            (
                node
                for node in region.nodes
                if node.tag == qn("w:p")
                and _is_figure_caption_text(_text_of(node))
            ),
            None,
        )
        non_caption_text = [
            _text_of(node).strip()
            for node in region.nodes
            if node.tag == qn("w:p") and _text_of(node).strip() and node is not caption
        ]
        is_figure = drawing is not None and caption is not None and not non_caption_text
        marked.append(
            _TemplateRegion(region.index, region.nodes, region.section, is_figure, drawing, caption)
        )
    return tuple(marked)


def _template_inline_figure_slots(regions) -> tuple[_TemplateRegion, ...]:
    """Find every image/caption pair, including pairs inside mixed regions.

    Some SI templates keep sample prose, figures, and captions in one section.
    A section-level figure test would miss those slots because the prose is not
    part of the figure itself.  Pairing the drawing paragraph with the next
    non-empty caption paragraph keeps the template's own section geometry while
    making every explicit figure slot reusable.
    """
    slots: list[_TemplateRegion] = []
    for region in regions:
        for index, node in enumerate(region.nodes):
            if node.tag != qn("w:p") or next(node.iter(qn("w:drawing")), None) is None:
                continue
            for following in region.nodes[index + 1:]:
                if following.tag != qn("w:p"):
                    continue
                text = _text_of(following).strip()
                if not text:
                    continue
                if _is_figure_caption_text(text):
                    slots.append(
                        _TemplateRegion(
                            region.index,
                            (node, following),
                            region.section,
                            True,
                            node,
                            following,
                        )
                    )
                break
    return tuple(slots)


def _template_prototypes(document: DocumentType, styles: Mapping[str, str], regions):
    paragraphs = list(document.paragraphs)
    by_style: dict[str, list] = {}
    for paragraph in paragraphs:
        by_style.setdefault(paragraph.style.style_id, []).append(paragraph._p)

    def choose(role: str, *, long: bool = False, text_match: str | None = None):
        values = by_style.get(styles[role], [])
        if text_match:
            values = [node for node in values if text_match in _text_of(node).lower()] or values
        values = [node for node in values if _text_of(node).strip()] or values
        if not values:
            return None
        return copy.deepcopy(max(values, key=lambda node: len(_text_of(node)))) if long else copy.deepcopy(values[0])

    # The reference file uses the same heading style IDs as the numbering
    # definitions; selecting an actual paragraph preserves its numPr/spacing.
    heading_1 = choose("heading_1")
    heading_2 = choose("heading_2")
    heading_3 = choose("heading_3")
    body_candidates = [
        (index, node)
        for index, node in enumerate(paragraphs)
        if node.style.style_id in {styles["body"], "a", "Normal"}
        and _text_of(node._p).strip()
        and "<w:drawing" not in node._p.xml
    ]
    heading_style_ids = {
        styles["heading_1"],
        styles["heading_2"],
        styles["heading_3"],
    }
    first_heading = next(
        (
            index
            for index, node in enumerate(paragraphs)
            if node.style.style_id in heading_style_ids and _text_of(node._p).strip()
        ),
        -1,
    )
    post_front = [
        node
        for index, node in body_candidates
        if index > first_heading
    ]
    body = copy.deepcopy(
        max(post_front or [node for _, node in body_candidates], key=lambda node: len(_text_of(node._p)))._p
    ) if (post_front or body_candidates) else choose("body")
    caption = choose("caption", long=True)
    references = choose("reference", long=True)
    references_heading = next((copy.deepcopy(paragraph._p) for paragraph in paragraphs if paragraph.text.strip().lower() in {"references", "bibliography"}), None)
    supporting_information = next(
        (copy.deepcopy(paragraph._p) for paragraph in paragraphs if paragraph.text.strip().lower() == "supporting information"),
        None,
    )
    si_title = None
    si_authors = None
    si_affiliations = None
    si_contacts = None
    if supporting_information is not None:
        front = [paragraph for paragraph in paragraphs[:first_heading] if paragraph.text.strip()]
        if len(front) >= 5:
            _, si_title, si_authors, si_affiliations, si_contacts = [copy.deepcopy(paragraph._p) for paragraph in front[:5]]
    authors = si_authors if si_authors is not None else choose("authors")
    if styles["authors"] == styles["title"]:
        same_style = by_style.get(styles["authors"], [])
        authors = copy.deepcopy(same_style[1]) if len(same_style) > 1 else None
    paragraphs_map = {
        "supporting_information": supporting_information,
        "title": si_title if si_title is not None else choose("title"),
        "authors": authors,
        "affiliations": si_affiliations if si_affiliations is not None else choose("affiliations"),
        "contacts": si_contacts if si_contacts is not None else choose("affiliations"),
        "abstract": choose("abstract", long=True),
        "body": body,
        "heading_1": heading_1,
        "heading_2": heading_2,
        "heading_3": heading_3,
        "caption": caption,
        "reference": references,
        "references_heading": references_heading if references_heading is not None else choose("references_heading", text_match="references"),
    }
    figure_regions = _template_inline_figure_slots(regions)
    if not figure_regions:
        figure_regions = tuple(region for region in regions if region.figure)
    return _TemplatePrototypes(paragraphs_map, figure_regions)


def _template_uses_superscript_citations(document: DocumentType) -> bool:
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            value = run.text.strip()
            marker = run._r.find(".//" + qn("w:vertAlign"))
            if marker is not None and marker.get(qn("w:val")) == "superscript" and re.fullmatch(r"[0-9]+(?:[,-][0-9]+)*", value):
                return True
    return False


def _materialize_superscript_citations(document: DocumentType) -> int:
    """Replace private citation sentinels with template-style superscripts."""
    converted = 0
    pattern = re.compile(r"(\ue000(?:S?[0-9]+)(?:,(?:S?[0-9]+))*\ue001)")
    # Materialize the run list before replacing nodes; mutating a live lxml
    # iterator otherwise skips sibling paragraphs after the first citation.
    for parent in list(document._element.body.iter(qn("w:r"))):
        text_nodes = [child for child in parent if child.tag == qn("w:t") and child.text and pattern.search(child.text)]
        if not text_nodes:
            continue
        grandparent = parent.getparent()
        if grandparent is None:
            continue
        base_properties = parent.find(qn("w:rPr"))
        replacements = []
        for node in text_nodes:
            for fragment in pattern.split(node.text):
                if not fragment:
                    continue
                run = OxmlElement("w:r")
                if base_properties is not None:
                    run.append(copy.deepcopy(base_properties))
                text = OxmlElement("w:t")
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                if fragment.startswith("\ue000"):
                    text.text = fragment[1:-1]
                    properties = run.find(qn("w:rPr"))
                    if properties is None:
                        properties = OxmlElement("w:rPr")
                        run.insert(0, properties)
                    for existing in list(properties.findall(qn("w:vertAlign"))):
                        properties.remove(existing)
                    marker = OxmlElement("w:vertAlign")
                    marker.set(qn("w:val"), "superscript")
                    properties.append(marker)
                    converted += 1
                else:
                    text.text = fragment
                run.append(text)
                replacements.append(run)
        position = list(grandparent).index(parent)
        grandparent.remove(parent)
        for offset, run in enumerate(replacements):
            grandparent.insert(position + offset, run)
    return converted


FIGURE_CAPTION_RE = re.compile(r"\s*(?:Figure|Scheme|Chart)\s+[A-Za-z0-9]+[.:|]?", re.IGNORECASE)
FIGURE_CAPTION_SLOT_RE = re.compile(
    r"\s*\[(?:Figure|Scheme|Chart)\s+Caption\]", re.IGNORECASE
)
CAPTION_LABEL_RE = re.compile(
    r"^\s*(Figure|Scheme|Chart)\s+[A-Za-z]*\d+\s*[.:]?\s*",
    re.IGNORECASE,
)


def _numbered_caption(caption: str, number: int, prefix: str) -> tuple[str, str]:
    match = CAPTION_LABEL_RE.match(caption)
    kind = match.group(1).capitalize() if match else "Figure"
    body = caption[match.end():].strip() if match else caption.strip()
    label = f"{kind} {prefix}{number}"
    return label, f"{label}. {body}".rstrip()


def _is_figure_caption_text(text: str) -> bool:
    """Recognize rendered captions and ACS-style placeholder caption slots."""
    return bool(FIGURE_CAPTION_RE.match(text) or FIGURE_CAPTION_SLOT_RE.match(text))


def _figure_groups(blocks: Sequence[Block]) -> tuple[tuple[tuple[Block, ...], Block | None, str], ...]:
    """Split Markdown into text runs and image/caption pairs.

    A caption immediately following an image is consumed as the caption.  An
    image alt text remains a usable fallback, so ordinary Markdown image
    syntax stays sufficient for callers.
    """
    groups: list[tuple[tuple[Block, ...], Block | None, str]] = []
    text: list[Block] = []
    index = 0
    while index < len(blocks):
        block = blocks[index]
        if block.kind != "image":
            text.append(block)
            index += 1
            continue
        groups.append((tuple(text), block, block.text.strip()))
        text = []
        index += 1
        if index < len(blocks) and blocks[index].kind == "paragraph" and FIGURE_CAPTION_RE.match(blocks[index].text):
            groups[-1] = (groups[-1][0], block, blocks[index].text.strip())
            index += 1
    groups.append((tuple(text), None, ""))
    return tuple(groups)


def _section_break_node(section) -> object:
    paragraph = OxmlElement("w:p")
    properties = OxmlElement("w:pPr")
    properties.append(copy.deepcopy(section))
    paragraph.append(properties)
    return paragraph


def _section_column_count(section) -> int:
    columns = section.find(qn("w:cols"))
    if columns is None:
        return 1
    return max(1, int(columns.get(qn("w:num"), "1")))


def _section_width_twips(section) -> int:
    page = section.find(qn("w:pgSz"))
    margins = section.find(qn("w:pgMar"))
    if page is None or margins is None:
        return 7200
    width = int(page.get(qn("w:w"), "12240"))
    left = int(margins.get(qn("w:left"), "0"))
    right = int(margins.get(qn("w:right"), "0"))
    columns = section.find(qn("w:cols"))
    count = _section_column_count(section)
    gap = int(columns.get(qn("w:space"), "0")) if columns is not None else 0
    return max(1, (width - left - right - gap * max(count - 1, 0)) // count)


def _with_column_layout(section, layout: str):
    if layout not in {"one", "two"}:
        raise ValueError("image columns must be one or two")
    section = copy.deepcopy(section)
    columns = section.find(qn("w:cols"))
    if columns is None:
        columns = OxmlElement("w:cols")
        section.append(columns)
    if layout == "one":
        columns.attrib.pop(qn("w:num"), None)
        columns.attrib.pop(qn("w:space"), None)
    else:
        columns.set(qn("w:num"), "2")
        columns.set(qn("w:space"), columns.get(qn("w:space"), "475"))
    return section


def _apply_template_run_formatting(element, prototype) -> None:
    if prototype is None:
        return
    base = _run_prototype(prototype, prefer_long=True)
    base_rpr = _sanitize_run_properties(base.find(qn("w:rPr")) if base is not None else None)
    for run in element.findall(".//" + qn("w:r")):
        source = run.find(qn("w:rPr"))
        if source is not None:
            run.remove(source)
        if base_rpr is not None:
            run.insert(0, copy.deepcopy(base_rpr))
        target_rpr = run.find(qn("w:rPr"))
        merged_rpr = _copy_positive_markers(source, target_rpr)
        if target_rpr is None and merged_rpr is not None and len(merged_rpr):
            run.insert(0, merged_rpr)


def _clone_rendered_block(
    block: Block,
    target: DocumentType,
    styles: Mapping[str, str],
    prototypes: _TemplatePrototypes,
    *,
    equation_number: int | None = None,
    table_number: int | None = None,
    number_prefix: str = "",
) -> list:
    """Render one block and graft the template's paragraph prototype."""
    p = prototypes.paragraphs
    if block.kind == "table_caption":
        label = f"Table {number_prefix}{table_number}. " if table_number is not None else "Table. "
        return [
            _new_paragraph(
                target,
                styles["caption"],
                label + block.text,
                prototype=p.get("caption"),
            )
        ]
    if block.kind in {"paragraph", "reference"}:
        return [_new_paragraph(target, styles["body"], block.text, prototype=p.get("body"))]
    if block.kind == "heading":
        level = min(max(block.level, 1), 3)
        role = f"heading_{level}"
        return [_new_paragraph(target, styles[role], block.text, prototype=p.get(role), bold_default=False, uppercase=level == 1)]
    if block.kind in {"paragraph", "reference", "quote", "ordered", "bullet", "code", "equation", "table", "separator"}:
        generated = render_blocks_to_doc([block], equation_start=equation_number or 1, number_prefix=number_prefix)
        result: list = []
        width = _body_column_width_twips(target)
        for child in generated._element.body.iterchildren():
            if child.tag == qn("w:sectPr"):
                continue
            clone = copy.deepcopy(child)
            remap_image_relationships(clone, generated, target)
            _remap_styles(clone, generated, target)
            _fit_tables(clone, width)
            if block.kind == "equation":
                properties = clone.find(qn("w:pPr"))
                if properties is None:
                    properties = OxmlElement("w:pPr")
                    clone.insert(0, properties)
                style = properties.find(qn("w:pStyle"))
                if style is None:
                    style = OxmlElement("w:pStyle")
                    properties.insert(0, style)
                style.set(qn("w:val"), styles["body"])
                _fit_equation_tabs(clone, width)
            if block.kind in {"paragraph", "reference", "quote", "ordered", "bullet"}:
                properties = clone.find(qn("w:pPr"))
                if properties is None:
                    properties = OxmlElement("w:pPr")
                    clone.insert(0, properties)
                old_properties = properties
                source_properties = _clean_paragraph_prototype(p.get("body")) if p.get("body") is not None else None
                # Keep list indentation and quote offsets from the generated
                # paragraph while importing the template's font/spacing.
                keep = {qn("w:ind"), qn("w:numPr"), qn("w:tabs"), qn("w:jc")}
                for child_prop in list(old_properties):
                    if child_prop.tag not in keep and child_prop.tag != qn("w:pStyle"):
                        old_properties.remove(child_prop)
                if source_properties is not None:
                    source_ppr = source_properties.find(qn("w:pPr"))
                    if source_ppr is not None:
                        for child_prop in source_ppr:
                            if child_prop.tag not in {qn("w:pStyle"), qn("w:ind"), qn("w:numPr")} and old_properties.find(child_prop.tag) is None:
                                old_properties.append(copy.deepcopy(child_prop))
                style = old_properties.find(qn("w:pStyle"))
                if style is None:
                    style = OxmlElement("w:pStyle")
                    old_properties.insert(0, style)
                style.set(qn("w:val"), styles["body"])
                _apply_template_run_formatting(clone, p.get("body"))
            result.append(clone)
        return result
    return []


def _clone_figure(
    target: DocumentType,
    image: Block,
    caption: str,
    prototype: _TemplateRegion | None,
    caption_prototype,
    caption_style_id: str,
    figure_number: int,
    section_width_twips: int | None = None,
    number_prefix: str = "",
) -> list:
    image_path = Path(image.path)
    if not image_path.is_file():
        raise FileNotFoundError(f"Markdown image does not exist: {image_path}")
    if prototype is not None and prototype.image_node is not None:
        paragraph = copy.deepcopy(prototype.image_node)
        properties = paragraph.find(qn("w:pPr"))
        if properties is None:
            properties = OxmlElement("w:pPr")
            paragraph.insert(0, properties)
        for child in list(paragraph):
            if child.tag != qn("w:pPr"):
                paragraph.remove(child)
        old_run = next(iter(prototype.image_node.findall(".//" + qn("w:r"))), None)
        run = OxmlElement("w:r")
        if old_run is not None and old_run.find(qn("w:rPr")) is not None:
            run.append(copy.deepcopy(old_run.find(qn("w:rPr"))))
        drawing = next(iter(prototype.image_node.findall(".//" + qn("w:drawing"))), None)
        if drawing is None:
            raise ValueError("Figure prototype has no drawing")
        drawing = copy.deepcopy(drawing)
        relation_id, _ = target.part.get_or_add_image(BytesIO(image_path.read_bytes()))
        for blip in drawing.iter(qn("a:blip")):
            blip.set(qn("r:embed"), relation_id)
        # Keep the reference slot width, but derive height from the actual
        # asset so the image is never stretched or cropped.
        extent = next(iter(drawing.iter(qn("wp:extent"))), None)
        source_extent = next(iter(prototype.image_node.iter(qn("wp:extent"))), None)
        if extent is not None and source_extent is not None:
            width = int(source_extent.get("cx", "1"))
            if section_width_twips is not None:
                width = min(width, section_width_twips * 635)
            with Image.open(image_path) as image_file:
                ratio = image_file.height / max(1, image_file.width)
            height = max(1, round(width * ratio))
            extent.set("cx", str(width))
            extent.set("cy", str(height))
            for xfrm in drawing.iter(qn("a:xfrm")):
                ext = next(iter(xfrm.iter(qn("a:ext"))), None)
                if ext is not None:
                    ext.set("cx", str(width))
                    ext.set("cy", str(height))
        for doc_pr in drawing.iter(qn("wp:docPr")):
            doc_pr.set("id", str(1000 + figure_number))
            doc_pr.set("name", image_path.stem)
            doc_pr.set("descr", caption[:255])
        run.append(drawing)
        paragraph.append(run)
    else:
        # Templates without explicit figure slots still receive a real inline
        # drawing, constrained to the current section column.
        temporary = Document()
        picture = temporary.add_paragraph()
        with Image.open(image_path) as image_file:
            available = section_width_twips * 635 if section_width_twips is not None else _body_column_width_twips(target) * 635
            width = min(914400 * 6.5, max(914400, available))
            picture.add_run().add_picture(str(image_path), width=width)
        paragraph = copy.deepcopy(picture._p)
        relation_id, _ = target.part.get_or_add_image(BytesIO(image_path.read_bytes()))
        for blip in paragraph.iter(qn("a:blip")):
            blip.set(qn("r:embed"), relation_id)
    _, caption_text = _numbered_caption(caption, figure_number, number_prefix)
    caption_text = re.sub(r"^((?:Figure|Scheme|Chart)\s+[A-Za-z0-9]+[.:])", r"**\1**", caption_text)
    cap = _new_paragraph(target, caption_style_id, caption_text, prototype=caption_prototype)
    return [paragraph, cap]


def _section_columns(document: DocumentType, index: int = 1) -> int:
    section = document.sections[min(index, len(document.sections) - 1)]
    columns = section._sectPr.find(qn("w:cols"))
    if columns is None:
        return 1
    return max(1, int(columns.get(qn("w:num"), "1")))


def _set_body_columns(
    section_breaks: Sequence,
    layout: str,
    document: DocumentType,
) -> None:
    """Apply an explicit one/two-column choice to the generated body section."""
    if layout == "template":
        return
    if layout not in {"one", "two"}:
        raise ValueError("columns must be one of: template, one, two")
    if len(section_breaks) >= 2:
        properties = section_breaks[1].find(qn("w:pPr"))
        section = properties.find(qn("w:sectPr"))
    else:
        section = document._element.body.sectPr
    columns = section.find(qn("w:cols"))
    if columns is None:
        columns = OxmlElement("w:cols")
        section.append(columns)
    if layout == "one":
        columns.attrib.pop(qn("w:num"), None)
        columns.attrib.pop(qn("w:space"), None)
    else:
        columns.set(qn("w:num"), "2")
        columns.set(qn("w:space"), columns.get(qn("w:space"), "475"))


def _insert_before(body, anchor, nodes: Iterable) -> None:
    index = list(body).index(anchor)
    for node in nodes:
        body.insert(index, node)
        index += 1


def _geometry(document: DocumentType) -> tuple[tuple[int, int, int, int, int, int], ...]:
    return tuple(
        (
            section.page_width,
            section.page_height,
            section.top_margin,
            section.bottom_margin,
            section.left_margin,
            section.right_margin,
        )
        for section in document.sections
    )


def _update_fields(document: DocumentType) -> None:
    """Disable automatic field updates that trigger Word's external-field prompt."""
    settings = document.settings.element
    node = settings.find(qn("w:updateFields"))
    if node is not None:
        settings.remove(node)


def _source_part(rels_name: str) -> str | None:
    path = PurePosixPath(rels_name)
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        return None
    return str(path.parent.parent / path.name[:-5])


def _target(source: str, value: str) -> str:
    if value.startswith("/"):
        return value[1:]
    return str(PurePosixPath(source).parent / value)


def _used_rel_ids(payload: bytes) -> set[str]:
    root = etree.fromstring(payload)
    return {
        value
        for element in root.iter()
        for attribute, value in element.attrib.items()
        if attribute.startswith("{" + REL_NS + "}") and value
    }


def _prune_images(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    targets: set[str] = set()
    for rels_name, payload in list(files.items()):
        source = _source_part(rels_name)
        if source is None or source not in files:
            continue
        root = etree.fromstring(payload)
        used = _used_rel_ids(files[source])
        changed = False
        for relationship in list(root):
            if not relationship.get("Type", "").endswith("/image"):
                continue
            relation_id = relationship.get("Id", "")
            if relation_id not in used:
                root.remove(relationship)
                changed = True
            else:
                targets.add(_target(source, relationship.get("Target", "")))
        if changed:
            files[rels_name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
    for name in list(files):
        if name.startswith("word/media/") and name not in targets:
            del files[name]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in files.items():
            archive.writestr(name, payload)


def _verify_images(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    targets: set[str] = set()
    for rels_name, payload in files.items():
        source = _source_part(rels_name)
        if source is None or source not in files:
            continue
        root = etree.fromstring(payload)
        used = _used_rel_ids(files[source])
        for relationship in root:
            if not relationship.get("Type", "").endswith("/image"):
                continue
            relation_id = relationship.get("Id", "")
            if relation_id not in used:
                raise RuntimeError(f"Unused image relationship: {rels_name}#{relation_id}")
            target = _target(source, relationship.get("Target", ""))
            if target not in files:
                raise RuntimeError(f"Missing image target: {target}")
            targets.add(target)
    leftovers = [name for name in files if name.startswith("word/media/") and name not in targets]
    if leftovers:
        raise RuntimeError(f"Unreferenced image parts remain: {leftovers[:10]}")


def _package_text(path: Path) -> str:
    fragments: list[str] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            try:
                root = etree.fromstring(archive.read(name))
            except etree.XMLSyntaxError:
                continue
            fragments.extend(node.text or "" for node in root.iter(qn("w:t")))
            fragments.extend(node.text or "" for node in root.iter(qn("w:instrText")))
    return "\n".join(fragments)


def verify_template_output(
    path: Path,
    *,
    expected_sections: int | None = None,
    expected_geometry: Sequence[tuple[int, int, int, int, int, int]] | None = None,
    expected_body_columns: int | None = None,
) -> dict[str, object]:
    verify_navigation_headings(path)
    verify_image_relationships(path)
    verify_clean_review_state(path)
    _verify_images(path)
    with zipfile.ZipFile(path) as archive:
        settings = etree.fromstring(archive.read("word/settings.xml"))
        update_fields = settings.find(qn("w:updateFields"))
        if update_fields is not None and update_fields.get(qn("w:val"), "true").lower() == "true":
            raise RuntimeError("Generated DOCX enables automatic field updates on open")
        unsafe = re.compile(r"\b(?:INCLUDETEXT|INCLUDEPICTURE|DDE(?:AUTO)?|LINK)\b", re.IGNORECASE)
        for name in archive.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = etree.fromstring(archive.read(name))
            field_codes = " ".join(node.text or "" for node in root.iter(qn("w:instrText")))
            if unsafe.search(field_codes):
                raise RuntimeError(f"Unsafe external-file field remains in generated DOCX: {name}")
    document = Document(path)
    geometry = _geometry(document)
    if expected_sections is not None and len(document.sections) != expected_sections:
        raise RuntimeError(
            f"Template section count changed: expected {expected_sections}, found {len(document.sections)}"
        )
    if expected_geometry is not None and tuple(expected_geometry) != geometry:
        raise RuntimeError("Template page geometry changed during assembly")
    body_columns = _section_columns(document)
    if expected_body_columns is not None and body_columns != expected_body_columns:
        raise RuntimeError(
            f"Template body column count changed: expected {expected_body_columns}, found {body_columns}"
        )
    text = _package_text(path)
    lower = text.lower()
    retained = [token for token in FORBIDDEN_TOKENS if token in lower]
    if retained:
        raise RuntimeError(f"Template placeholder text remains: {retained}")
    if CITATION_RE.search(text):
        raise RuntimeError("Raw Markdown citation token remains in generated DOCX")
    if "\ue000" in text or "\ue001" in text:
        raise RuntimeError("Unmaterialized citation marker remains in generated DOCX")
    return {
        "clean_review_state": True,
        "image_relationships": "resolved",
        "placeholder_tokens": [],
        "raw_citation_tokens": 0,
        "automatic_field_updates": False,
        "sections": len(document.sections),
        "body_columns": body_columns,
        "section_geometry": [list(value) for value in geometry],
    }


def assemble_markdown_template(
    inputs: Sequence[Path],
    *,
    template_path: Path,
    output: Path,
    metadata_path: Path | None = None,
    bibliography_path: Path | None = None,
    citation_base_path: Path | None = None,
    citation_format: str = "template",
    title: str = "",
    keep_comments: bool = False,
    skip_images: bool = False,
    columns: str = "template",
    font_family: str | None = None,
    east_asia_font: str | None = None,
    style_profile: str = "template",
    line_numbers: str = "template",
    include_title: bool = True,
    strip_level_one_headings: bool = False,
    heading_before: float | None = None,
    numbering_prefix: str = "",
    bibliography_scope: str = "all",
    include_metadata_back_matter: bool = True,
    force: bool = False,
) -> AssemblyResult:
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    if not inputs:
        raise ValueError("Provide at least one Markdown input")
    if not template_path.is_file():
        raise FileNotFoundError(f"Template DOCX does not exist: {template_path}")
    if metadata_path is not None and not metadata_path.is_file():
        raise FileNotFoundError(f"Metadata Markdown does not exist: {metadata_path}")
    if bibliography_path is not None and not bibliography_path.is_file():
        raise FileNotFoundError(f"Bibliography JSON does not exist: {bibliography_path}")
    if citation_base_path is not None and not citation_base_path.is_file():
        raise FileNotFoundError(f"Citation base manifest does not exist: {citation_base_path}")
    if line_numbers not in {"template", "on", "off"}:
        raise ValueError("line_numbers must be one of: template, on, off")
    if columns not in {"template", "one", "two"}:
        raise ValueError("columns must be one of: template, one, two")
    if heading_before is not None and heading_before < 0:
        raise ValueError("heading_before must be nonnegative")
    if not re.fullmatch(r"[A-Za-z]*", numbering_prefix):
        raise ValueError("numbering_prefix must contain only ASCII letters")
    if bibliography_scope not in {"all", "new-only"}:
        raise ValueError("bibliography_scope must be one of: all, new-only")

    if citation_format not in {"template", "superscript", "bracketed"}:
        raise ValueError("citation_format must be one of: template, superscript, bracketed")
    template = Document(template_path)
    citation_superscript = citation_format == "superscript" or (
        citation_format == "template" and _template_uses_superscript_citations(template)
    )
    blocks = tuple(
        block
        for path in inputs
        for block in parse_markdown(path, strip_comments=not keep_comments)
    )
    if skip_images:
        blocks = tuple(block for block in blocks if block.kind != "image")
    bibliography = load_bibliography(bibliography_path) if bibliography_path else {}
    cited = _citation_keys(blocks)
    missing = sorted(set(cited) - set(bibliography))
    if missing:
        raise ValueError(f"Unknown citation key(s): {', '.join(missing)}")
    citation_base = _load_citation_base(citation_base_path) if citation_base_path else None
    used, mapping = _citation_plan(cited, bibliography, citation_base)
    rendered = _replace_block_citations(blocks, mapping, superscript=citation_superscript)
    abstract, body_blocks = _extract_abstract(rendered)
    if strip_level_one_headings:
        body_blocks = tuple(
            block for block in body_blocks
            if not (block.kind == "heading" and block.level == 1)
        )

    metadata = parse_metadata(metadata_path) if metadata_path else ManuscriptMetadata()
    resolved_title = title.strip() or metadata.title
    if not resolved_title:
        raise ValueError("Template assembly requires --title or a TITLE metadata value")

    sections_before = len(template.sections)
    geometry = _geometry(template)
    styles = _resolve_style_profile(template, discover_template_styles(template), style_profile)
    regions = _template_regions(template)
    prototypes = _template_prototypes(template, styles, regions)
    # Use the actual paragraph style carried by the reference body prototype;
    # MolCrysKit's body paragraphs are ``Normal`` while its semantic discovery
    # role may resolve to a custom style in another template.
    body_prototype = prototypes.paragraphs.get("body")
    if body_prototype is not None:
        body_ppr = body_prototype.find(qn("w:pPr"))
        body_style = body_ppr.find(qn("w:pStyle")) if body_ppr is not None else None
        if body_style is not None:
            styles = dict(styles)
            styles["body"] = body_style.get(qn("w:val"), styles["body"])
    body_regions = [region for region in regions if not region.figure]
    if len(body_regions) > 1 and body_regions[0].index == 0:
        front_region = body_regions.pop(0)
    else:
        front_region = None
    if not body_regions:
        body_regions = [front_region] if front_region is not None else [regions[0]]
    body_columns_before = _section_column_count(body_regions[0].section)

    def adjusted(section, *, apply_columns: bool = True):
        section = copy.deepcopy(section)
        line_numbers_node = section.find(qn("w:lnNumType"))
        if line_numbers == "off":
            if line_numbers_node is not None:
                section.remove(line_numbers_node)
        elif line_numbers == "on":
            if line_numbers_node is None:
                line_numbers_node = OxmlElement("w:lnNumType")
                section.append(line_numbers_node)
            line_numbers_node.set(qn("w:countBy"), "1")
            line_numbers_node.set(qn("w:distance"), "120")
            line_numbers_node.set(qn("w:restart"), "newPage")
        elif line_numbers_node is not None:
            # The reference omits a distance value.  LibreOffice then lets
            # three-digit labels touch the column text; the explicit gutter
            # keeps the inherited line-number contract without changing page
            # dimensions or margins.
            line_numbers_node.set(qn("w:distance"), "120")
        if apply_columns and columns in {"one", "two"}:
            col = section.find(qn("w:cols"))
            if col is None:
                col = OxmlElement("w:cols")
                section.append(col)
            if columns == "one":
                col.attrib.pop(qn("w:num"), None)
                col.attrib.pop(qn("w:space"), None)
            else:
                col.set(qn("w:num"), "2")
                col.set(qn("w:space"), col.get(qn("w:space"), "475"))
        return section

    body_columns_after = _section_column_count(adjusted(body_regions[0].section))
    front = []
    if prototypes.paragraphs.get("supporting_information") is not None:
        front.append(_new_metadata_paragraph(template, styles["body"], "Supporting Information", prototype=prototypes.paragraphs.get("supporting_information")))
    if include_title:
        front.append(_new_metadata_paragraph(template, styles["title"], resolved_title, prototype=prototypes.paragraphs.get("title")))
    if metadata.authors:
        front.append(_new_metadata_paragraph(template, styles["authors"], metadata.authors, prototype=prototypes.paragraphs.get("authors")))
    if metadata.affiliations:
        front.append(_new_metadata_paragraph(template, styles["affiliations"], metadata.affiliations, prototype=prototypes.paragraphs.get("affiliations")))
    for contact in metadata.contacts:
        front.append(_new_metadata_paragraph(template, styles["affiliations"], contact, prototype=prototypes.paragraphs.get("contacts")))
    if abstract:
        abstract_prototype = prototypes.paragraphs.get("abstract")
        # MolCrysKit carries the label and the abstract in one styled paragraph.
        front.append(_new_paragraph(template, styles["abstract"], f"**ABSTRACT:** {abstract}", prototype=abstract_prototype))

    # Build a fresh body while reusing the template's section properties.  The
    # sample text and sample images are discarded; their section geometry,
    # numbering, headers/footers and line-number settings are retained.
    output_nodes: list = []
    section_sources: list[int] = []
    if front_region is not None:
        output_nodes.extend(front)
        output_nodes.append(_section_break_node(adjusted(front_region.section)))
        section_sources.append(front_region.index)
    else:
        output_nodes.extend(front)

    groups = _figure_groups(body_blocks)
    figures: list[dict[str, object]] = []
    tables: list[dict[str, object]] = []
    body_index = 0
    figure_index = 0
    equation_index = 1
    table_index = 1
    for text_blocks, image, caption in groups:
        selected_body = body_regions[min(body_index, len(body_regions) - 1)]
        section = adjusted(selected_body.section)
        # _clone_rendered_block computes table widths from the document's
        # active section; the selected section is normally the same geometry,
        # and the explicit fit below handles differing templates.
        for block in text_blocks:
            nodes = _clone_rendered_block(
                block,
                template,
                styles,
                prototypes,
                equation_number=equation_index if block.kind == "equation" else None,
                table_number=table_index if block.kind == "table_caption" else None,
                number_prefix=numbering_prefix,
            )
            for node in nodes:
                _fit_tables(node, _section_width_twips(section))
            output_nodes.extend(nodes)
            if block.kind == "equation":
                equation_index += 1
            if block.kind == "table_caption":
                label = f"Table {numbering_prefix}{table_index}" if numbering_prefix else f"Table {table_index}"
                tables.append({"number": table_index, "label": label, "caption": block.text})
                table_index += 1
        if image is None:
            break
        if prototypes.figure_regions:
            figure_region = prototypes.figure_regions[min(figure_index, len(prototypes.figure_regions) - 1)]
            requested_columns = dict(image.options).get("columns")
            if requested_columns in {"single", "one"}:
                figure_section = _with_column_layout(figure_region.section, "one")
            elif requested_columns in {"double", "two"}:
                figure_section = _with_column_layout(figure_region.section, "two")
            elif requested_columns is None:
                figure_section = copy.deepcopy(figure_region.section)
            else:
                raise ValueError("Image columns marker must be one, two, single, or double")
            figure_section = adjusted(figure_section, apply_columns=False)
            previous_figure = prototypes.figure_regions[min(figure_index - 1, len(prototypes.figure_regions) - 1)] if figure_index > 0 else None
            needs_body_break = (
                bool(text_blocks)
                or previous_figure is None
                or previous_figure.index != figure_region.index
            )
            if needs_body_break:
                output_nodes.append(_section_break_node(section))
                section_sources.append(selected_body.index)
            output_nodes.extend(
                _clone_figure(
                    template,
                    image,
                    caption or image.text or f"Figure {figure_index + 1}.",
                    figure_region,
                    prototypes.paragraphs.get("caption"),
                    styles["caption"],
                    figure_index + 1,
                    section_width_twips=_section_width_twips(figure_section),
                    number_prefix=numbering_prefix,
                )
            )
            output_nodes.append(_section_break_node(figure_section))
            section_sources.append(figure_region.index)
            figures.append(
                {
                    "number": figure_index + 1,
                    "label": _numbered_caption(caption or image.text, figure_index + 1, numbering_prefix)[0],
                    "path": str(Path(image.path).resolve()),
                    "sha256": sha256_file(Path(image.path)),
                    "caption": caption or image.text,
                    "template_region": figure_region.index,
                    "placement_after_body_region": selected_body.index,
                    "columns": _section_column_count(figure_section),
                }
            )
            figure_index += 1
        else:
            # A blank single-section template has no full-width slot.  Keep
            # the image inline in the current column and do not invent a new
            # page geometry.
            output_nodes.extend(
                _clone_figure(
                    template,
                    image,
                    caption or image.text or f"Figure {figure_index + 1}.",
                    None,
                    prototypes.paragraphs.get("caption"),
                    styles["caption"],
                    figure_index + 1,
                    section_width_twips=_section_width_twips(section),
                    number_prefix=numbering_prefix,
                )
            )
            figures.append(
                {
                    "number": figure_index + 1,
                    "label": _numbered_caption(caption or image.text, figure_index + 1, numbering_prefix)[0],
                    "path": str(Path(image.path).resolve()),
                    "sha256": sha256_file(Path(image.path)),
                    "caption": caption or image.text,
                    "template_region": None,
                    "placement_after_body_region": selected_body.index,
                    "columns": _section_column_count(section),
                }
            )
            figure_index += 1
        body_index += 1

    reference_keys = tuple(key for key in used if bibliography_scope == "all" or citation_base is None or key not in citation_base)
    if include_metadata_back_matter and _is_filled_metadata(metadata.acknowledgement):
        output_nodes.append(_new_paragraph(template, styles["heading_1"], "ACKNOWLEDGMENTS", prototype=prototypes.paragraphs.get("heading_1"), uppercase=False))
        output_nodes.append(_new_paragraph(template, styles["body"], metadata.acknowledgement, prototype=prototypes.paragraphs.get("body")))
    if include_metadata_back_matter and _is_filled_metadata(metadata.author_contributions):
        output_nodes.append(_new_paragraph(template, styles["heading_1"], "AUTHOR CONTRIBUTIONS", prototype=prototypes.paragraphs.get("heading_1"), uppercase=False))
        output_nodes.append(_new_paragraph(template, styles["body"], metadata.author_contributions, prototype=prototypes.paragraphs.get("body")))
    if include_metadata_back_matter and _is_filled_metadata(metadata.code_availability):
        output_nodes.append(_new_paragraph(template, styles["heading_1"], "CODE AVAILABILITY", prototype=prototypes.paragraphs.get("heading_1"), uppercase=False))
        output_nodes.append(_new_paragraph(template, styles["body"], metadata.code_availability, prototype=prototypes.paragraphs.get("body")))
    if reference_keys:
        reference_prototype = prototypes.paragraphs.get("references_heading")
        if reference_prototype is None:
            reference_prototype = prototypes.paragraphs.get("heading_1")
        output_nodes.append(_new_paragraph(template, styles["references_heading"], "REFERENCES", prototype=reference_prototype, uppercase=False))
        for key in reference_keys:
            output_nodes.append(
                _new_paragraph(template, styles["reference"], f"{mapping[key]}.\t{bibliography[key]}", prototype=prototypes.paragraphs.get("reference"))
            )
    final_region = body_regions[min(body_index, len(body_regions) - 1)]
    # The final section belongs to the document body.  A body-level sectPr is
    # required for Word/LibreOffice to balance the last two-column region.
    output_nodes.append(adjusted(final_region.section))
    section_sources.append(final_region.index)

    heading_style_ids = {
        styles["heading_1"],
        styles["heading_2"],
        styles["heading_3"],
        styles["references_heading"],
    }
    for node in output_nodes:
        _override_heading_before(node, heading_style_ids, heading_before)
        _override_run_fonts(node, font_family, east_asia_font)
    body_element = template._element.body
    for child in list(body_element):
        body_element.remove(child)
    for node in output_nodes:
        body_element.append(node)
    template.core_properties.title = resolved_title
    template.core_properties.author = metadata.authors
    _update_fields(template)
    citation_converted = _materialize_superscript_citations(template) if citation_superscript else 0
    normalize_document_typography(template)
    output.parent.mkdir(parents=True, exist_ok=True)
    template.save(output)
    accept_docx_revisions(output)
    remove_docx_comments(output)
    _prune_images(output)
    convert_unicode_scripts_in_docx(output)
    selected_geometry = tuple(geometry[index] for index in section_sources if index < len(geometry))
    verification = verify_template_output(
        output,
        expected_sections=len(section_sources),
        expected_geometry=selected_geometry,
        expected_body_columns=body_columns_after,
    )
    verification["citation_format"] = "superscript" if citation_superscript else "bracketed"
    verification["citation_count"] = citation_converted
    return AssemblyResult(
        output=output,
        title=resolved_title,
        citation_map=mapping,
        used_citations=used,
        reference_keys=reference_keys,
        style_map=styles,
        template_sections_before=sections_before,
        template_sections_after=len(Document(output).sections),
        section_geometry=selected_geometry,
        body_columns_before=body_columns_before,
        body_columns_after=body_columns_after,
        column_layout=columns,
        verification=verification,
        font_family=font_family,
        east_asia_font=east_asia_font,
        style_profile=style_profile,
        line_numbers=line_numbers,
        contacts=metadata.contacts,
        include_title=include_title,
        strip_level_one_headings=strip_level_one_headings,
        heading_before=heading_before,
        section_sources=tuple(section_sources),
        figures=tuple(figures),
        tables=tuple(tables),
        numbering_prefix=numbering_prefix,
        bibliography_scope=bibliography_scope,
        include_metadata_back_matter=include_metadata_back_matter,
    )


def _descriptor(path: Path) -> dict[str, object]:
    return {"path": str(path.resolve()), "sha256": sha256_file(path), "size_bytes": path.stat().st_size}


def write_assembly_sidecars(
    result: AssemblyResult,
    *,
    inputs: Sequence[Path],
    template_path: Path,
    metadata_path: Path | None,
    bibliography_path: Path | None,
    command: Sequence[str],
    citation_base_path: Path | None = None,
) -> tuple[Path, Path]:
    import docx
    import lxml
    import PIL

    descriptors: dict[str, object] = {
        "markdown_inputs": [_descriptor(path) for path in inputs],
        "template": _descriptor(template_path),
    }
    if metadata_path is not None:
        descriptors["metadata"] = _descriptor(metadata_path)
    if bibliography_path is not None:
        descriptors["bibliography"] = _descriptor(bibliography_path)
    if citation_base_path is not None:
        descriptors["citation_base"] = _descriptor(citation_base_path)
    output = _descriptor(result.output)
    manifest = {
        "schema": "docforge.md2docx.template-manifest.v1",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "command": list(command),
        "inputs": descriptors,
        "output": output,
        "title": result.title,
        "metadata_fields": {
            "contacts": list(result.contacts),
            "include_back_matter": result.include_metadata_back_matter,
        },
        "rendering": {
            "font_family": result.font_family,
            "east_asia_font": result.east_asia_font,
            "style_profile": result.style_profile,
            "line_numbers": result.line_numbers,
            "include_title": result.include_title,
            "strip_level_one_headings": result.strip_level_one_headings,
            "heading_before": result.heading_before
        },
        "citation_map": dict(result.citation_map),
        "used_citations": list(result.used_citations),
        "reference_keys": list(result.reference_keys),
        "bibliography_scope": result.bibliography_scope,
        "style_map": dict(result.style_map),
        "figures": [dict(figure) for figure in result.figures],
        "tables": [dict(table) for table in result.tables],
        "numbering_prefix": result.numbering_prefix,
        "template_sections": {
            "before": result.template_sections_before,
            "after": result.template_sections_after,
            "geometry": [list(value) for value in result.section_geometry],
            "body_columns_before": result.body_columns_before,
            "body_columns_after": result.body_columns_after,
            "column_layout": result.column_layout,
            "source_indices": list(result.section_sources),
        },
        "verification": dict(result.verification),
        "environment": {
            "python": sys.version.split()[0],
            "python_docx": docx.__version__,
            "lxml": lxml.__version__,
            "pillow": PIL.__version__,
        },
    }
    manifest_path = result.output.with_suffix(".manifest.json")
    checksum_path = result.output.with_suffix(".sha256")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    checksum_path.write_text(f"{output['sha256']}  {result.output.name}\n", encoding="utf-8")
    return manifest_path, checksum_path
