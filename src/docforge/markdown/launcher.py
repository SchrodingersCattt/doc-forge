"""Markdown block parser and DOCX renderer.

This is the project-agnostic core of the former NSFC proposal converter.
Proposal-specific concerns (official heading lists, variant templates,
résumé section layout, the B0510 positive-scope language policy) live in
the consuming workflow, not here.
"""

from __future__ import annotations

import copy
import re
import tempfile
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

import lxml.etree as etree

from .blocks import Block, SectionSource
from ..tex.tokenize import tokenize_tex

__all__ = [
    "Block",
    "SectionSource",
    "parse_markdown",
    "add_block",
    "add_heading",
    "add_body_paragraph",
    "add_list_paragraph",
    "add_code_block",
    "add_equation",
    "add_table_caption",
    "add_quote",
    "add_image",
    "add_separator",
    "add_table",
    "setup_styles",
    "make_empty_doc",
    "override_document_fonts",
]

TABLE_SEPARATOR_RE = re.compile(r"^\s*:?-{3,}:?\s*$")
ORDERED_RE = re.compile(r"^(\s*)(\d+)\.\s+(.*)$")
BULLET_RE = re.compile(r"^(\s*)[-*+]\s+(.*)$")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
REFERENCE_RE = re.compile(r"^\[(\d+)\]\s+(.*)$")
TABLE_CAPTION_RE = re.compile(r"^Table\s*:\s*(.+)$", re.IGNORECASE)
IMAGE_RE = re.compile(r"^!\[(?P<caption>[^\]]*)\]\((?P<path>[^)]+)\)$")
COMMENT_RE = re.compile(r"<!--.*?-->", re.DOTALL)
BANG_COMMENT_RE = re.compile(r"^[!！](?:\s+|$)")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

INLINE_TOKEN_RE = re.compile(
    r"(<u>\*\*.*?\*\*</u>|<sup>[^<\n]*?</sup>|<sub>[^<\n]*?</sub>|<!--.*?-->|\*\*.*?\*\*|(?<!\*)\*[^*\n]+?\*(?!\*)|`[^`\n]+`|\$[^$\n]+?\$)"
)

# Unicode script ⇄ plain-text maps used when transferring Word runs.
UNICODE_SUBSCRIPT_MAP = str.maketrans(
    "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜₓₔ",
    "0123456789+-=()aehijklmnoprstxə",
)
UNICODE_SUPERSCRIPT_MAP = str.maketrans(
    "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂ",
    "0123456789+-=()abcdefghijklmnoprstuvwxyzABDEGHIJKLMNOPRTUVW",
)
# Word chemical formulae use the en dash for a superscript negative charge.
UNICODE_SUPERSCRIPT_MAP[ord("⁻")] = ord("–")
UNICODE_SCRIPT_CHARS = frozenset(
    chr(codepoint) for codepoint in (*UNICODE_SUBSCRIPT_MAP.keys(), *UNICODE_SUPERSCRIPT_MAP.keys())
)
EN_DASH_ENTITY_RE = re.compile(r"&(?:ndash|#8211|#x0*2013);", re.IGNORECASE)
BOND_HYPHEN_RE = re.compile(r"(?<![A-Za-z])([A-Z][a-z]?)-([A-Z][a-z]?)(?=(?:[0-9]|[\s,.;:)\]/]|$))")


# ── parsing ────────────────────────────────────────────────────────────────


def normalize_text(value: str) -> str:
    return re.sub(r"[\s\u00a0\u3000]+", "", value).rstrip("：:")


def normalize_typography(value: str) -> str:
    """Normalize dash entities and convert em dashes to en dashes."""
    value = EN_DASH_ENTITY_RE.sub("\u2013", value)
    value = value.replace("\u2014", "\u2013")
    return BOND_HYPHEN_RE.sub(r"\1–\2", value)


def clean_markdown(text: str, strip_comments: bool) -> str:
    if strip_comments:
        text = COMMENT_RE.sub("", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return normalize_typography(text)


def split_table_row(line: str) -> tuple[str, ...]:
    stripped = line.strip().strip("|")
    return tuple(cell.strip() for cell in stripped.split("|"))


def is_table_start(lines: list[str], index: int) -> bool:
    if index + 1 >= len(lines):
        return False
    first = lines[index].strip()
    second = lines[index + 1].strip()
    if not (first.startswith("|") and first.endswith("|") and second.startswith("|") and second.endswith("|")):
        return False
    return all(TABLE_SEPARATOR_RE.match(cell) for cell in split_table_row(second))


def _looks_like_table_row(line: str) -> bool:
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 2


def parse_markdown(path: Path, strip_comments: bool = True) -> list[Block]:
    """Parse a Markdown file into the block model (see ``Block``)."""
    text = clean_markdown(path.read_text(encoding="utf-8-sig"), strip_comments)
    lines = text.splitlines()
    blocks: list[Block] = []
    paragraph_lines: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph_lines
        if paragraph_lines:
            joined = " ".join(line.strip() for line in paragraph_lines if line.strip()).strip()
            if joined:
                blocks.append(Block("paragraph", joined))
            paragraph_lines = []

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            flush_paragraph()
            i += 1
            continue
        if BANG_COMMENT_RE.match(stripped) and not stripped.startswith("!["):
            flush_paragraph()
            if not strip_comments:
                blocks.append(Block("paragraph", f"[TODO: {stripped[1:].strip()}]"))
            i += 1
            continue
        heading = HEADING_RE.match(stripped)
        if heading:
            flush_paragraph()
            blocks.append(Block("heading", heading.group(2).strip(), level=len(heading.group(1))))
            i += 1
            continue
        image = IMAGE_RE.match(stripped)
        if image:
            flush_paragraph()
            image_path = (path.parent / image.group("path").strip()).resolve()
            raw_caption = image.group("caption").strip()
            caption_parts = [part.strip() for part in raw_caption.split("|")]
            options: list[tuple[str, str]] = []
            if len(caption_parts) > 1:
                kept = [caption_parts[0]]
                for part in caption_parts[1:]:
                    if "=" not in part:
                        kept.append(part)
                        continue
                    key, value = (piece.strip() for piece in part.split("=", 1))
                    options.append((key.lower(), value.lower()))
                raw_caption = " | ".join(kept)
            blocks.append(
                Block("image", text=raw_caption, path=str(image_path), options=tuple(options))
            )
            i += 1
            continue
        if stripped.startswith("```"):
            flush_paragraph()
            language = stripped[3:].strip()
            code: list[str] = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            i += 1 if i < len(lines) else 0
            blocks.append(Block("code", "\n".join(code), language=language))
            continue
        if stripped == "$$":
            flush_paragraph()
            equation: list[str] = []
            opening_line = i + 1
            i += 1
            while i < len(lines) and lines[i].strip() != "$$":
                equation.append(lines[i])
                i += 1
            if i >= len(lines):
                raise ValueError(f"Unclosed display equation in {path}:{opening_line}")
            if not any(line.strip() for line in equation):
                raise ValueError(f"Empty display equation in {path}:{opening_line}")
            i += 1
            blocks.append(Block("equation", "\n".join(equation).strip()))
            continue
        if is_table_start(lines, i):
            flush_paragraph()
            rows = [split_table_row(lines[i])]
            i += 2
            while i < len(lines):
                row = lines[i].strip()
                if not (row.startswith("|") and row.endswith("|")):
                    break
                rows.append(split_table_row(row))
                i += 1
            blocks.append(Block("table", rows=tuple(rows)))
            continue
        if _looks_like_table_row(line) and i + 1 < len(lines) and _looks_like_table_row(lines[i + 1]):
            raise ValueError(f"Malformed Markdown table in {path}:{i + 1}; separator row must use at least three hyphens per cell")
        ordered = ORDERED_RE.match(line)
        if ordered:
            flush_paragraph()
            blocks.append(
                Block("ordered", f"{ordered.group(2)}. {ordered.group(3)}", level=len(ordered.group(1)) // 2)
            )
            i += 1
            continue
        bullet = BULLET_RE.match(line)
        if bullet:
            flush_paragraph()
            blocks.append(Block("bullet", bullet.group(2), level=len(bullet.group(1)) // 2))
            i += 1
            continue
        if stripped.startswith(">"):
            flush_paragraph()
            blocks.append(Block("quote", stripped.lstrip("> ")))
            i += 1
            continue
        if stripped in {"---", "***", "___"}:
            flush_paragraph()
            blocks.append(Block("separator"))
            i += 1
            continue
        if REFERENCE_RE.match(stripped):
            flush_paragraph()
            blocks.append(Block("reference", stripped))
            i += 1
            continue
        table_caption = TABLE_CAPTION_RE.match(stripped)
        if table_caption:
            flush_paragraph()
            blocks.append(Block("table_caption", table_caption.group(1).strip()))
            i += 1
            continue
        paragraph_lines.append(line)
        i += 1
    flush_paragraph()
    return blocks


def make_empty_doc(*, configure_normal: bool = True) -> DocumentType:
    """A blank A4 document with the Heading/Table Grid styles configured."""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)
    setup_styles(doc, configure_normal=configure_normal)
    return doc


# ── inline rendering ───────────────────────────────────────────────────────


def set_run_font(run, *, chinese: str = "仿宋", latin: str = "Times New Roman", size: float = 12.0) -> None:
    run.font.name = latin
    run.font.size = Pt(size)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), chinese)


def override_document_fonts(
    doc: DocumentType,
    *,
    latin: str | None = None,
    east_asia: str | None = None,
) -> None:
    """Override fonts without changing run-level semantic formatting."""
    if not latin and not east_asia:
        return
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
            if latin:
                run.font.name = latin
                r_fonts.set(qn("w:ascii"), latin)
                r_fonts.set(qn("w:hAnsi"), latin)
                r_fonts.set(qn("w:cs"), latin)
            if east_asia:
                r_fonts.set(qn("w:eastAsia"), east_asia)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
                        if latin:
                            run.font.name = latin
                            r_fonts.set(qn("w:ascii"), latin)
                            r_fonts.set(qn("w:hAnsi"), latin)
                            r_fonts.set(qn("w:cs"), latin)
                        if east_asia:
                            r_fonts.set(qn("w:eastAsia"), east_asia)


def set_paragraph_spacing(
    paragraph, *, before: float = 0, after: float = 0, line: float = 1.5
) -> None:
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def apply_east_asian_line_break_rules(paragraph) -> None:
    """Keep Chinese closing punctuation off the start of wrapped lines."""
    p_pr = paragraph._p.get_or_add_pPr()
    for tag, value in (("kinsoku", "1"), ("overflowPunct", "1")):
        node = p_pr.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            p_pr.append(node)
        node.set(qn("w:val"), value)


def apply_base_format(paragraph, *, size: float = 12.0, first_indent: bool = False) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_spacing(paragraph, after=3, line=1.5)
    apply_east_asian_line_break_rules(paragraph)
    if first_indent:
        paragraph.paragraph_format.first_line_indent = Pt(size * 2)
    for run in paragraph.runs:
        set_run_font(run, size=size)


def set_highlight(run, fill: str = "FFF2CC") -> None:
    rpr = run._element.get_or_add_rPr()
    shading = rpr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        rpr.append(shading)
    shading.set(qn("w:fill"), fill)


def append_inline_math(
    paragraph, expression: str, *, bold_default: bool = False, size: float = 11.0
) -> None:
    r"""Render inline TeX as ordinary Word runs with true script formatting.

    Inline formulae are intentionally not OMML objects.  Variables retain the
    normal italic convention, ``\mathrm`` text and numerals remain upright,
    and ``^``/``_`` become Word ``w:vertAlign`` run properties.  Standalone
    display formulae use :func:`add_equation` and native OMML instead.
    """
    spans = tokenize_tex(f"${expression}$")
    binary_operators = {"+", "–", "=", "<", ">", "≤", "≥", "≠", "≈", "∼", "±", "∓", "→", "⟶", "←", "⟵"}
    normalized_text = [span.text for span in spans]
    for index, span in enumerate(spans):
        if span.superscript or span.subscript:
            continue
        stripped = span.text.strip()
        if stripped not in binary_operators or index == 0 or index + 1 >= len(spans):
            continue
        normalized_text[index - 1] = normalized_text[index - 1].rstrip()
        normalized_text[index] = f" {stripped} "
        normalized_text[index + 1] = normalized_text[index + 1].lstrip()
    for index, span in enumerate(spans):
        text = normalized_text[index]
        if not span.superscript and not span.subscript:
            text = re.sub(r"\s*(→|⟶|←|⟵|≤|≥|≠|≈|∼|±|∓|=|<|>)\s*", r" \1 ", text)
            text = re.sub(r"(?<=[A-Za-z0-9)])\s*([+–])\s*(?=[A-Za-z])", r" \1 ", text)
        run = paragraph.add_run(text)
        run.bold = bold_default or span.bold
        run.italic = span.italic
        run.font.subscript = span.subscript
        run.font.superscript = span.superscript
        set_run_font(run, size=size)
        if span.color:
            try:
                run.font.color.rgb = RGBColor.from_string(span.color)
            except ValueError:
                pass


def add_inline(
    paragraph, text: str, *, bold_default: bool = False, size: float = 11.0
) -> None:
    """Render inline Markdown tokens (bold, italic, code, math, comments)."""
    text = normalize_typography(text)
    position = 0
    for match in INLINE_TOKEN_RE.finditer(text):
        if match.start() > position:
            segment = text[position : match.start()]
            if paragraph.runs and paragraph.runs[-1].text.endswith(" ") and segment.startswith(" "):
                segment = segment.lstrip(" ")
            run = paragraph.add_run(segment)
            run.bold = bold_default
            set_run_font(run, size=size)
        token = match.group(0)
        if token.startswith("<u>**"):
            run = paragraph.add_run(token[5:-6])
            run.bold = True
            run.underline = True
            set_run_font(run, size=size)
        elif token.lower().startswith("<sup>"):
            run = paragraph.add_run(token[5:-6])
            run.bold = bold_default
            run.font.superscript = True
            set_run_font(run, size=size)
        elif token.lower().startswith("<sub>"):
            run = paragraph.add_run(token[5:-6])
            run.bold = bold_default
            run.font.subscript = True
            set_run_font(run, size=size)
        elif token.startswith("<!--"):
            inner = token[4:-3].strip()
            if inner.upper().startswith("TODO:"):
                inner = inner[5:].strip()
            run = paragraph.add_run(f"[TODO: {inner}]")
            set_run_font(run, chinese="宋体", latin="Times New Roman", size=size)
            run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        elif token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
            set_run_font(run, size=size)
        elif token.startswith("*"):
            run = paragraph.add_run(token[1:-1])
            run.italic = True
            set_run_font(run, size=size)
        elif token.startswith("`"):
            run = paragraph.add_run(token[1:-1])
            set_run_font(run, chinese="等线", latin="Consolas", size=max(size - 1, 9))
        elif token.startswith("$"):
            before = paragraph.runs[-1] if paragraph.runs else None
            first_math_run = len(paragraph.runs)
            append_inline_math(paragraph, token[1:-1], bold_default=bold_default, size=size)
            if before is not None and len(paragraph.runs) > first_math_run:
                first = paragraph.runs[first_math_run]
                if before.text.endswith(" ") and first.text.startswith(" "):
                    before.text = before.text.rstrip(" ")
        position = match.end()
    if position < len(text):
        remainder = text[position:]
        if paragraph.runs and paragraph.runs[-1].text.endswith(" ") and remainder.startswith(" "):
            remainder = remainder.lstrip(" ")
        run = paragraph.add_run(remainder)
        run.bold = bold_default
        set_run_font(run, size=size)


# ── block rendering ────────────────────────────────────────────────────────


def add_heading(doc: DocumentType, text: str, level: int):
    paragraph = doc.add_paragraph()
    word_level = min(level + 1, 9)
    paragraph.style = doc.styles[f"Heading {word_level}"]
    p_pr = paragraph._p.get_or_add_pPr()
    p_style = p_pr.find(qn("w:pStyle"))
    if p_style is None:
        p_style = OxmlElement("w:pStyle")
        p_pr.insert(0, p_style)
    p_style.set(qn("w:val"), doc.styles[f"Heading {word_level}"].style_id)
    outline_level = p_pr.find(qn("w:outlineLvl"))
    if outline_level is None:
        outline_level = OxmlElement("w:outlineLvl")
        p_pr.append(outline_level)
    outline_level.set(qn("w:val"), str(word_level - 1))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.keep_with_next = True
    if level == 1:
        size, before, after = 14.0, 12, 5
    elif level == 2:
        size, before, after = 13.0, 9, 3
    elif level == 3:
        size, before, after = 12.0, 6, 2
    else:
        size, before, after = 11.0, 4, 1
    set_paragraph_spacing(paragraph, before=before, after=after, line=1.2)
    add_inline(paragraph, text, bold_default=True, size=size)
    for run in paragraph.runs:
        run.bold = True
        run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        if level == 1:
            set_run_font(run, chinese="楷体", latin="Times New Roman", size=size)
        else:
            set_run_font(run, chinese="黑体", latin="Arial", size=size)
    return paragraph


def add_body_paragraph(doc: DocumentType, text: str, *, indent: bool = True):
    paragraph = doc.add_paragraph()
    add_inline(paragraph, text)
    apply_base_format(paragraph, first_indent=indent)
    return paragraph


def add_list_paragraph(doc: DocumentType, text: str, level: int, ordered: bool):
    paragraph = doc.add_paragraph()
    prefix = "" if ordered else "• "
    add_inline(paragraph, prefix + text)
    apply_base_format(paragraph, first_indent=False)
    paragraph.paragraph_format.left_indent = Cm(0.74 + 0.6 * level)
    paragraph.paragraph_format.first_line_indent = Cm(-0.37 if not ordered else 0)
    return paragraph


def add_code_block(doc: DocumentType, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(paragraph, before=3, after=3, line=1.0)
    paragraph.paragraph_format.left_indent = Cm(0.5)
    run = paragraph.add_run(text)
    set_run_font(run, chinese="等线", latin="Consolas", size=9)
    set_highlight(run, "F2F2F2")


def add_equation(
    doc: DocumentType,
    text: str,
    *,
    number: int | None = None,
    number_prefix: str = "",
) -> None:
    from ..math import latex_to_omml

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(paragraph, before=5, after=5, line=1.0)
    paragraph._p.append(latex_to_omml(text))
    if number is not None:
        # A right tab keeps the number at the edge of the receiving column;
        # template assembly retargets this tab to the actual column width.
        paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
        tab = paragraph.add_run("\t")
        set_run_font(tab, size=11.0)
        marker = paragraph.add_run(f"({number_prefix}{number})")
        set_run_font(marker, size=11.0)


def add_table_caption(
    doc: DocumentType,
    text: str,
    *,
    number: int | None = None,
    number_prefix: str = "",
) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(paragraph, before=3, after=2, line=1.0)
    label = f"Table {number_prefix}{number}. " if number is not None else "Table. "
    add_inline(paragraph, label + text, bold_default=False, size=10.0)
    for run in paragraph.runs:
        set_run_font(run, size=10.0)


def add_quote(doc: DocumentType, text: str) -> None:
    paragraph = doc.add_paragraph()
    add_inline(paragraph, text)
    apply_base_format(paragraph, first_indent=False)
    paragraph.paragraph_format.left_indent = Cm(0.74)
    paragraph.paragraph_format.right_indent = Cm(0.74)
    for run in paragraph.runs:
        run.italic = True


def add_image(doc: DocumentType, path: str, caption: str) -> None:
    image_path = Path(path)
    if not image_path.exists():
        raise FileNotFoundError(f"Markdown image does not exist: {image_path}")
    figure = doc.add_paragraph()
    figure.alignment = WD_ALIGN_PARAGRAPH.CENTER
    figure.paragraph_format.keep_with_next = True
    set_paragraph_spacing(figure, before=5, after=2, line=1.0)
    figure.add_run().add_picture(str(image_path), width=Cm(14.5))
    if caption:
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.keep_with_next = False
        set_paragraph_spacing(paragraph, before=0, after=5, line=1.0)
        add_inline(paragraph, caption, size=10.0)
        for run in paragraph.runs:
            set_run_font(run, size=10.0)


def add_separator(doc: DocumentType) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(2)
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def set_cell_text(cell, text: str, bold: bool = False) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    add_inline(paragraph, text, bold_default=bold)


def set_cell_margins(cell, top: int = 80, start: int = 80, bottom: int = 80, end: int = 80) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def set_table_three_line_borders(table) -> None:
    """Apply a journal-style three-line table with no fill or vertical rules."""
    table_properties = table._tbl.tblPr
    borders = table_properties.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_properties.append(borders)
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            borders.append(node)
        if side in {"top", "bottom"}:
            node.set(qn("w:val"), "single")
            node.set(qn("w:sz"), "10")
            node.set(qn("w:space"), "0")
            node.set(qn("w:color"), "000000")
        else:
            node.set(qn("w:val"), "nil")
            for attribute in ("sz", "space", "color"):
                node.attrib.pop(qn(f"w:{attribute}"), None)
    if not table.rows:
        return
    for cell in table.rows[0].cells:
        cell_properties = cell._tc.get_or_add_tcPr()
        cell_borders = cell_properties.find(qn("w:tcBorders"))
        if cell_borders is None:
            cell_borders = OxmlElement("w:tcBorders")
            cell_properties.append(cell_borders)
        bottom = cell_borders.find(qn("w:bottom"))
        if bottom is None:
            bottom = OxmlElement("w:bottom")
            cell_borders.append(bottom)
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def add_table(doc: DocumentType, rows: tuple[tuple[str, ...], ...]) -> None:
    if not rows:
        return
    width = max(len(row) for row in rows)
    table = doc.add_table(rows=len(rows), cols=width)
    table.style = "Table Grid"
    table.autofit = True
    for row_index, row_data in enumerate(rows):
        row = table.rows[row_index]
        for col_index in range(width):
            value = row_data[col_index] if col_index < len(row_data) else ""
            set_cell_text(row.cells[col_index], value, bold=row_index == 0)
            set_cell_margins(row.cells[col_index])
            for paragraph in row.cells[col_index].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_paragraph_spacing(paragraph, after=0, line=1.05)
                for run in paragraph.runs:
                    set_run_font(run, size=9.5)
        if row_index == 0:
            set_repeat_table_header(row)
    set_table_three_line_borders(table)
    doc.add_paragraph()


def add_block(
    doc: DocumentType,
    block: Block,
    *,
    equation_number: int | None = None,
    table_number: int | None = None,
    number_prefix: str = "",
) -> None:
    if block.kind == "heading":
        add_heading(doc, block.text, block.level)
    elif block.kind == "paragraph":
        add_body_paragraph(doc, block.text)
    elif block.kind == "ordered":
        add_list_paragraph(doc, block.text, block.level, ordered=True)
    elif block.kind == "bullet":
        add_list_paragraph(doc, block.text, block.level, ordered=False)
    elif block.kind == "reference":
        add_body_paragraph(doc, block.text, indent=False)
    elif block.kind == "code":
        add_code_block(doc, block.text)
    elif block.kind == "equation":
        add_equation(doc, block.text, number=equation_number, number_prefix=number_prefix)
    elif block.kind == "table_caption":
        add_table_caption(doc, block.text, number=table_number, number_prefix=number_prefix)
    elif block.kind == "table":
        add_table(doc, block.rows)
    elif block.kind == "quote":
        add_quote(doc, block.text)
    elif block.kind == "image":
        add_image(doc, block.path, block.text)
    elif block.kind == "separator":
        add_separator(doc)
    else:
        raise ValueError(f"Unsupported block kind: {block.kind}")


# ── styles ─────────────────────────────────────────────────────────────────


def setup_styles(doc: DocumentType, *, configure_normal: bool = True) -> None:
    normal = doc.styles["Normal"]
    if configure_normal:
        normal.font.name = "Times New Roman"
        normal._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
        normal.font.size = Pt(12)
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.space_after = Pt(3)
    for level in range(1, 10):
        name = f"Heading {level}"
        try:
            style = doc.styles[name]
        except KeyError:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.base_style = normal
        style.hidden = False
        style.quick_style = True
        style.priority = level
        style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        p_pr = style._element.get_or_add_pPr()
        outline_level = p_pr.find(qn("w:outlineLvl"))
        if outline_level is None:
            outline_level = OxmlElement("w:outlineLvl")
            p_pr.append(outline_level)
        outline_level.set(qn("w:val"), str(level - 1))
    if "Table Grid" not in [style.name for style in doc.styles]:
        doc.styles.add_style("Table Grid", WD_STYLE_TYPE.TABLE)


# ── verification helpers (shared by CLI and tests) ─────────────────────────


def verify_navigation_headings(path: Path) -> None:
    """Require every Heading paragraph to retain a valid outline level."""
    document = Document(str(path))
    failures: list[str] = []
    for paragraph in document.paragraphs:
        if not paragraph.style.name.startswith("Heading"):
            continue
        p_pr = paragraph._p.pPr
        revision_properties = p_pr.find(qn("w:rPr")) if p_pr is not None else None
        if revision_properties is not None and revision_properties.find(qn("w:del")) is not None:
            # A deleted heading remains in the all-markup XML as an empty
            # paragraph carrying its former Heading style.  It is absent from
            # the accepted navigation tree and therefore has no outline-level
            # requirement in the final document.
            continue
        match = re.search(r"(\d+)$", paragraph.style.name)
        if match is None:
            failures.append(f"{paragraph.text!r}: unrecognized style {paragraph.style.name!r}")
            continue
        expected = int(match.group(1)) - 1
        outline = p_pr.find(qn("w:outlineLvl")) if p_pr is not None else None
        actual = int(outline.get(qn("w:val"))) if outline is not None else None
        if actual != expected:
            failures.append(
                f"{paragraph.text!r}: {paragraph.style.name} has outline level {actual}; expected {expected}"
            )
    if failures:
        raise RuntimeError("DOCX navigation heading validation failed:\n" + "\n".join(failures[:20]))


def _script_chunks(text: str) -> list[tuple[str, str]]:
    """Return consecutive normal/subscript/superscript chunks with baseline text."""
    chunks: list[tuple[str, str]] = []
    for character in text:
        codepoint = ord(character)
        if codepoint in UNICODE_SUBSCRIPT_MAP:
            mode, value = "subscript", chr(UNICODE_SUBSCRIPT_MAP[codepoint])
        elif codepoint in UNICODE_SUPERSCRIPT_MAP:
            mode, value = "superscript", chr(UNICODE_SUPERSCRIPT_MAP[codepoint])
        else:
            mode, value = "normal", character
        if chunks and chunks[-1][0] == mode:
            chunks[-1] = (mode, chunks[-1][1] + value)
        else:
            chunks.append((mode, value))
    return chunks


def _run_with_child(run_properties, child, mode: str):
    run = etree.Element(qn("w:r"))
    if run_properties is not None:
        properties = copy.deepcopy(run_properties)
        if mode != "normal":
            for marker in properties.findall(qn("w:vertAlign")):
                properties.remove(marker)
            marker = etree.SubElement(properties, qn("w:vertAlign"))
            marker.set(qn("w:val"), mode)
        run.append(properties)
    elif mode != "normal":
        properties = etree.SubElement(run, qn("w:rPr"))
        marker = etree.SubElement(properties, qn("w:vertAlign"))
        marker.set(qn("w:val"), mode)
    run.append(child)
    return run


def convert_unicode_scripts_in_docx(path: Path) -> dict[str, int]:
    """Convert Unicode script glyphs throughout a DOCX package atomically."""
    totals = {"subscript": 0, "superscript": 0}
    with zipfile.ZipFile(path) as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]
    temporary_name = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{path.stem}-scripts-",
            suffix=".docx",
            dir=path.parent,
            delete=False,
        ) as temporary:
            temporary_name = Path(temporary.name)
        with zipfile.ZipFile(temporary_name, "w") as target:
            for info, payload in entries:
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    root = etree.fromstring(payload)
                    converted = _convert_unicode_scripts_in_xml(root)
                    for mode in totals:
                        totals[mode] += converted[mode]
                    payload = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                target.writestr(info, payload)
        for attempt in range(10):
            try:
                temporary_name.replace(path)
                break
            except PermissionError:
                if attempt == 9:
                    raise
                time.sleep(0.5)
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot update script formatting because the DOCX is open or locked: {path}"
        ) from exc
    finally:
        if temporary_name is not None and temporary_name.exists():
            temporary_name.unlink()
    return totals


def _convert_unicode_scripts_in_xml(root) -> dict[str, int]:
    """Replace Unicode script glyphs in Word runs with true run formatting."""
    counts = {"subscript": 0, "superscript": 0}
    text_tags = {qn("w:t"), qn("w:delText")}
    for run in list(root.iter(qn("w:r"))):
        text_nodes = [child for child in run if child.tag in text_tags]
        if not any(any(character in UNICODE_SCRIPT_CHARS for character in (node.text or "")) for node in text_nodes):
            continue
        parent = run.getparent()
        if parent is None:
            continue
        run_properties = run.find(qn("w:rPr"))
        replacements = []
        for child in run:
            if child.tag == qn("w:rPr"):
                continue
            if child.tag not in text_tags:
                replacements.append(_run_with_child(run_properties, copy.deepcopy(child), "normal"))
                continue
            original = child.text or ""
            for mode, value in _script_chunks(original):
                text_node = copy.deepcopy(child)
                text_node.text = value
                if value[:1].isspace() or value[-1:].isspace():
                    text_node.set(XML_SPACE, "preserve")
                replacements.append(_run_with_child(run_properties, text_node, mode))
                if mode != "normal":
                    counts[mode] += len(value)
        position = parent.index(run)
        parent.remove(run)
        for replacement in replacements:
            parent.insert(position, replacement)
            position += 1
    # Reviewed baselines may already contain superscript ASCII hyphens from an
    # earlier conversion.  Normalize those runs across both final and redline text.
    for run in root.iter(qn("w:r")):
        properties = run.find(qn("w:rPr"))
        marker = properties.find(qn("w:vertAlign")) if properties is not None else None
        if marker is None or marker.get(qn("w:val")) != "superscript":
            continue
        for child in run:
            if child.tag in text_tags and child.text:
                child.text = child.text.replace("-", "–")
    return counts


def document_contains_em_dash(doc: DocumentType) -> bool:
    for part in doc.part.package.parts:
        element = getattr(part, "_element", None)
        if element is None:
            continue
        for tag in ("w:t", "m:t"):
            if any("\u2014" in (node.text or "") for node in element.iter(qn(tag))):
                return True
    return False


def document_contains_hyphenated_numeric_range(doc: DocumentType) -> bool:
    pattern = re.compile(r"\[[0-9]+-[0-9]+(?:,[0-9]+)?\]|\b(?:19|20)[0-9]{2}-(?:19|20)[0-9]{2}(?=年)")
    for part in doc.part.package.parts:
        element = getattr(part, "_element", None)
        if element is None:
            continue
        for tag in ("w:t", "m:t"):
            if any(pattern.search(node.text or "") for node in element.iter(qn(tag))):
                return True
    return False


def verify_image_relationships(path: Path) -> None:
    """Require every image relationship in the saved package to resolve."""
    check = Document(str(path))
    for blip in check._element.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        relationship = check.part.rels.get(rid) if rid else None
        if relationship is None or relationship.reltype != RT.IMAGE:
            raise RuntimeError(f"Generated DOCX contains an unresolved image relationship: {rid}")


def render_blocks_to_doc(
    blocks: Iterable[Block],
    *,
    title: str = "",
    equation_start: int = 1,
    number_prefix: str = "",
    heading_before: float | None = None,
    font_family: str | None = None,
    east_asia_font: str | None = None,
) -> DocumentType:
    """Render parsed blocks into a fresh styled document (standalone usage)."""
    doc = make_empty_doc()
    if title:
        add_heading(doc, title, 1)
    equation_number = equation_start
    table_number = 1
    for block in blocks:
        add_block(
            doc,
            block,
            equation_number=equation_number if block.kind == "equation" else None,
            table_number=table_number if block.kind == "table_caption" else None,
            number_prefix=number_prefix,
        )
        if block.kind == "heading" and heading_before is not None and doc.paragraphs:
            doc.paragraphs[-1].paragraph_format.space_before = Pt(max(0, heading_before))
        if block.kind == "equation":
            equation_number += 1
        if block.kind == "table_caption":
            table_number += 1
    override_document_fonts(doc, latin=font_family, east_asia=east_asia_font)
    return doc


# ── template preservation helpers ──────────────────────────────────────────
#
# These were originally part of the NSFC proposal workflow and are kept here
# because any template-driven Markdown→DOCX job needs them: locate a marker
# paragraph, remove a body range, splice in generated content while remapping
# image relationships and heading style IDs, and normalize typography across
# the whole package (including headers/footers).


def replace_paragraph_text_preserve_format(paragraph, text: str) -> None:
    """Replace text while preserving the template paragraph and first-run formatting."""
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def copy_run_format(source, target) -> None:
    """Copy only direct run formatting from a template run."""
    source_rpr = source._r.rPr
    if source_rpr is not None:
        target._r.insert(0, copy.deepcopy(source_rpr))


def replace_cell_text_preserve_format(cell, text: str) -> None:
    """Fill a retained template cell without rebuilding its paragraph formatting."""
    paragraph = next((p for p in cell.paragraphs if p.runs), cell.paragraphs[0])
    replace_paragraph_text_preserve_format(paragraph, text)
    for extra in cell.paragraphs:
        # python-docx may return a new Paragraph wrapper for the same XML node.
        # Comparing wrapper identity clears the paragraph we just populated.
        if extra._p is paragraph._p:
            continue
        for run in extra.runs:
            run.text = ""


def set_cell_run_font_only(cell, font_name: str) -> None:
    """Set every run font slot in a cell while preserving size and other formatting."""
    for run in cell._tc.iter(qn("w:r")):
        properties = run.find(qn("w:rPr"))
        if properties is None:
            properties = OxmlElement("w:rPr")
            run.insert(0, properties)
        fonts = properties.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            properties.insert(0, fonts)
        for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(f"w:{attribute}"), font_name)
        for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn(f"w:{attribute}"), None)


def find_abstract_value_cell(doc: DocumentType, label: str):
    """Return the retained template cell paired with an abstract row label."""
    for table in doc.tables:
        row = next(
            (item for item in table.rows if normalize_text(item.cells[0].text) == normalize_text(label)),
            None,
        )
        if row is not None:
            return row.cells[1]
    raise ValueError(f"Could not locate abstract field in template: {label}")


def verify_english_abstract_font(doc: DocumentType) -> None:
    """Require explicit Times New Roman in every textual English-abstract run."""
    cell = find_abstract_value_cell(doc, "英文摘要")
    textual_runs = [
        run
        for run in cell._tc.iter(qn("w:r"))
        if any((node.text or "") for tag in ("w:t", "w:delText") for node in run.iter(qn(tag)))
    ]
    if not textual_runs:
        raise RuntimeError("English abstract contains no text runs")
    expected = "Times New Roman"
    for run in textual_runs:
        properties = run.find(qn("w:rPr"))
        fonts = properties.find(qn("w:rFonts")) if properties is not None else None
        values = {
            attribute: fonts.get(qn(f"w:{attribute}")) if fonts is not None else None
            for attribute in ("ascii", "hAnsi", "eastAsia", "cs")
        }
        if any(value != expected for value in values.values()):
            raise RuntimeError(f"English abstract run has incorrect font slots: {values}")


def find_body_index(body, text: str) -> int:
    target = normalize_text(text)
    for index, child in enumerate(body.iterchildren()):
        if child.tag == qn("w:p") and normalize_text(paragraph_text(child)) == target:
            return index
    raise ValueError(f"Could not find template paragraph: {text}")


def paragraph_text(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t"))).strip()


def find_body_index_after(body, start: int, candidates: tuple[str, ...]) -> int:
    """Find the first retained-section marker after the report body starts."""
    targets = {normalize_text(value) for value in candidates}
    children = list(body.iterchildren())
    for index, child in enumerate(children[start + 1 :], start=start + 1):
        if child.tag == qn("w:p") and normalize_text(paragraph_text(child)) in targets:
            return index
    raise ValueError(f"Could not find a retained-section marker after body index {start}: {candidates}")


def remove_range(body, start: int, end: int) -> None:
    children = list(body.iterchildren())
    for child in children[start:end]:
        body.remove(child)


def remap_image_relationships(element, source: DocumentType, target: DocumentType) -> None:
    """Copy image parts referenced by cloned XML into the target DOCX package."""
    for blip in element.iter(qn("a:blip")):
        old_rid = blip.get(qn("r:embed"))
        if not old_rid:
            continue
        source_rel = source.part.rels.get(old_rid)
        if source_rel is None or source_rel.reltype != RT.IMAGE:
            raise ValueError(f"Generated image has an invalid relationship: {old_rid}")
        new_rid, _ = target.part.get_or_add_image(BytesIO(source_rel.target_part.blob))
        blip.set(qn("r:embed"), new_rid)


def remap_heading_style_ids(element, target: DocumentType) -> None:
    """Map temporary-document heading IDs onto the reviewed template styles."""
    for paragraph in element.iter(qn("w:p")):
        p_pr = paragraph.find(qn("w:pPr"))
        if p_pr is None:
            continue
        p_style = p_pr.find(qn("w:pStyle"))
        if p_style is None:
            continue
        source_id = p_style.get(qn("w:val"), "")
        match = re.fullmatch(r"Heading(\d+)", source_id)
        if match is None:
            continue
        level = int(match.group(1))
        p_style.set(qn("w:val"), target.styles[f"Heading {level}"].style_id)
        outline = p_pr.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            p_pr.append(outline)
        outline.set(qn("w:val"), str(level - 1))


def replace_report_and_resume(template: DocumentType, generated: DocumentType) -> None:
    """Splice generated body blocks between marker paragraphs in the template."""
    body = template._element.body
    start = find_body_index(body, "报告正文")
    end = find_body_index_after(
        body,
        start,
        ("依托单位推荐意见", "依托单位：中山大学", "附件信息"),
    )
    preserved_end = list(body.iterchildren())[end]
    remove_range(body, start, end)

    insert_index = list(body.iterchildren()).index(preserved_end)
    generated_children = list(generated._element.body.iterchildren())
    for child in generated_children:
        if child.tag == qn("w:sectPr"):
            continue
        clone = copy.deepcopy(child)
        remap_image_relationships(clone, generated, template)
        remap_heading_style_ids(clone, template)
        body.insert(insert_index, clone)
        insert_index += 1


def normalize_document_typography(doc: DocumentType) -> None:
    """Normalize generated and preserved template text, including headers/footers."""
    seen: set[int] = set()
    for part in doc.part.package.parts:
        element = getattr(part, "_element", None)
        if element is None or id(element) in seen:
            continue
        seen.add(id(element))
        for tag in ("w:t", "m:t"):
            for node in element.iter(qn(tag)):
                if node.text:
                    node.text = normalize_typography(node.text)


def accept_docx_revisions(path: Path) -> None:
    """Materialize the reviewed document's final view before clean generation."""
    with zipfile.ZipFile(path, "r") as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    discard = {"del", "moveFrom"}
    retain = {"ins", "moveTo"}
    change_history = {
        "rPrChange",
        "pPrChange",
        "sectPrChange",
        "tblPrChange",
        "tblGridChange",
        "trPrChange",
        "tcPrChange",
    }
    property_parents = {"rPr", "pPr", "sectPr", "tblPr", "trPr", "tcPr"}
    for name, payload in list(files.items()):
        if not (name.startswith("word/") and name.endswith(".xml")):
            continue
        try:
            root = etree.fromstring(payload)
        except etree.XMLSyntaxError:
            continue
        changed = False
        for node in list(root.iter()):
            local = etree.QName(node).localname
            parent = node.getparent()
            if parent is None or local not in discard:
                continue
            parent.remove(node)
            changed = True
        for node in list(root.iter()):
            local = etree.QName(node).localname
            parent = node.getparent()
            if parent is None or local not in retain:
                continue
            if etree.QName(parent).localname in property_parents:
                parent.remove(node)
            else:
                position = parent.index(node)
                for child in list(node):
                    node.remove(child)
                    parent.insert(position, child)
                    position += 1
                parent.remove(node)
            changed = True
        for node in list(root.iter()):
            if etree.QName(node).localname not in change_history or node.getparent() is None:
                continue
            node.getparent().remove(node)
            changed = True
        if name == "word/settings.xml":
            for marker in list(root.xpath("//*[local-name()='trackRevisions']")):
                if marker.getparent() is not None:
                    marker.getparent().remove(marker)
                    changed = True
        if changed:
            files[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in files.items():
            archive.writestr(name, payload)


def remove_docx_comments(path: Path) -> None:
    """Remove review metadata from the clean DOCX package.

    A reviewed baseline can carry modern comment metadata and the setting that
    enables revision tracking.  The clean output must start from the same
    reviewed layout while opening as an ordinary, comment-free document.
    """
    review_parts = {
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsExtensible.xml",
        "word/commentsIds.xml",
        "word/people.xml",
    }
    comment_tags = {qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")}
    with zipfile.ZipFile(path, "r") as archive:
        files = {name: archive.read(name) for name in archive.namelist() if name not in review_parts}
    for name, payload in list(files.items()):
        if not name.endswith(".xml") and not name.endswith(".rels"):
            continue
        try:
            root = etree.fromstring(payload)
        except etree.XMLSyntaxError:
            continue
        changed = False
        for node in list(root.iter()):
            if node.tag in comment_tags and node.getparent() is not None:
                node.getparent().remove(node)
                changed = True
        for relationship in list(root.xpath("//*[local-name()='Relationship']")):
            rel_type = relationship.get("Type", "").lower()
            target = relationship.get("Target", "").lower()
            if (
                ("comment" in rel_type or rel_type.endswith("/people") or "people.xml" in target)
                and relationship.getparent() is not None
            ):
                relationship.getparent().remove(relationship)
                changed = True
        if name == "[Content_Types].xml":
            for override in list(root.xpath("//*[local-name()='Override']")):
                part_name = override.get("PartName", "").lower()
                if ("comment" in part_name or part_name.endswith("/people.xml")) and override.getparent() is not None:
                    override.getparent().remove(override)
                    changed = True
        if name == "word/settings.xml":
            for marker in list(root.xpath("//*[local-name()='trackRevisions']")):
                if marker.getparent() is not None:
                    marker.getparent().remove(marker)
                    changed = True
        if changed:
            files[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in files.items():
            archive.writestr(name, payload)


def verify_clean_review_state(path: Path) -> None:
    """Require a clean package with accepted revisions and removed comments."""
    review_parts = {
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsExtensible.xml",
        "word/commentsIds.xml",
        "word/people.xml",
    }
    revision_names = {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "rPrChange",
        "pPrChange",
        "sectPrChange",
        "tblPrChange",
        "tblGridChange",
        "trPrChange",
        "tcPrChange",
        "cellIns",
        "cellDel",
        "cellMerge",
        "numberingChange",
    }
    review_markers = {"commentRangeStart", "commentRangeEnd", "commentReference", "trackRevisions"}
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        retained_parts = sorted(review_parts & names)
        if retained_parts:
            raise RuntimeError(f"Clean DOCX retained review parts: {retained_parts}")
        retained_nodes: list[tuple[str, str]] = []
        for name in names:
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            root = etree.fromstring(archive.read(name))
            for node in root.iter():
                local = etree.QName(node).localname
                if local in revision_names or local in review_markers:
                    retained_nodes.append((name, local))
        if retained_nodes:
            raise RuntimeError(f"Clean DOCX retained review nodes: {retained_nodes[:20]}")


def verify_retained_metadata(
    doc: DocumentType,
    metadata: dict[str, str],
    chinese: str,
    english: str,
    title: str,
) -> None:
    """Fail generation if retained basic-information or abstract fields were lost."""
    retained_text = "\n".join(
        cell.text
        for table in doc.tables
        for row in table.rows
        for cell in row.cells
    )
    required = {
        "项目名称": title,
        "姓名": metadata.get("姓名", ""),
        "性别": metadata.get("性别", ""),
        "出生年月": metadata.get("出生年月", ""),
        "民族": metadata.get("民族", ""),
        "国别/地区": metadata.get("国别/地区", ""),
        "博士入学年份": metadata.get("博士入学年份", ""),
        "博士类别": metadata.get("博士类别", ""),
        "博士年级": metadata.get("博士年级", ""),
        "专业": metadata.get("专业", ""),
        "英文名称": metadata.get("英文名称", ""),
        "申请代码1": metadata.get("申请代码1", ""),
        "研究方向": metadata.get("研究方向", ""),
        "中文关键词": metadata.get("中文关键词", ""),
        "英文关键词": metadata.get("英文关键词", ""),
        "中文摘要": chinese,
        "英文摘要": english,
    }
    missing = [label for label, value in required.items() if not value or value not in retained_text]
    if missing:
        raise RuntimeError(f"Generated DOCX lost retained metadata fields: {', '.join(missing)}")
