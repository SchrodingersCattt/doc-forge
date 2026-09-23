"""Project-agnostic TeX to DOCX conversion and validation."""
from __future__ import annotations
import os, re, tempfile, zipfile
from contextvars import ContextVar
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
from ..docxdiff import Package
from ..docxdiff.redline import W as DIFF_W, NS as DIFF_NS, _blocks, visible_text
from ..math import latex_to_omml
from ..output import validate_output_path
from .bib import CitationResolver, parse_bib, _compress_labels
from .tokenize import Span, TableCell, tokenize_tex, spans_to_plain, plain_tex
FONT_BODY="Times New Roman"
FONT_HEADING="Arial"
PT_BODY=Pt(10)
PT_CAPTION=Pt(9)
PT_REF=Pt(10)
PT_HALF_LINE=Pt(5)
DOCX_NBSP="\u00A0"

def _scan_labels(text: str, fig_offset: int = 0, tbl_offset: int = 0,
                  prefix: str = "") -> dict[str, str]:
    """Scan text for labels and return a key->display map."""
    body_m = re.search(r"\\begin\{document\}", text)
    body = text[body_m.end():] if body_m else text
    result: dict[str, str] = {}
    fig_n = fig_offset
    tbl_n = tbl_offset
    alg_n = 0
    ed_fig_n = 0
    ed_tbl_n = 0
    sec_n = 0
    subsec_n = 0
    in_extended_data = False
    for m in re.finditer(
        r"\\section\*?\{([^}]*)\}|\\subsection\*?\{([^}]*)\}|\\begin\{(figure|table|longtable|algorithm)\b|\\label\{([^}]*)\}",
        body
    ):
        if m.group(1):  # section
            heading = re.sub(r"\\[a-zA-Z]+\*?", "", m.group(1)).strip()
            if heading == "Extended Data":
                in_extended_data = True
            else:
                sec_n += 1
                subsec_n = 0
            continue
        if m.group(2):  # subsection
            subsec_n += 1
            continue
        env = m.group(3)
        if env == "figure":
            if in_extended_data:
                ed_fig_n += 1
            else:
                fig_n += 1
        elif env in ("table", "longtable"):
            if in_extended_data:
                ed_tbl_n += 1
            else:
                tbl_n += 1
        elif env == "algorithm":
            alg_n += 1
        elif m.group(4):
            key = m.group(4)
            if key.startswith("fig:"):
                result[key] = f"{ed_fig_n}" if key.startswith("fig:ed_") else f"{prefix}{fig_n}"
            elif key.startswith("tab:"):
                result[key] = f"{ed_tbl_n}" if key.startswith("tab:ed_") else f"{prefix}{tbl_n}"
            elif key.startswith("alg:"):
                result[key] = f"{prefix}{alg_n}"
            elif key.startswith("si:"):
                # SI labels: use section/subsection numbering
                if subsec_n > 0:
                    result[key] = f"{prefix}{sec_n}.{subsec_n}"
                else:
                    result[key] = f"{prefix}{sec_n}"
            elif key.startswith("sec:"):
                result[key] = f"{prefix}{sec_n}"
            else:
                result[key] = key
    return result


def _build_label_map(full_text: str, aux_text: str | None = None,
                     current_prefix: str = "", aux_prefix: str = "S") -> dict[str,str]:
    labels=_scan_labels(full_text, prefix=current_prefix)
    if aux_text:
        for key,value in _scan_labels(aux_text,prefix=aux_prefix).items(): labels.setdefault(key,value)
    return labels


def _normalize_docx_whitespace(text: str) -> str:
    return text.replace(DOCX_NBSP, " ")


def _find_matching_brace(s: str, start: int) -> int:
    assert s[start] == "{"
    depth, pos = 1, start + 1
    while pos < len(s) and depth:
        if s[pos] == "{": depth += 1
        elif s[pos] == "}": depth -= 1
        pos += 1
    return pos


def _match_cmd_with_braced_arg(line: str, prefix_re: str):
    match = re.match(prefix_re, line)
    if not match: return None
    brace_start = match.end() - 1
    end = _find_matching_brace(line, brace_start)
    return line[brace_start + 1:end - 1], line[end:]


# ═══════════════════════════════════════════════════════════════════════════
#  Docx paragraph helpers
# ═══════════════════════════════════════════════════════════════════════════

TEX_COLOR_MAP = {
    "red": "FF0000", "blue": "0000FF", "green": "008000",
    "yellow": "FFFF00", "black": "000000", "white": "FFFFFF",
    "gray": "808080", "orange": "FF8C00", "cyan": "00FFFF",
}

OOXML_HL_MAP = {
    "yellow": "yellow", "red": "red", "green": "green", "blue": "blue",
    "cyan": "cyan", "magenta": "magenta", "darkBlue": "darkBlue",
}


def _tex_color_to_hex(c: str) -> str:
    c = c.strip().lower()
    if c in TEX_COLOR_MAP:
        return TEX_COLOR_MAP[c]
    if "!" in c:
        parts = c.split("!")
        base = TEX_COLOR_MAP.get(parts[0], "000000")
        return base
    return "000000"


def _tex_hl_name(c: str) -> str:
    c = c.strip().lower()
    return OOXML_HL_MAP.get(c, "yellow")


def _apply_span(run, span: Span, font_size=PT_BODY, font_name=FONT_BODY):
    run.text = _normalize_docx_whitespace(span.text)
    run.font.size = font_size
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:eastAsia="{font_name}"/>')
    rPr.insert(0, rFonts)
    if span.bold:
        run.bold = True
    if span.italic:
        run.italic = True
    if span.superscript:
        run.font.superscript = True
    if span.subscript:
        run.font.subscript = True
    if span.color:
        run.font.color.rgb = RGBColor(
            int(span.color[0:2], 16),
            int(span.color[2:4], 16),
            int(span.color[4:6], 16),
        )
    if span.highlight:
        rPr.append(parse_xml(
            f'<w:highlight {nsdecls("w")} w:val="{span.highlight}"/>'
        ))


def _reference_bookmark_name(label: str) -> str:
    return "ref_" + re.sub(r"[^A-Za-z0-9_]", "_", label)


def add_bookmark(paragraph, name: str):
    """Add a Word bookmark around an existing paragraph."""
    context=_context()
    bookmark_id=str(context.bookmark_counter)
    context.bookmark_counter+=1

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bookmark_id)
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bookmark_id)

    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _append_superscript_run(paragraph, text: str, font_size=PT_BODY, font_name=FONT_BODY):
    run = paragraph.add_run(text)
    run.font.size = font_size
    run.font.name = font_name
    run.font.superscript = True
    return run


def _append_internal_hyperlink(paragraph, text: str, anchor: str,
                               font_size=PT_BODY, font_name=FONT_BODY,
                               superscript: bool = True):
    """Append an internal Word hyperlink run."""
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")

    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    r_fonts = OxmlElement("w:rFonts")
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_pr.append(r_fonts)

    size = OxmlElement("w:sz")
    size.set(qn("w:val"), str(int(font_size.pt * 2)))
    r_pr.append(size)

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    r_pr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)

    if superscript:
        vert = OxmlElement("w:vertAlign")
        vert.set(qn("w:val"), "superscript")
        r_pr.append(vert)

    run.append(r_pr)
    text_el = OxmlElement("w:t")
    text_el.text = text
    run.append(text_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def _append_citation(paragraph, keys_str: str, resolver: CitationResolver,
                     font_size=PT_BODY, font_name=FONT_BODY,
                     use_inherited: bool = True,
                     force_superscript: bool | None = None):
    for text, target_label in resolver.citation_parts(keys_str, use_inherited=use_inherited):
        is_superscript = True if force_superscript is None else force_superscript
        if target_label:
            _append_internal_hyperlink(
                paragraph,
                text,
                _reference_bookmark_name(target_label),
                font_size=font_size,
                font_name=font_name,
                superscript=is_superscript,
            )
        else:
            if is_superscript:
                _append_superscript_run(
                    paragraph,
                    text,
                    font_size=font_size,
                    font_name=font_name,
                )
            else:
                run = paragraph.add_run(text)
                run.font.size = font_size
                run.font.name = font_name


def add_rich_text(paragraph, tex: str, resolver: CitationResolver | None = None,
                  font_size=PT_BODY, font_name=FONT_BODY,
                  citations_superscript: bool = True):
    """Tokenize *tex* and append formatted runs to *paragraph*."""
    def append_tex(chunk: str):
        spans = tokenize_tex(chunk, resolve_ref=lambda key:_context().label_map.get(key,"??"))
        for sp in spans:
            run = paragraph.add_run()
            _apply_span(run, sp, font_size=font_size, font_name=font_name)

    if not resolver:
        append_tex(tex)
        return

    pos = 0
    for m in re.finditer(r"\\(cite|mainref)\{([^}]+)\}", tex):
        append_tex(tex[pos:m.start()])
        _append_citation(
            paragraph,
            m.group(2),
            resolver,
            font_size=font_size,
            font_name=font_name,
            use_inherited=m.group(1) == "mainref",
            force_superscript=citations_superscript,
        )
        pos = m.end()
    append_tex(tex[pos:])


def _set_justify(p):
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _is_invisible_tex_line(line: str) -> bool:
    """Return True for LaTeX structural commands that should not emit text."""
    return bool(re.fullmatch(
        r"(?:\\phantomsection\s*)?(?:\\label\{[^}]*\}\s*)+",
        line.strip(),
    ))


def _strip_latex_comments(text: str) -> str:
    """Remove unescaped LaTeX line comments while preserving escaped percent signs."""
    placeholder = "\x00PERCENT\x00"
    lines = []
    for line in text.splitlines():
        protected = line.replace(r"\%", placeholder)
        if "%" in protected:
            protected = protected[:protected.index("%")]
        lines.append(protected.replace(placeholder, r"\%"))
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════
#  Document-level style setup
# ═══════════════════════════════════════════════════════════════════════════

def setup_styles(doc: Document):
    style = doc.styles["Normal"]
    style.font.name = FONT_BODY
    style.font.size = PT_BODY
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for i in range(1, 4):
        h = doc.styles[f"Heading {i}"]
        h.font.name = FONT_HEADING
        h.font.color.rgb = RGBColor(0, 0, 0)
        h.font.bold = True
        h.font.size = Pt([0, 14, 12, 10][i])
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(0)
        h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    if "Title" in doc.styles:
        ts = doc.styles["Title"]
        ts.font.name = FONT_HEADING
        ts.font.size = Pt(16)
        ts.font.bold = True
        ts.font.color.rgb = RGBColor(0, 0, 0)
        ts.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if "Caption" in doc.styles:
        cap = doc.styles["Caption"]
        cap.font.name = FONT_BODY
        cap.font.size = PT_CAPTION
        cap.font.italic = False
        cap.font.color.rgb = RGBColor(0, 0, 0)
        cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


# ═══════════════════════════════════════════════════════════════════════════
#  High-level add_* helpers
# ═══════════════════════════════════════════════════════════════════════════

class ParaTracker:
    def __init__(self): self._first=True
    def heading_seen(self): self._first=True
    def get_indent(self) -> Cm | None:
        if self._first:
            self._first=False
            return None
        return Cm(0.74)

@dataclass
class ConversionContext:
    figure_dir: Path
    project_root: Path
    label_map: dict[str,str] = field(default_factory=dict)
    fig_counter: int=0
    tbl_counter: int=0
    ed_fig_counter: int=0
    ed_tbl_counter: int=0
    eq_counter: int=0
    alg_counter: int=0
    bookmark_counter: int=0
    tracker: ParaTracker = field(default_factory=ParaTracker)

_ACTIVE_CONTEXT: ContextVar[ConversionContext | None] = ContextVar("docforge_tex_context", default=None)

def _context() -> ConversionContext:
    context=_ACTIVE_CONTEXT.get()
    if context is None: raise RuntimeError("TeX rendering helper called outside a conversion")
    return context


def add_title(doc: Document, title_tex: str, si_mode: bool = False):
    # Handle multi-line titles (e.g., SI with "Supplementary Information \\ Main Title")
    # Split on \\ with optional spacing/size commands
    parts = re.split(r"\\\\\s*(?:\[[\d.]+em\])?\s*", title_tex)

    for part_idx, part in enumerate(parts):
        if not part.strip():
            continue
        # Remove size commands like \large, \Large, \LARGE, \huge, \Huge
        part_clean = re.sub(r"\\(large|Large|LARGE|huge|Huge|small|footnotesize|scriptsize|tiny|normalsize)\b\s*", "", part)
        if not part_clean.strip():
            continue

        p = doc.add_paragraph(style="Title")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # First part (e.g., "Supplementary Information") gets slightly smaller font
        font_size = Pt(14) if part_idx == 0 and len(parts) > 1 else Pt(16)
        is_si_header = si_mode and part_idx == 0
        add_rich_text(p, part_clean, font_size=font_size, font_name=FONT_HEADING)
        for run in p.runs:
            if is_si_header:
                run.bold = False
                run.italic = True
            else:
                run.bold = True


def add_authors(doc: Document, authors: list, affils: list[str]):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, a in enumerate(authors):
        if isinstance(a, tuple):
            name, is_corr = a
        else:
            name, is_corr = a, False
        if not name.strip():
            continue
        add_rich_text(p, name, font_size=Pt(11))
        if is_corr:
            r = p.add_run("*")
            r.font.size = Pt(11)
            r.font.name = FONT_BODY
            r.font.superscript = True
        if i < len(authors) - 1:
            r = p.add_run(", ")
            r.font.size = Pt(11)
            r.font.name = FONT_BODY
    for aff in affils:
        pa = doc.add_paragraph()
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_rich_text(pa, aff, font_size=Pt(9))
        for run in pa.runs:
            run.italic = True


def add_abstract(doc: Document, text: str, resolver: CitationResolver):
    p_head = doc.add_paragraph()
    _set_justify(p_head)
    p_head.paragraph_format.space_before = Pt(12)
    r = p_head.add_run("Abstract")
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = FONT_HEADING

    # Collapse blank lines inside abstract: join non-empty lines
    text = re.sub(r"\n\s*\n", " ", text)
    text = re.sub(r"\n", " ", text)
    text = re.sub(r"  +", " ", text).strip()

    pa = doc.add_paragraph()
    _set_justify(pa)
    pa.paragraph_format.first_line_indent = Cm(0)
    add_rich_text(pa, text, resolver)


def add_heading(doc: Document, title_tex: str, level: int = 1):
    p = doc.add_heading(level=level)
    add_rich_text(p, title_tex, font_size=Pt([0, 14, 12, 10][level]),
                  font_name=FONT_HEADING)
    for run in p.runs:
        run.bold = True
    _context().tracker.heading_seen()


def add_body_paragraph(doc: Document, text: str, resolver: CitationResolver,
                       bold_prefix_tex: str | None = None):
    text_clean = re.sub(r"\s+", " ", text).strip()
    if not text_clean:
        return
    p = doc.add_paragraph()
    _set_justify(p)
    indent = _context().tracker.get_indent()
    if _starts_lowercase_continuation(text_clean):
        p.paragraph_format.first_line_indent = Cm(0)
    elif indent is not None:
        p.paragraph_format.first_line_indent = indent
    else:
        p.paragraph_format.first_line_indent = Cm(0)

    if bold_prefix_tex:
        spans = tokenize_tex(bold_prefix_tex)
        for sp in spans:
            sp.bold = True
            run = p.add_run(sp.text)
            _apply_span(run, sp)
        r = p.add_run(" ")
        r.font.size = PT_BODY
        r.font.name = FONT_BODY

    add_rich_text(p, text_clean, resolver)


def _set_paragraph_border(p, *, top: bool = False, bottom: bool = False):
    """Apply simple paragraph borders used for algorithm blocks."""
    p_pr = p._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)

    for side, enabled in (("top", top), ("bottom", bottom)):
        if not enabled:
            continue
        border = p_bdr.find(qn(f"w:{side}"))
        if border is None:
            border = OxmlElement(f"w:{side}")
            p_bdr.append(border)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "8")
        border.set(qn("w:space"), "1")
        border.set(qn("w:color"), "000000")


def _algorithm_braced_command(line: str, command: str) -> tuple[str, str] | None:
    m = re.match(re.escape(command) + r"\{", line)
    if not m:
        return None
    brace_start = m.end() - 1
    brace_end = _find_matching_brace(line, brace_start)
    return line[brace_start + 1: brace_end - 1], line[brace_end:].strip()


def _algorithmic_lines(block: str) -> list[tuple[int | None, str, str]]:
    """Return (line_number, text, kind) rows from an algorithmic environment."""
    rows: list[tuple[int | None, str, str]] = []
    m = re.search(r"\\begin\{algorithmic\}(?:\[[^\]]*\])?(.*?)\\end\{algorithmic\}", block, re.DOTALL)
    if not m:
        return rows

    indent = 0
    line_no = 1
    for raw in m.group(1).splitlines():
        line = raw.strip()
        if not line or line.startswith("%") or _is_invisible_tex_line(line):
            continue

        def emit(text: str, level: int | None = None):
            nonlocal line_no
            if level is None:
                rows.append((None, text, "meta"))
                return
            rows.append((line_no, text, f"code:{max(level, 0)}"))
            line_no += 1

        if line.startswith(r"\Require"):
            emit(r"\textbf{Require:} " + line[len(r"\Require"):].strip())
            continue
        if line.startswith(r"\Ensure"):
            emit(r"\textbf{Ensure:} " + line[len(r"\Ensure"):].strip())
            continue
        if line.startswith(r"\State"):
            emit(line[len(r"\State"):].strip(), indent)
            continue

        parsed = _algorithm_braced_command(line, r"\For")
        if parsed:
            condition, _ = parsed
            emit(r"\textbf{for} " + condition + r" \textbf{do}", indent)
            indent += 1
            continue
        if line.startswith(r"\EndFor"):
            indent = max(indent - 1, 0)
            emit(r"\textbf{end for}", indent)
            continue

        parsed = _algorithm_braced_command(line, r"\While")
        if parsed:
            condition, _ = parsed
            emit(r"\textbf{while} " + condition + r" \textbf{do}", indent)
            indent += 1
            continue
        if line.startswith(r"\EndWhile"):
            indent = max(indent - 1, 0)
            emit(r"\textbf{end while}", indent)
            continue

        parsed = _algorithm_braced_command(line, r"\If")
        if parsed:
            condition, rest = parsed
            inline_end = r"\EndIf" in rest
            rest = rest.replace(r"\EndIf", "").strip()
            suffix = f" {rest}" if rest else ""
            emit(r"\textbf{if} " + condition + r" \textbf{then}" + suffix, indent)
            if inline_end:
                emit(r"\textbf{end if}", indent)
            else:
                indent += 1
            continue
        if line.startswith(r"\EndIf"):
            indent = max(indent - 1, 0)
            emit(r"\textbf{end if}", indent)
            continue

        emit(line, indent)

    return rows


def _algorithm_text_tex(tex: str) -> str:
    """Keep algorithm math as compact inline pseudocode text."""
    tex = tex.replace(r"\arg\min", r"\arg \min")
    return tex.replace("$", "")


def add_algorithm(doc: Document, block: str, resolver: CitationResolver,
                  float_prefix: str = ""):
    context=_context(); context.alg_counter+=1
    label=f"Algorithm {float_prefix}{context.alg_counter} "
    caption_tex = _extract_braced_arg(block, r"\caption")

    cap = doc.add_paragraph()
    _set_justify(cap)
    cap.paragraph_format.first_line_indent = Cm(0)
    cap.paragraph_format.space_before = PT_HALF_LINE
    cap.paragraph_format.space_after = Pt(0)
    r = cap.add_run(label)
    r.bold = True
    r.font.size = PT_CAPTION
    r.font.name = FONT_BODY
    if caption_tex:
        add_rich_text(cap, caption_tex, resolver, font_size=PT_CAPTION)

    rows = _algorithmic_lines(block)
    last_para = None
    for row_idx, (number, text, kind) in enumerate(rows):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        if row_idx == 0:
            _set_paragraph_border(p, top=True)
        if number is None:
            p.paragraph_format.first_line_indent = Cm(0)
            add_rich_text(p, _algorithm_text_tex(text), resolver, font_size=PT_CAPTION)
        else:
            indent_level = int(kind.split(":", 1)[1])
            p.paragraph_format.left_indent = Cm(0.55 + 0.45 * indent_level)
            p.paragraph_format.first_line_indent = Cm(-0.55)
            num_run = p.add_run(f"{number}: ")
            num_run.font.size = Pt(8)
            num_run.font.name = FONT_BODY
            add_rich_text(p, _algorithm_text_tex(text), resolver, font_size=PT_CAPTION)
        last_para = p

    if last_para is not None:
        _set_paragraph_border(last_para, bottom=True)


def _normalize_display_math_source(s: str) -> str:
    s = s.strip()
    if s.startswith(r"\["):
        s = s[2:]
    if s.endswith(r"\]"):
        s = s[:-2]
    s = re.sub(r"\\begin\{equation\*?\}", "", s)
    s = re.sub(r"\\end\{equation\*?\}", "", s)
    s = re.sub(r"\\begin\{(?:array|aligned|align)\}(?:\{[^}]*\})?", "", s)
    s = re.sub(r"\\end\{(?:array|aligned|align)\}", "", s)
    s = re.sub(r"\\left\s*([(\[|.])", r"\1", s)
    s = re.sub(r"\\right\s*([)\]|.])", r"\1", s)
    s = re.sub(r"\\vdet\b", lambda _: r"V_{\mathrm{det}}", s)
    s = re.sub(r"\\etasq\b", lambda _: r"\eta^{2}", s)
    s = s.replace(r"\,", " ")
    s = re.sub(r"\\qquad\b", "    ", s)
    s = re.sub(r"\\quad\b", "  ", s)
    return s.strip()


def _display_math_rows(math_tex: str) -> list[str]:
    s = _normalize_display_math_source(math_tex)
    rows = re.split(r"\\\\", s)
    clean_rows = []
    for row in rows:
        row = re.sub(r"\s*&\s*=\s*&\s*", " \u2009=\u2009 ", row)
        row = row.replace("&", " ")
        row = re.sub(r"\s+", " ", row).strip(" ,;")
        if row:
            clean_rows.append(row)
    return clean_rows


def _starts_lowercase_continuation(text: str) -> bool:
    match=re.search(r"[A-Za-z]",text)
    return bool(match and match.group(0).islower())


def add_display_math(doc: Document, math_tex: str, resolver: CitationResolver, equation_prefix: str=""):
    rows=_display_math_rows(math_tex)
    if not rows: return
    context=_context(); context.eq_counter+=1
    label=f"({equation_prefix}{context.eq_counter})"
    for index,row in enumerate(rows):
        p=doc.add_paragraph(); p.paragraph_format.first_line_indent=Cm(0)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(8),WD_TAB_ALIGNMENT.CENTER)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16),WD_TAB_ALIGNMENT.RIGHT)
        run=p.add_run("\t"); run.font.size=PT_BODY; run.font.name=FONT_BODY
        p._p.append(latex_to_omml(row))
        if index==len(rows)-1:
            run=p.add_run(f"\t{label}"); run.font.size=PT_BODY; run.font.name=FONT_BODY


# ═══════════════════════════════════════════════════════════════════════════
#  Figure insertion
# ═══════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════
#  Figure insertion
# ═══════════════════════════════════════════════════════════════════════════

def _find_figure(name: str) -> Path | None:
    base = name.split(".")[0] if "." in name else name
    for prefix in ("figures/", "figures\\"):
        if base.startswith(prefix):
            base = base[len(prefix):]
    context=_context()
    for ext in (".png", ".jpg", ".jpeg"):
        path=context.figure_dir/f"{base}{ext}"
        if path.exists(): return path
    if "." in name:
        path=context.project_root/name
        if path.exists() and path.suffix.lower() in (".png", ".jpg", ".jpeg"): return path
    return None


def add_figure(doc: Document, fig_path: str | list[str], caption_tex: str,
               resolver: CitationResolver, width_inches: float = 6.0,
               extended: bool = False, float_prefix: str = ""):
    context=_context()
    if extended:
        context.ed_fig_counter+=1
        label=f"Extended Data Fig. {context.ed_fig_counter}. "
    else:
        context.fig_counter+=1
        label=f"Figure {float_prefix}{context.fig_counter}. "
    fig_paths = [fig_path] if isinstance(fig_path, str) else fig_path
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    for idx, one_path in enumerate(fig_paths):
        img = _find_figure(one_path)
        if img:
            run = p.add_run()
            run.add_picture(str(img), width=Inches(width_inches))
            if idx < len(fig_paths) - 1:
                run.add_break()
        else:
            run = p.add_run(f"[Figure: {one_path} - not found as PNG/JPG]")
            run.font.color.rgb = RGBColor(0xCC, 0, 0)
            run.italic = True

    cp = doc.add_paragraph()
    _set_justify(cp)
    cp.paragraph_format.first_line_indent = Cm(0)
    cp.paragraph_format.space_after = PT_HALF_LINE
    r = cp.add_run(label)
    r.bold = True
    r.font.size = PT_CAPTION
    r.font.name = FONT_BODY
    if caption_tex:
        caption_tex = re.sub(r"\n\s*", " ", caption_tex)
        caption_tex = re.sub(r"  +", " ", caption_tex).strip()
        add_rich_text(cp, caption_tex, resolver, font_size=PT_CAPTION)


# ═══════════════════════════════════════════════════════════════════════════
#  Table insertion
# ═══════════════════════════════════════════════════════════════════════════

def _cell_text(cell: TableCell | str) -> str:
    return cell.text if isinstance(cell, TableCell) else str(cell)


def _extract_braced_arg_at(s: str, brace_pos: int) -> tuple[str, int] | None:
    """Extract a brace-balanced argument starting at *brace_pos*."""
    if brace_pos >= len(s) or s[brace_pos] != "{":
        return None
    end = _find_matching_brace(s, brace_pos)
    return s[brace_pos + 1:end - 1], end


def _parse_multirow_cell(cell_tex: str) -> TableCell:
    """Parse a cell, preserving explicit LaTeX \\multirow row spans."""
    cell_tex = cell_tex.strip()
    m = re.search(r"\\multirow\s*", cell_tex)
    if not m:
        return TableCell(cell_tex)

    pos = m.end()
    args: list[str] = []
    for _ in range(3):
        while pos < len(cell_tex) and cell_tex[pos].isspace():
            pos += 1
        extracted = _extract_braced_arg_at(cell_tex, pos)
        if extracted is None:
            return TableCell(cell_tex)
        arg, pos = extracted
        args.append(arg)

    try:
        rowspan = max(1, int(args[0].strip()))
    except ValueError:
        rowspan = 1

    merged_text = (cell_tex[:m.start()] + args[2] + cell_tex[pos:]).strip()
    return TableCell(merged_text, rowspan=rowspan)


def _parse_multicolumn_cell(cell_tex: str) -> TableCell:
    """Parse a cell, preserving explicit LaTeX \\multicolumn column spans."""
    cell_tex = cell_tex.strip()
    m = re.search(r"\\multicolumn\s*", cell_tex)
    if not m:
        return TableCell(cell_tex)

    pos = m.end()
    args: list[str] = []
    for _ in range(3):
        while pos < len(cell_tex) and cell_tex[pos].isspace():
            pos += 1
        extracted = _extract_braced_arg_at(cell_tex, pos)
        if extracted is None:
            return TableCell(cell_tex)
        arg, pos = extracted
        args.append(arg)

    try:
        colspan = max(1, int(args[0].strip()))
    except ValueError:
        colspan = 1

    merged_text = (cell_tex[:m.start()] + args[2] + cell_tex[pos:]).strip()
    return TableCell(merged_text, colspan=colspan)


def _parse_table_cell(cell_tex: str) -> TableCell:
    cell = _parse_multicolumn_cell(cell_tex)
    parsed = _parse_multirow_cell(cell.text)
    return TableCell(parsed.text, rowspan=parsed.rowspan, colspan=cell.colspan)


def _parse_table_cells(line: str) -> list[TableCell]:
    cells = []
    for raw in line.split("&"):
        raw = raw.strip()
        cells.append(_parse_table_cell(raw))
    return cells


def _parse_tabular_rows(body: str) -> list[list[TableCell]]:
    m_begin = re.search(r"\\begin\{tabular\}", body)
    m_end = re.search(r"\\end\{tabular\}", body)
    if m_begin and m_end:
        start = m_begin.end()
        while start < len(body) and body[start].isspace():
            start += 1
        if start < len(body) and body[start] == "{":
            start = _find_matching_brace(body, start)
        inner = body[start:m_end.start()]
    else:
        inner = body
    inner = _strip_latex_comments(inner)
    inner = re.sub(r"\\toprule|\\midrule|\\bottomrule|\\hline", "", inner)
    inner = re.sub(r"\\centering", "", inner)
    inner = re.sub(r"\\label\{[^}]*\}", "", inner)
    lines = [l.strip() for l in inner.split("\\\\") if l.strip()]
    rows = []
    for line in lines:
        line = re.sub(r"\\(?:addlinespace|smallskip|medskip|bigskip)\b(?:\[[^\]]*\])?", "", line).strip()
        if not line:
            continue
        cells = _parse_table_cells(line)
        if any(_cell_text(c).strip() for c in cells):
            rows.append(cells)
    return rows


def _parse_longtable_rows(body: str) -> list[list[TableCell]]:
    inner = body

    # Remove \caption{...} with brace-balanced extraction
    while r"\caption{" in inner:
        m = re.search(r"\\caption\{", inner)
        if not m:
            break
        start = m.start()
        pos = m.end()
        depth = 1
        while pos < len(inner) and depth > 0:
            if inner[pos] == "{": depth += 1
            elif inner[pos] == "}": depth -= 1
            pos += 1
        inner = inner[:start] + inner[pos:]

    # Find and extract the first-page header (before \endfirsthead)
    first_header = ""
    m_firsthead = re.search(r"\\endfirsthead", inner)
    if m_firsthead:
        # Find the start of the header (after \label)
        start_pos = 0
        m_label = re.search(r"\\label\{[^}]*\}\s*\\\\\s*", inner)
        if m_label:
            start_pos = m_label.end()
        first_header = inner[start_pos:m_firsthead.start()]
        # Remove everything up to and including \endfirsthead
        inner = inner[m_firsthead.end():]

    # Remove continuation header (everything before \endhead, which includes "continued")
    m_endhead = re.search(r"\\endhead", inner)
    if m_endhead:
        inner = inner[m_endhead.end():]

    # Remove footer markers
    for tag in (r"\\endfoot", r"\\endlastfoot"):
        inner = re.sub(tag, "", inner)

    # Clean up table environment and formatting commands
    inner = re.sub(r"\\begin\{longtable\}\{(?:[^{}]|\{[^{}]*\})*\}", "", inner)
    first_header = re.sub(r"\\begin\{longtable\}\{(?:[^{}]|\{[^{}]*\})*\}", "", first_header)

    for tag in (r"\\end\{longtable\}",
                r"\\toprule", r"\\midrule", r"\\bottomrule", r"\\hline",
                r"\\label\{[^}]*\}", r"\\thetable\{[^}]*\}", r"\\tablename",
                r"\\centering", r"\\footnotesize", r"\\addlinespace(?:\[[^\]]*\])?"):
        inner = re.sub(tag, "", inner)
        first_header = re.sub(tag, "", first_header)

    # Multicolumn and multirow are preserved as span metadata while parsing cells.
    inner = _strip_latex_comments(inner)
    first_header = _strip_latex_comments(first_header)

    # Parse first header rows
    rows = []
    if first_header.strip():
        header_lines = [l.strip() for l in first_header.split("\\\\") if l.strip()]
        for line in header_lines:
            if line.startswith("%") or re.match(r"^\s*$", line):
                continue
            cells = _parse_table_cells(line)
            if any(_cell_text(c).strip() for c in cells):
                rows.append(cells)

    # Parse data rows
    lines = [l.strip() for l in inner.split("\\\\") if l.strip()]
    for line in lines:
        if line.startswith("%") or re.match(r"^\s*$", line):
            continue
        cells = _parse_table_cells(line)
        if any(_cell_text(c).strip() for c in cells):
            rows.append(cells)

    return rows


def _set_cell_border(cell, top=None, bottom=None):
    """Set individual cell borders. *top*/*bottom* are tuples (size_eighths, style)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        tcPr.append(borders)
    for side, val in [("top", top), ("bottom", bottom)]:
        if val is None:
            continue
        sz, style = val
        el = borders.find(qn(f"w:{side}"))
        if el is not None:
            borders.remove(el)
        borders.append(parse_xml(
            f'<w:{side} {nsdecls("w")} w:val="{style}" w:sz="{sz}" w:space="0" w:color="000000"/>'
        ))


def _clear_all_borders(tbl):
    """Remove all borders from a table and its cells."""
    tblPr = tbl._tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is not None:
        tblPr.remove(borders)
    none_border = (
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0"/>'
        '<w:left w:val="none" w:sz="0" w:space="0"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0"/>'
        '<w:right w:val="none" w:sz="0" w:space="0"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0"/>'
        '</w:tblBorders>'
    )
    tblPr.append(parse_xml(none_border))


def _combine_multilevel_header(rows: list[list[TableCell]]) -> list[list[TableCell]]:
    """Flatten common LaTeX multicolumn header pairs for Word tables."""
    if len(rows) < 2:
        return rows
    ncols = max(_row_width(r) for r in rows)
    top, sub = rows[0], rows[1]
    if _row_width(top) >= ncols or len(top) < 2 or not sub or _cell_text(sub[0]).strip():
        return rows
    group_count = len(top) - 1
    remaining = ncols - 1
    if group_count <= 0 or remaining % group_count:
        return rows
    span = remaining // group_count
    combined = [TableCell(_cell_text(top[0]))]
    pos = 1
    for group in top[1:]:
        for subcell in sub[pos:pos + span]:
            combined.append(TableCell(f"{_cell_text(group)} {_cell_text(subcell)}".strip()))
        pos += span
    if len(combined) != ncols:
        return rows
    return [combined] + rows[2:]


def _row_width(row: list[TableCell]) -> int:
    return sum(max(1, cell.colspan) for cell in row)


def _iter_logical_cells(row: list[TableCell]):
    col = 0
    for cell in row:
        yield col, cell
        col += max(1, cell.colspan)


def _cell_at_col(row: list[TableCell], col_idx: int) -> TableCell | None:
    for start_col, cell in _iter_logical_cells(row):
        if start_col <= col_idx < start_col + max(1, cell.colspan):
            return cell
    return None


def _clear_cell_content(cell):
    for p in cell.paragraphs:
        p.clear()


def _merge_vertical_cells(tbl, start_row: int, end_row: int, col: int):
    """Merge a vertical cell range while preserving only the top cell content."""
    if end_row <= start_row:
        return
    for row_idx in range(start_row + 1, end_row + 1):
        _clear_cell_content(tbl.cell(row_idx, col))
    merged = tbl.cell(start_row, col).merge(tbl.cell(end_row, col))
    merged.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _apply_explicit_rowspans(tbl, rows: list[list[TableCell]], header_rows: int) -> set[tuple[int, int]]:
    """Apply LaTeX \\multirow spans and return occupied continuation cells."""
    occupied: set[tuple[int, int]] = set()
    nrows = len(rows)
    for i, row in enumerate(rows):
        if i < header_rows:
            continue
        for j, cell in _iter_logical_cells(row):
            if cell.rowspan <= 1:
                continue
            end = min(nrows - 1, i + cell.rowspan - 1)
            _merge_vertical_cells(tbl, i, end, j)
            for r in range(i, end + 1):
                occupied.add((r, j))
    return occupied


def _is_auto_merge_column(header: str) -> bool:
    header_norm = re.sub(r"[^A-Za-z]+", " ", header).strip().lower()
    safe_headers = {
        "level", "group", "category", "class", "type", "family",
        "site", "site level", "factor", "block",
    }
    return header_norm in safe_headers or header_norm.endswith(" level")


def _apply_repeated_cell_merges(
    tbl,
    rows: list[list[TableCell]],
    header_rows: int,
    occupied: set[tuple[int, int]],
):
    """Merge consecutive identical non-empty grouping cells in safe columns."""
    if not rows:
        return
    nrows = len(rows)
    ncols = max(_row_width(r) for r in rows)
    headers = []
    for j in range(ncols):
        if header_rows <= 0:
            headers.append("")
        else:
            header_parts = [
                _cell_text(cell)
                for i in range(min(header_rows, nrows))
                if (cell := _cell_at_col(rows[i], j)) is not None
            ]
            headers.append(" ".join(part for part in header_parts if part).strip())

    for j, header in enumerate(headers):
        if not _is_auto_merge_column(header):
            continue
        i = header_rows
        while i < nrows:
            cell = _cell_at_col(rows[i], j)
            if (i, j) in occupied or cell is None or cell.colspan > 1:
                i += 1
                continue
            value = _cell_text(cell).strip()
            if not value:
                i += 1
                continue
            end = i
            while (
                end + 1 < nrows
                and (end + 1, j) not in occupied
                and (next_cell := _cell_at_col(rows[end + 1], j)) is not None
                and next_cell.colspan == 1
                and _cell_text(next_cell).strip() == value
            ):
                end += 1
            if end > i:
                _merge_vertical_cells(tbl, i, end, j)
                for r in range(i, end + 1):
                    occupied.add((r, j))
            i = end + 1


# ---------------------------------------------------------------------------
#  Column-width estimation
# ---------------------------------------------------------------------------

def _set_table_column_widths(tbl, ncols: int, rows: list[list[TableCell]]):
    """Assign proportional column widths based on header text length.

    Uses Word percentage-width (pct) units to avoid DXA/EMU conversion
    and to keep columns within page bounds regardless of page size.
    """
    if not rows or ncols <= 0:
        return
    # Set table preferred width to 100 % (5000 pct).
    tblPr = tbl._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
        tblPr.insert(0, tblW)
    else:
        tblW.set(qn("w:w"), "5000")
        tblW.set(qn("w:type"), "pct")

    # Estimate column weights from header text lengths.
    header_rows = min(2, len(rows))
    weights = [1.0] * ncols
    for i in range(header_rows):
        for j in range(min(ncols, len(rows[i]))):
            text = _cell_text(rows[i][j])
            char_len = len(re.sub(r"\s+", "", text)) + 1  # +1 avoids zero
            weights[j] = max(weights[j], char_len)

    total = sum(weights)
    if total <= 0:
        return

    # Convert weights to pct values (5000 = 100 %), clamping extremes.
    pct_min = 400   # ~8 %
    pct_max = 2500  # 50 %
    pct_vals: list[int] = []
    for w in weights:
        pct = int(round(w / total * 5000))
        pct = max(pct_min, min(pct_max, pct))
        pct_vals.append(pct)

    # Redistribute so they sum to exactly 5000.
    surplus = 5000 - sum(pct_vals)
    # Distribute surplus proportionally from largest to smallest.
    idxs = sorted(range(ncols), key=lambda k: pct_vals[k], reverse=True)
    for k in idxs:
        if surplus <= 0:
            break
        add = min(surplus, pct_max - pct_vals[k])
        pct_vals[k] += add
        surplus -= add

    for j in range(ncols):
        tcW_val = str(pct_vals[j])
        for row in tbl.rows:
            tc = row.cells[j]._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = parse_xml(
                    f'<w:tcW {nsdecls("w")} w:w="{tcW_val}" w:type="pct"/>'
                )
                tcPr.append(tcW)
            else:
                tcW.set(qn("w:w"), tcW_val)
                tcW.set(qn("w:type"), "pct")


def add_table(doc: Document, rows: list[list[TableCell]], caption_tex: str | None,
              resolver: CitationResolver, extended: bool = False,
              float_prefix: str = ""):
    context=_context()
    if extended:
        context.ed_tbl_counter+=1
        label=f"Extended Data Table {context.ed_tbl_counter}. "
    else:
        context.tbl_counter+=1
        label=f"Table {float_prefix}{context.tbl_counter}. "
    cp = doc.add_paragraph()
    _set_justify(cp)
    cp.paragraph_format.first_line_indent = Cm(0)
    cp.paragraph_format.space_before = PT_HALF_LINE
    r = cp.add_run(label)
    r.bold = True
    r.font.size = PT_CAPTION
    r.font.name = FONT_BODY
    if caption_tex:
        caption_tex = re.sub(r"\n\s*", " ", caption_tex)
        caption_tex = re.sub(r"  +", " ", caption_tex).strip()
        add_rich_text(cp, caption_tex, resolver, font_size=PT_CAPTION)

    if not rows:
        return
    rows = _combine_multilevel_header(rows)
    ncols = max(_row_width(r) for r in rows)
    header_rows = 1
    if len(rows) > 1 and (
        not _cell_text(rows[1][0]).strip() or len(rows[0]) < ncols
    ):
        header_rows = 2

    tbl = doc.add_table(rows=len(rows), cols=ncols)
    # Use "Table Normal" (borderless) instead of "Table Grid" to avoid
    # style-level border conflicts that force manual DOCX cleanup.
    # Borders are set explicitly below via _clear_all_borders / _set_cell_border.
    try:
        tbl.style = doc.styles["Table Normal"]
    except KeyError:
        pass
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Distribute available page width proportionally across columns.
    _set_table_column_widths(tbl, ncols, rows)

    for i, row_data in enumerate(rows):
        for j, cell_tex in _iter_logical_cells(row_data):
            if j >= ncols:
                break
            cell = tbl.cell(i, j)
            if cell_tex.colspan > 1:
                end_col = min(ncols - 1, j + cell_tex.colspan - 1)
                for merge_col in range(j + 1, end_col + 1):
                    _clear_cell_content(tbl.cell(i, merge_col))
                cell = cell.merge(tbl.cell(i, end_col))
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.text = ""
            p = cell.paragraphs[0]
            cell_text = _cell_text(cell_tex)
            is_indented_label = cell_text.lstrip().startswith(r"\quad")
            if is_indented_label:
                cell_text = re.sub(r"^\s*(?:\\quad\s*)+", "", cell_text)
                p.paragraph_format.left_indent = Cm(0.35)
            p.alignment = (
                WD_ALIGN_PARAGRAPH.LEFT
                if cell_tex.colspan > 1 or is_indented_label
                else WD_ALIGN_PARAGRAPH.CENTER
            )
            p.paragraph_format.first_line_indent = Cm(0)
            add_rich_text(p, cell_text, resolver, font_size=Pt(8), citations_superscript=False)
            if i < header_rows:
                for run in p.runs:
                    run.bold = True

    occupied = _apply_explicit_rowspans(tbl, rows, header_rows)
    _apply_repeated_cell_merges(tbl, rows, header_rows, occupied)

    _clear_all_borders(tbl)
    nrows = len(rows)
    SZ_THICK = 8   # 1.0 pt = 8 eighths
    SZ_THIN = 4    # 0.5 pt = 4 eighths
    for j in range(ncols):
        _set_cell_border(tbl.cell(0, j), top=(SZ_THICK, "single"))
        _set_cell_border(tbl.cell(header_rows - 1, j), bottom=(SZ_THIN, "single"))
        _set_cell_border(tbl.cell(nrows - 1, j), bottom=(SZ_THICK, "single"))


# ═══════════════════════════════════════════════════════════════════════════
#  Brace-balanced argument extractor
# ═══════════════════════════════════════════════════════════════════════════

def _extract_braced_arg(text: str, command: str) -> str:
    pat = re.escape(command) + r"\{"
    m = re.search(pat, text)
    if not m:
        return ""
    start = m.end()
    depth = 1
    pos = start
    while pos < len(text) and depth > 0:
        if text[pos] == "{":
            depth += 1
        elif text[pos] == "}":
            depth -= 1
        pos += 1
    return text[start:pos - 1]


# ═══════════════════════════════════════════════════════════════════════════
#  Preamble extraction
# ═══════════════════════════════════════════════════════════════════════════

def _extract_preamble(full_text: str) -> dict:
    data: dict = {"title": "", "authors": [], "affils": [], "abstract": ""}
    m = re.search(r"\\title\{", full_text)
    if m:
        data["title"] = _extract_braced_arg(full_text, r"\title")

    for am in re.finditer(r"\\author(?:\[([^\]]*)\])?\{([^}]+)\}", full_text):
        annotation = am.group(1) or ""
        name = am.group(2)
        is_corr = "*" in annotation
        data["authors"].append((name, is_corr))

    for af in re.finditer(r"\\affil(?:\[[^\]]*\])?\{", full_text):
        start = af.end()
        depth = 1
        pos = start
        while pos < len(full_text) and depth > 0:
            if full_text[pos] == "{":
                depth += 1
            elif full_text[pos] == "}":
                depth -= 1
            pos += 1
        content = full_text[start:pos - 1].strip()
        if content:
            data["affils"].append(content)

    abm = re.search(r"\\begin\{abstract\}(.*?)\\end\{abstract\}", full_text, re.DOTALL)
    if abm:
        data["abstract"] = abm.group(1).strip()
    return data


# ═══════════════════════════════════════════════════════════════════════════
#  Main converter (line-oriented state machine)
# ═══════════════════════════════════════════════════════════════════════════

def _expand_inputs(text: str, base_dir: Path, seen: set[Path] | None = None) -> str:
    """Expand local \\input/\\include fragments used by a TeX source."""
    if seen is None:
        seen = set()

    def repl(m: re.Match) -> str:
        name = m.group(2).strip()
        if not name.endswith(".tex"):
            name += ".tex"
        path = (base_dir / name).resolve()
        if path in seen or not path.exists():
            return m.group(0)
        seen.add(path)
        nested = path.read_text(encoding="utf-8")
        return _expand_inputs(nested, path.parent, seen)

    return re.sub(r"\\(input|include)\{([^}]+)\}", repl, text)


def _body_from_text(text: str) -> str:
    m0 = re.search(r"\\begin\{document\}", text)
    m1 = re.search(r"\\end\{document\}", text)
    if m0 and m1:
        text = text[m0.end():m1.start()]

    # Remove \title{...}, \author{...}, \date{...} with brace-balanced extraction
    for cmd in (r"\title", r"\author", r"\date"):
        while cmd + "{" in text:
            m = re.search(re.escape(cmd) + r"\{", text)
            if not m:
                break
            start = m.start()
            pos = m.end()
            depth = 1
            while pos < len(text) and depth > 0:
                if text[pos] == "{":
                    depth += 1
                elif text[pos] == "}":
                    depth -= 1
                pos += 1
            text = text[:start] + text[pos:]

    # Remove other document structure commands
    for pat in (r"\\tableofcontents", r"\\newpage", r"\\maketitle",
                r"\\clearpage", r"\\thispagestyle\{[^}]*\}", r"\\flushbottom"):
        text = re.sub(pat, "", text)
    return text


def _body_of(path: Path) -> str:
    text = _expand_inputs(path.read_text(encoding="utf-8"), path.parent)
    return _body_from_text(text)


def _convert_tex(tex_path: Path, out_path: Path, bib: dict[str, dict],
                aux_path: Path | None = None,
                inherited_citations: dict[str, int] | None = None,
                citation_prefix: str = "",
                float_prefix: str = "",
                main_authors=None,
                main_affils=None):
    full_text = _expand_inputs(tex_path.read_text(encoding="utf-8"), tex_path.parent)
    aux_text = (
        _expand_inputs(aux_path.read_text(encoding="utf-8"), aux_path.parent)
        if aux_path and aux_path.exists() else None
    )
    _context().label_map=_build_label_map(
        full_text,aux_text,current_prefix=float_prefix,
        aux_prefix="" if float_prefix else "S",
    )
    preamble = _extract_preamble(full_text)
    body = _body_from_text(full_text)

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    setup_styles(doc)
    resolver = CitationResolver(
        bib,
        inherited_numbers=inherited_citations,
        local_prefix=citation_prefix,
    )

    is_si_doc = float_prefix == "S"

    if preamble["title"]:
        add_title(doc, preamble["title"], si_mode=is_si_doc)
    if preamble["authors"]:
        add_authors(doc, preamble["authors"], preamble["affils"])
    elif is_si_doc and main_authors:
        add_authors(doc, main_authors, main_affils or [])
    if preamble["abstract"]:
        add_abstract(doc, preamble["abstract"], resolver)

    lines = body.split("\n")
    i = 0
    in_extended_data = False
    si_section_count = 0
    si_subsection_count = 0
    while i < len(lines):
        line = lines[i].strip()
        # Remove inline comments (% and everything after it, unless % is escaped as \%)
        # First protect escaped \% temporarily
        line = line.replace("\\%", "\x00PERCENT\x00")
        if "%" in line:
            line = line[:line.index("%")].strip()
        # Restore escaped %
        line = line.replace("\x00PERCENT\x00", "\\%")

        if not line or line.startswith("%") or _is_invisible_tex_line(line):
            i += 1
            continue

        # --- display math: \[ ... \] or \begin{equation} ... \end{equation} ---
        if line.startswith(r"\[") or re.match(r"\\begin\{equation\*?\}", line):
            math_block = line
            j = i + 1
            end_marker = r"\]" if line.startswith(r"\[") else r"\end{equation"
            while end_marker not in math_block and j < len(lines):
                math_block += "\n" + lines[j].strip()
                j += 1
            add_display_math(doc, math_block, resolver, equation_prefix=float_prefix)
            i = j
            continue

        # --- headings (brace-balanced to handle nested {2+} etc.) ---
        sec_match = _match_cmd_with_braced_arg(line, r"\\section\*?\{")
        if sec_match:
            sec_title, _ = sec_match
            if is_si_doc:
                if si_section_count > 0:
                    doc.add_page_break()
                si_section_count += 1
            add_heading(doc, sec_title, 1)
            if re.sub(r"\\[a-zA-Z]+\*?", "", sec_title).strip() == "Extended Data":
                in_extended_data = True
            i += 1
            continue
        subsec_match = _match_cmd_with_braced_arg(line, r"\\subsection\*?\{")
        if subsec_match:
            subsec_title, _ = subsec_match
            if is_si_doc:
                if si_subsection_count > 0:
                    spacer = doc.add_paragraph()
                    spacer.paragraph_format.first_line_indent = Cm(0)
                si_subsection_count += 1
            add_heading(doc, subsec_title, 2)
            i += 1
            continue
        subsubsec_match = _match_cmd_with_braced_arg(line, r"\\subsubsection\*?\{")
        if subsubsec_match:
            subsubsec_title, _ = subsubsec_match
            add_heading(doc, subsubsec_title, 3)
            i += 1
            continue

        # --- \paragraph*{...} (brace-balanced) ---
        para_match = _match_cmd_with_braced_arg(line, r"\\paragraph\*?\{")
        if para_match:
            para_title, para_rest = para_match
            para_content = para_rest.strip()
            j = i + 1
            while j < len(lines):
                nl = lines[j].strip()
                # Remove inline comments
                nl = nl.replace("\\%", "\x00PERCENT\x00")
                if "%" in nl:
                    nl = nl[:nl.index("%")].strip()
                nl = nl.replace("\x00PERCENT\x00", "\\%")

                if not nl:
                    break
                if nl.startswith("%") or _is_invisible_tex_line(nl):
                    j += 1
                    continue
                if nl.startswith(r"\[") or re.match(r"\\(section|subsection|paragraph|begin\{)", nl):
                    break
                para_content += " " + nl
                j += 1
            _context().tracker.heading_seen()
            add_body_paragraph(doc, para_content, resolver, bold_prefix_tex=para_title)
            i = j
            continue

        # --- \begin{itemize/enumerate} ... \end{itemize/enumerate} ---
        m_list = re.match(r"\\begin\{(itemize|enumerate)\}", line)
        if m_list:
            list_env = m_list.group(1)
            j = i + 1
            items = []
            while j < len(lines):
                nl = lines[j].strip()
                # Remove inline comments
                nl = nl.replace("\\%", "\x00PERCENT\x00")
                if "%" in nl:
                    nl = nl[:nl.index("%")].strip()
                nl = nl.replace("\x00PERCENT\x00", "\\%")
                if re.match(rf"\\end\{{{list_env}\}}", nl):
                    j += 1
                    break
                if re.match(r"\\setlength\\itemsep\{[^}]*\}", nl):
                    j += 1
                    continue
                if nl.startswith(r"\item"):
                    item_text = nl[5:].strip()  # Remove \item
                    # Collect continuation lines
                    k = j + 1
                    while k < len(lines):
                        cont = lines[k].strip()
                        cont = cont.replace("\\%", "\x00PERCENT\x00")
                        if "%" in cont:
                            cont = cont[:cont.index("%")].strip()
                        cont = cont.replace("\x00PERCENT\x00", "\\%")
                        if not cont or cont.startswith("%") or _is_invisible_tex_line(cont):
                            k += 1
                            continue
                        if re.match(r"\\setlength\\itemsep\{[^}]*\}", cont):
                            k += 1
                            continue
                        if cont.startswith(r"\item") or re.match(rf"\\end\{{{list_env}\}}", cont):
                            break
                        item_text += " " + cont
                        k += 1
                    items.append(item_text)
                    j = k
                else:
                    j += 1

            # Add items as bullets or numbered list entries.
            list_style = "List Bullet" if list_env == "itemize" else "List Number"
            for item in items:
                p = doc.add_paragraph(style=list_style)
                _set_justify(p)
                p.paragraph_format.first_line_indent = Cm(0)
                add_rich_text(p, item, resolver)

            i = j
            continue

        # --- algorithm / algorithmic ---
        if r"\begin{algorithm" in line:
            alg_block = line
            j = i + 1
            while j < len(lines) and r"\end{algorithm}" not in alg_block:
                alg_block += "\n" + lines[j]
                j += 1
            add_algorithm(doc, alg_block, resolver, float_prefix=float_prefix)
            i = j
            continue

        # --- figure ---
        if r"\begin{figure" in line:
            fig_block = line
            j = i + 1
            while j < len(lines) and r"\end{figure" not in fig_block:
                fig_block += "\n" + lines[j]
                j += 1
            fig_names = re.findall(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}", fig_block)
            fig_name = fig_names if len(fig_names) > 1 else (fig_names[0] if fig_names else "")
            caption_raw = _extract_braced_arg(fig_block, r"\caption")
            add_figure(
                doc,
                fig_name,
                caption_raw,
                resolver,
                extended=in_extended_data,
                float_prefix=float_prefix,
            )
            i = j
            continue

        # --- table / longtable ---
        if r"\begin{table" in line or r"\begin{longtable" in line:
            is_lt = "longtable" in line
            blk = line
            j = i + 1
            end = r"\end{longtable}" if is_lt else r"\end{table"
            while j < len(lines) and end not in blk:
                blk += "\n" + lines[j]
                j += 1
            caption_raw = _extract_braced_arg(blk, r"\caption")
            rows = _parse_longtable_rows(blk) if is_lt else _parse_tabular_rows(blk)
            add_table(
                doc,
                rows,
                caption_raw,
                resolver,
                extended=in_extended_data,
                float_prefix=float_prefix,
            )
            i = j
            continue

        # --- bibliography ---
        if r"\bibliography{" in line:
            refs = resolver.reference_items()
            if refs:
                if is_si_doc:
                    if si_section_count > 0:
                        doc.add_page_break()
                    si_section_count += 1
                add_heading(doc, "References", 1)
                for label, ref_text in refs:
                    p = doc.add_paragraph()
                    _set_justify(p)
                    p.paragraph_format.first_line_indent = Cm(0)
                    r = p.add_run(f"[{label}] {ref_text}")
                    r.font.size = PT_REF
                    r.font.name = FONT_BODY
                    add_bookmark(p, _reference_bookmark_name(label))
            i += 1
            continue

        # --- skip no-ops ---
        if re.match(r"\\(begin|end)\{(document|abstract)\}", line):
            i += 1
            continue
        if re.match(r"\\(bibliographystyle|usepackage|renewcommand|newcommand|setcounter)", line):
            i += 1
            continue

        # --- regular paragraph ---
        para_lines = [line]
        j = i + 1
        while j < len(lines):
            nl = lines[j].strip()
            # Remove inline comments from this line too
            nl = nl.replace("\\%", "\x00PERCENT\x00")
            if "%" in nl:
                nl = nl[:nl.index("%")].strip()
            nl = nl.replace("\x00PERCENT\x00", "\\%")

            if not nl:
                break
            if nl.startswith("%") or _is_invisible_tex_line(nl):
                j += 1
                continue
            if nl.startswith(r"\[") or re.match(r"\\(section|subsection|subsubsection|paragraph|begin\{figure|begin\{table|begin\{longtable|begin\{algorithm|begin\{itemize\}|begin\{enumerate\}|begin\{equation\*?\}|bibliography\{|end\{)", nl):
                break
            para_lines.append(nl)
            j += 1
        add_body_paragraph(doc, " ".join(para_lines), resolver)
        i = max(j, i + 1)

    os.makedirs(out_path.parent, exist_ok=True)
    doc.save(str(out_path))
    print(f"  -> {out_path}  ({out_path.stat().st_size / 1024:.0f} KB)")



_LATEX_LEAK_RE = re.compile(
    r"\\(?:begin|end|cite\w*|includegraphics|section|subsection|label|ref)\b"
)


def validate_docx_package(path: Path, tex_path: Path, auxiliary_tex_path: Path | None = None) -> dict[str, int]:
    """Validate a generated DOCX against the TeX source it represents."""
    with zipfile.ZipFile(path) as package:
        corrupt = package.testzip()
        if corrupt is not None:
            raise ValueError(f"corrupt DOCX member: {corrupt}")
        required = {"[Content_Types].xml", "word/document.xml"}
        missing = required.difference(package.namelist())
        if missing:
            raise ValueError(f"DOCX is missing required parts: {sorted(missing)}")

    document = Document(str(path))
    root = Package.load(path).xml("word/document.xml")
    visible = "\n".join(
        visible_text(block.element, "final") for block in _blocks(root)[1]
    )
    normalized_visible = re.sub(r"\s+", " ", visible).strip()

    def hidden_in_final(element) -> bool:
        parent = element
        while parent is not None:
            if parent.tag in {
                f"{{{DIFF_W}}}del",
                f"{{{DIFF_W}}}moveFrom",
            }:
                return True
            if parent.tag == f"{{{DIFF_W}}}tr":
                properties = parent.find("./w:trPr", DIFF_NS)
                if (
                    properties is not None
                    and properties.find("./w:del", DIFF_NS) is not None
                ):
                    return True
            parent = parent.getparent()
        return False
    primary_source = _strip_latex_comments(_expand_inputs(tex_path.read_text(encoding="utf-8"),tex_path.parent))
    source = primary_source
    if auxiliary_tex_path is not None and Path(auxiliary_tex_path).exists():
        aux=Path(auxiliary_tex_path)
        source += "\n"+_strip_latex_comments(_expand_inputs(aux.read_text(encoding="utf-8"),aux.parent))
    preamble = _extract_preamble(primary_source)
    expected_title = re.sub(r"\s+", " ", plain_tex(preamble["title"])).strip()
    expected_title = expected_title.replace("\\\\", " ").replace("\\", " ")
    expected_title = re.sub(r"\s+", " ", expected_title).strip()
    if expected_title and expected_title not in normalized_visible:
        raise ValueError(f"DOCX title does not match {tex_path.name}")
    leak = _LATEX_LEAK_RE.search(visible)
    if leak:
        raise ValueError(f"LaTeX command leaked into DOCX: {leak.group(0)}")
    if r"\bibliography{" in primary_source and "References" not in visible:
        raise ValueError("DOCX is missing the References heading")

    if "[Figure:" in visible:
        raise ValueError("DOCX contains a missing-figure placeholder")
    expected_figures = len(re.findall(
        r"\\includegraphics(?:\[[^\]]*\])?\{[^}]+\}", primary_source
    ))
    actual_figures = sum(
        not hidden_in_final(drawing)
        for drawing in root.findall(".//w:drawing", DIFF_NS)
    )
    if actual_figures != expected_figures:
        raise ValueError(
            f"figure count mismatch: expected {expected_figures}, found {actual_figures}"
        )
    expected_tables = len(re.findall(
        r"\\begin\{(?:table\*?|longtable)\}", primary_source
    ))
    actual_tables = 0
    for table in root.findall("./w:body/w:tbl", DIFF_NS):
        rows = table.findall("./w:tr", DIFF_NS)
        if not rows or any(not hidden_in_final(row) for row in rows):
            actual_tables += 1
    if actual_tables != expected_tables:
        raise ValueError(
            f"table count mismatch: expected {expected_tables}, found {actual_tables}"
        )
    return {
        "figures": actual_figures,
        "tables": actual_tables,
        "paragraphs": len(document.paragraphs),
    }


def build_validated_docx(
    output_path: Path,
    tex_path: Path,
    builder: Callable[[Path], None],
    *,
    atomic: bool = True,
    auxiliary_tex_path: Path | None = None,
) -> dict[str, int]:
    """Build, validate and optionally atomically publish a DOCX."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not atomic:
        builder(output_path)
        return validate_docx_package(output_path, tex_path, auxiliary_tex_path)

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}-", suffix=".docx", dir=output_path.parent
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        builder(temporary_path)
        summary = validate_docx_package(temporary_path, tex_path, auxiliary_tex_path)
        os.replace(temporary_path, output_path)
        return summary
    finally:
        temporary_path.unlink(missing_ok=True)




def convert_tex(tex_path: Path,out_path: Path,bib: dict[str,dict],aux_path: Path|None=None,
                inherited_citations: dict[str,int]|None=None,citation_prefix: str="",
                float_prefix: str="",main_authors=None,main_affils=None,*,figure_dir: Path|None=None)->Document:
    tex_path=Path(tex_path).resolve(); out_path=Path(out_path)
    context=ConversionContext(Path(figure_dir).resolve() if figure_dir else tex_path.parent/"figures",tex_path.parent)
    token=_ACTIVE_CONTEXT.set(context)
    try:
        _convert_tex(tex_path,out_path,bib,aux_path,inherited_citations,citation_prefix,float_prefix,main_authors,main_affils)
        return Document(str(out_path))
    finally:
        _ACTIVE_CONTEXT.reset(token)


def _citation_numbers(full_text: str,bib: dict[str,dict])->dict[str,int]:
    resolver=CitationResolver(bib)
    preamble=_extract_preamble(full_text)
    for match in re.finditer(r"\\(?:cite|mainref)\{([^}]+)\}",preamble["abstract"]): resolver.resolve(match.group(1))
    body=_body_from_text(full_text)
    body=_strip_latex_comments(body)
    for match in re.finditer(r"\\(?:cite|mainref)\{([^}]+)\}",body): resolver.resolve(match.group(1))
    return resolver.citation_numbers()


def latex_to_docx(main_text: str,si_text: str|None=None,bib: dict[str,dict]|None=None,*,
                  fig_dir: Path|None=None,output: Path|None=None,target: str="main",
                  inherited_citations: dict[str,int]|None=None,citation_prefix: str="",float_prefix: str="")->Document:
    if target not in {"main","si"}: raise ValueError("target must be 'main' or 'si'")
    if target=="si" and si_text is None: raise ValueError("target='si' requires si_text")
    with tempfile.TemporaryDirectory(prefix="docforge-tex-") as temp:
        root=Path(temp); main_path=root/"main.tex"; si_path=root/"SI.tex"
        main_path.write_text(main_text,encoding="utf-8")
        if si_text is not None: si_path.write_text(si_text,encoding="utf-8")
        source=main_path if target=="main" else si_path
        out=Path(output) if output else root/f"{target}.docx"
        return convert_tex(source,out,bib or {},aux_path=si_path if target=="main" and si_text else main_path if target=="si" else None,
            inherited_citations=inherited_citations,citation_prefix=citation_prefix,
            float_prefix=float_prefix or ("S" if target=="si" else ""),figure_dir=fig_dir)


def convert_files(main_path: Path,si_path: Path|None=None,bib_path: Path|None=None,*,output: Path|None=None,
                  target: str="main",inherited_citations: dict[str,int]|None=None,
                  citation_prefix: str="",float_prefix: str="")->Document:
    main_path=Path(main_path).resolve(); si_path=Path(si_path).resolve() if si_path else None
    if target not in {"main","si"}: raise ValueError("target must be 'main' or 'si'")
    source=main_path if target=="main" else si_path
    if source is None: raise ValueError("target='si' requires si_path")
    bibliography=parse_bib(Path(bib_path)) if bib_path else {}
    preamble=_extract_preamble(_expand_inputs(main_path.read_text(encoding="utf-8"),main_path.parent))
    with tempfile.TemporaryDirectory(prefix="docforge-tex-output-") as temporary:
        out=Path(output) if output is not None else Path(temporary)/f"{target}.docx"
        return convert_tex(source,out,bibliography,aux_path=si_path if target=="main" else main_path,
            inherited_citations=inherited_citations,citation_prefix=citation_prefix,
            float_prefix=float_prefix or ("S" if target=="si" else ""),figure_dir=main_path.parent/"figures",
            main_authors=preamble["authors"],main_affils=preamble["affils"])



def build_label_map(full_text: str, auxiliary_text: str | None = None, *, current_prefix: str = "", auxiliary_prefix: str = "S") -> dict[str,str]:
    return _build_label_map(full_text, auxiliary_text, current_prefix, auxiliary_prefix)


def collect_citation_numbers(full_text: str, bib: dict[str,dict]) -> dict[str,int]:
    return _citation_numbers(full_text, bib)


__all__=["ConversionContext","convert_tex","convert_files","latex_to_docx","validate_docx_package","build_validated_docx","build_label_map","collect_citation_numbers","parse_bib","CitationResolver"]
