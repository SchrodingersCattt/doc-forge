"""LaTeX -> DOCX converter.

``latex_to_docx`` converts a manuscript written in the documented LaTeX/TeX
subset into a styled Word document: sections/subsections, figures, tables,
longtables, algorithms, inline and display math, and a resolved bibliography.
Cross-references between ``main.tex`` and ``si.tex`` follow ``\\externaldocument``
prefixed-label conventions (e.g. ``[S-]``).
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from .bib import CitationResolver, parse_bib
from ..output import validate_output_path
from .tokenize import TableCell, spans_to_plain, tokenize_tex
from ..markdown.omml import normalize_math_source, parse_math_omml

FONT_HEADING = "Arial"
FONT_BODY = "Times New Roman"


def _scan_labels(text: str, fig_offset: int = 0, tbl_offset: int = 0, prefix: str = "") -> dict[str, str]:
    """Scan a .tex file for labels and return a key->display map."""
    body = text
    m = re.search(r"\\begin\{document\}", text)
    if m:
        body = text[m.end() :]
    result: dict[str, str] = {}
    fig_n = fig_offset
    tbl_n = tbl_offset
    alg_n = 0
    ed_fig_n = 0
    ed_tbl_n = 0
    sec_n = 0
    subsec_n = 0
    suppnote_n = 0
    supprecord_n = 0
    in_extended_data = False
    for m in re.finditer(
        r"\\section\*?\{([^}]*)\}|\\subsection\*?\{([^}]*)\}|"
        r"\\begin\{(figure|figure\*|table|table\*|longtable|algorithm)\b|\\label\{([^}]*)\}",
        body,
    ):
        if m.group(1):
            heading = re.sub(r"\\[a-zA-Z]+\*?", "", m.group(1)).strip()
            note_match = re.match(r"Supplementary Note\s+(\d+)", heading)
            record_match = re.match(r"Supplementary Record\s+(\d+)", heading)
            if note_match:
                suppnote_n = int(note_match.group(1))
            if record_match:
                supprecord_n = int(record_match.group(1))
            if heading == "Extended Data":
                in_extended_data = True
            else:
                sec_n += 1
                subsec_n = 0
            continue
        if m.group(2):
            subsec_n += 1
            continue
        env = m.group(3)
        if env in ("figure", "figure*"):
            if in_extended_data:
                ed_fig_n += 1
            else:
                fig_n += 1
        elif env in ("table", "table*", "longtable"):
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
                result[key] = f"{prefix}{sec_n}.{subsec_n}" if subsec_n > 0 else f"{prefix}{sec_n}"
            elif key.startswith("sn:"):
                if supprecord_n:
                    result[key] = f"{supprecord_n}"
                elif suppnote_n:
                    result[key] = f"{suppnote_n}"
                else:
                    result[key] = f"{sec_n}"
            elif key.startswith("sec:"):
                result[key] = f"{prefix}{sec_n}"
            else:
                result[key] = key
    return result


def build_label_map(full_text: str, aux_text: str | None = None) -> dict[str, str]:
    """Combine label maps for one .tex file and an optional cross-referenced SI."""
    result = _scan_labels(full_text)
    if aux_text:
        aux_labels = _scan_labels(aux_text, prefix="S-")
        for k, v in aux_labels.items():
            if k not in result:
                result[k] = v
            prefixed_key = f"S-{k}"
            if prefixed_key not in result:
                result[prefixed_key] = v
    return result


def _reference_prefix(text: str) -> str:
    m = re.search(r"\\externaldocument\[([^\]]*)\]\{([^}]*)\}", text)
    return m.group(1) if m else ""


def latex_to_docx(
    main_text: str,
    si_text: str | None = None,
    bib: dict[str, dict] | None = None,
    *,
    fig_dir: Path | None = None,
    output: Path | None = None,
) -> Document:
    """Convert main (and optional SI) LaTeX text into a styled DOCX.

    *fig_dir* resolves ``\\includegraphics`` paths; *output* saves the document
    if given. Returns the opened Document.
    """
    validate_output_path(output)
    resolver = CitationResolver(bib or {})

    def resolve_ref(key: str) -> str:
        label_map = build_label_map(main_text, si_text)
        return label_map.get(key, "??")

    doc = Document()
    _setup_doc_styles(doc)
    _convert_text(doc, main_text, fig_dir=fig_dir, resolve_ref=resolve_ref)

    if si_text:
        doc.add_page_break()
        _convert_text(doc, si_text, fig_dir=fig_dir, resolve_ref=resolve_ref, si=True)

    if bib:
        _append_bibliography(doc, resolver)

    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output))
    return doc


def _setup_doc_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONT_BODY
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(4)
    for level in range(1, 6):
        name = f"Heading {level}"
        try:
            style = doc.styles[name]
        except KeyError:
            from docx.enum.style import WD_STYLE_TYPE

            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.base_style = normal
        style.font.name = FONT_HEADING
        style.font.color.rgb = None
        style.quick_style = True


def _convert_text(
    doc: Document,
    text: str,
    *,
    fig_dir: Path | None,
    resolve_ref,
    si: bool = False,
) -> None:
    pending_refs: list[str] = []
    environment_stack: list[str] = []
    in_math_display = False
    math_buffer: list[str] = []

    def flush_math() -> None:
        nonlocal math_buffer, in_math_display
        if math_buffer:
            _add_equation(doc, " ".join(math_buffer))
            math_buffer = []
            in_math_display = False

    for line in text.splitlines():
        stripped = line.strip()
        # figure/tabular/algorithm float boundaries
        if re.match(r"\\begin\{(figure|figure\*|table|table\*|longtable|algorithm|tabular)\}", stripped):
            if re.match(r"\\begin\{(figure|figure\*|table|table\*|longtable)\}", stripped):
                environment_stack.append("float")
            elif re.match(r"\\begin\{(tabular)\}", stripped):
                environment_stack.append("tabular")
            else:
                environment_stack.append("algorithm")
            continue
        if re.match(r"\\end\{(figure|figure\*|table|table\*|longtable|algorithm|tabular)\}", stripped):
            if environment_stack:
                environment_stack.pop()
            continue
        if stripped.startswith("\\begin{document}") or stripped.startswith("\\end{document}"):
            continue
        if stripped.startswith("%"):
            continue
        if environment_stack and environment_stack[-1] == "float":
            # caption / label / includegraphics inside a float
            m_cap = re.match(r"\\caption\{(\s*.*?)\s*\}", stripped)
            if m_cap:
                _add_paragraph(doc, m_cap.group(1).strip(), bold=True, italic=False)
                continue
            m_img = re.match(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]*)\}", stripped)
            if m_img:
                _add_image(doc, m_img.group(1), fig_dir)
                continue
            continue
        if stripped == r"\begin{equation*}" or stripped == r"\begin{equation}":
            in_math_display = True
            math_buffer = []
            continue
        if stripped == r"\end{equation*}" or stripped == r"\end{equation}":
            flush_math()
            continue
        if stripped == r"\begin{center}":
            continue
        if stripped == r"\end{center}":
            continue
        m_section = re.match(r"\\section\*?\{([^}]*)\}", stripped)
        if m_section:
            flush_math()
            _add_heading(doc, m_section.group(1), 1)
            continue
        m_sub = re.match(r"\\subsection\*?\{([^}]*)\}", stripped)
        if m_sub:
            flush_math()
            _add_heading(doc, m_sub.group(1), 2)
            continue
        m_subsub = re.match(r"\\subsubsection\*?\{([^}]*)\}", stripped)
        if m_subsub:
            flush_math()
            _add_heading(doc, m_subsub.group(1), 3)
            continue
        m_cite = re.search(r"\\cite\{([^}]*)\}", stripped)
        if m_cite:
            for key in m_cite.group(1).split(","):
                key = key.strip()
                if key and key not in pending_refs:
                    pending_refs.append(key)
            continue
        if stripped == r"\bibliography{ref}":
            continue
        if stripped.startswith(r"\par"):
            continue
        if re.match(r"\\label\{", stripped):
            continue
        if stripped.startswith(r"\externaldocument"):
            continue
        if in_math_display and stripped:
            math_buffer.append(stripped)
            continue
        if stripped:
            _add_paragraph(doc, stripped, resolve_ref=resolve_ref)

    flush_math()
    if pending_refs:
        for key in pending_refs:
            resolver.number_for(key)


def _add_heading(doc: Document, text: str, level: int) -> None:
    paragraph = doc.add_paragraph()
    paragraph.style = doc.styles[f"Heading {min(level, 5)}"]
    paragraph.style.font.name = FONT_HEADING
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = FONT_HEADING


def _add_paragraph(doc: Document, text: str, *, bold: bool = False, italic: bool = False, resolve_ref=None) -> None:
    spans = tokenize_tex(text, resolve_ref=resolve_ref)
    paragraph = doc.add_paragraph()
    for span in spans:
        run = paragraph.add_run(span.text)
        run.bold = bold or span.bold
        run.italic = italic or span.italic
        if span.superscript:
            run.font.superscript = True
        if span.subscript:
            run.font.subscript = True
        if span.highlight:
            run.font.highlight_color = span.highlight
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return paragraph


def _add_equation(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    math_para = OxmlElement("m:oMathPara")
    math = OxmlElement("m:oMath")
    for element in parse_math_omml(normalize_math_source(text)):
        math.append(element)
    math_para.append(math)
    paragraph._p.append(math_para)


def _add_image(doc: Document, rel_path: str, fig_dir: Path | None) -> None:
    path = fig_dir / rel_path if fig_dir else Path(rel_path)
    if not path.exists():
        raise FileNotFoundError(f"Figure file does not exist: {path}")
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Inches(5.8))


def _append_bibliography(doc: Document, resolver: CitationResolver) -> None:
    _add_heading(doc, "References", 1)
    for key, number in sorted(resolver.numbers.items(), key=lambda item: item[1]):
        entry = resolver.bib.get(key)
        if entry is None:
            _add_paragraph(doc, f"{number}. {key} (missing BibTeX entry)")
            continue
        text = spans_to_plain(tokenize_tex(f"{number}. {_format_entry(entry)}"))
        _add_paragraph(doc, text)


def _format_entry(entry: dict) -> str:
    authors = entry.get("author", "unknown")
    authors = authors.replace(" and ", ", ").replace(" and others", " et al.")
    parts = [part for part in (authors, entry.get("year"), entry.get("title")) if part]
    venue = entry.get("journal", entry.get("booktitle", ""))
    if venue:
        parts.append(venue)
    return ". ".join(parts)


def convert_files(
    main_path: Path,
    si_path: Path | None = None,
    bib_path: Path | None = None,
    *,
    output: Path | None = None,
) -> Document:
    """CLI-friendly wrapper: read files, convert, save."""
    main_text = main_path.read_text(encoding="utf-8")
    si_text = si_path.read_text(encoding="utf-8") if si_path else None
    bib = parse_bib(bib_path) if bib_path else None
    fig_dir = main_path.parent
    return latex_to_docx(main_text, si_text, bib, fig_dir=fig_dir, output=output)


__all__ = ["latex_to_docx", "convert_files", "build_label_map", "parse_bib", "CitationResolver"]