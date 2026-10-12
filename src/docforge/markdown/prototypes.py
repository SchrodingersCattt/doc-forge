"""Shared semantic prototype discovery for template and source editing.

The two Markdown writers consume different XML wrappers (``w:p`` elements
versus python-docx ``Paragraph`` objects), but list-heading detection must be
identical.  This module deliberately accepts either representation and keeps
the decision based on explicit numbering and clean bold text rather than
indentation heuristics.
"""

from __future__ import annotations

import copy

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def element_of(value):
    """Return a ``w:p`` element from either a paragraph wrapper or an element."""
    return getattr(value, "_p", value)


def has_clean_direct_bold_run(value) -> bool:
    """Whether a paragraph has a direct bold text run without layout markers."""
    paragraph = element_of(value)
    for run in paragraph.findall(f"{{{W}}}r"):
        properties = run.find(f"{{{W}}}rPr")
        bold = properties.find(f"{{{W}}}b") if properties is not None else None
        if bold is None and properties is not None:
            bold = properties.find(f"{{{W}}}bCs")
        if bold is None or (bold.get(f"{{{W}}}val", "true").lower() in {"0", "false", "off", "none"}):
            continue
        if any(
            node.tag in {
                f"{{{W}}}tab",
                f"{{{W}}}instrText",
                f"{{{W}}}fldChar",
                f"{{{W}}}fldSimple",
                f"{{{W}}}hyperlink",
            }
            for node in run.iter()
        ):
            continue
        if run.find(f".//{{{W}}}highlight") is not None:
            continue
        if any(node.text for node in run.iter(f"{{{W}}}t")):
            return True
    return False


def numbered_bold_prototypes(values, *, copy_nodes: bool = False) -> dict[int, object]:
    """Return the first clean numbered-bold paragraph at each list level."""
    result: dict[int, object] = {}
    for value in values:
        element = element_of(value)
        numbering = element.find(f"{{{W}}}pPr/{{{W}}}numPr")
        if numbering is None or not has_clean_direct_bold_run(element):
            continue
        if any(
            node.tag in {
                f"{{{W}}}tab",
                f"{{{W}}}instrText",
                f"{{{W}}}fldChar",
                f"{{{W}}}fldSimple",
                f"{{{W}}}hyperlink",
            }
            for node in element.iter()
        ):
            continue
        if any(
            token in (node.text or "").upper()
            for node in element.iter(f"{{{W}}}instrText")
            for token in ("PAGEREF", "HYPERLINK", "TOC")
        ):
            continue
        level_node = numbering.find(f"{{{W}}}ilvl")
        try:
            level = int(level_node.get(f"{{{W}}}val", "0")) if level_node is not None else 0
        except (TypeError, ValueError):
            continue
        result.setdefault(level, copy.deepcopy(element) if copy_nodes else element)
    return result


__all__ = ["element_of", "has_clean_direct_bold_run", "numbered_bold_prototypes"]
