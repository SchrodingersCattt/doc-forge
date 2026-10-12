"""Apply a Markdown paragraph delta onto a source DOCX.

Template assembly redraws every paragraph and drops layout that Markdown
cannot carry: title-page styles, index leaders, equations, and tables.
This editor leaves the source package untouched except for paragraphs whose
baseline Markdown text uniquely matches a source paragraph and whose edited
Markdown text changed. New paragraphs are inserted beside that anchor and
copy its paragraph and run properties.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import shutil
import tempfile
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from lxml import etree

from ..docxdiff.package import Package
from ..docxmd.inline import Atom, parse as parse_inline
from ..output import validate_output_path
from .prototypes import numbered_bold_prototypes
from .inline import display, key as inline_key

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
TAG_RE = re.compile(r"</?[^>]+>")
ESCAPE_RE = re.compile(r"\\([\\`*{}\[\]()#+\-.!_>~|])")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
DASHES = str.maketrans(
    {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
        "\u00ad": "",
    }
)


def _local(element: etree._Element) -> str:
    return etree.QName(element).localname


def _opaque_kind(element: etree._Element) -> str | None:
    """Return the stable block kind for a non-editable document node."""
    name = _local(element)
    if name == "tbl":
        return "table"
    if name == "sectPr":
        return "sectPr"
    if name != "p":
        return "opaque"
    if element.find(f".//{{{W}}}drawing") is not None:
        return "drawing"
    if element.find(f".//{{{M}}}oMath") is not None or element.find(f".//{{{M}}}oMathPara") is not None:
        return "equation"
    if element.find(f".//{{{W}}}fldSimple") is not None or element.find(f".//{{{W}}}instrText") is not None:
        return "field"
    if element.find(f".//{{{W}}}fldChar") is not None:
        return "field"
    return None


def _element_signature(element: etree._Element) -> str:
    """Hash an opaque node after removing volatile Word bookkeeping ids."""
    clone = copy.deepcopy(element)
    for node in clone.iter():
        for key in list(node.attrib):
            local = key.rsplit("}", 1)[-1].casefold()
            if local.endswith("id") or local.startswith("rsid") or local in {"id", "paraid", "textid"}:
                del node.attrib[key]
    return hashlib.sha256(etree.tostring(clone, encoding="UTF-8")).hexdigest()


def export_block_map(source: Path, output: Path | None = None) -> dict[str, Any]:
    """Export the source document's ordered editable and opaque block map.

    Text entries carry the body-child index and paragraph index.  Tables,
    drawings, equations, fields, and the terminal section properties receive
    deterministic opaque ids and signatures so a caller can detect an edit
    before writing an output package.
    """
    package = Package.load(source)
    root = package.xml("word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        raise ValueError("DOCX has no document body")
    entries: list[dict[str, Any]] = []
    paragraph_index = 0
    opaque_counts: dict[str, int] = {}
    for source_index, child in enumerate(list(body)):
        kind = _opaque_kind(child)
        if kind is None:
            entry = {
                "kind": "text",
                "source_index": source_index,
                "source_paragraph_index": paragraph_index,
                "text": "".join(node.text or "" for node in child.iter(f"{{{W}}}t")),
            }
            paragraph_index += 1
        else:
            number = opaque_counts.get(kind, 0)
            opaque_counts[kind] = number + 1
            entry = {
                "kind": kind,
                "source_index": source_index,
                "id": f"{kind}-{number}",
                "opaque_id": f"{kind}-{number}",
                "signature": _element_signature(child),
            }
            if _local(child) == "p":
                entry["source_paragraph_index"] = paragraph_index
                paragraph_index += 1
        entries.append(entry)
    result = {
        "version": 1,
        "source_part": "word/document.xml",
        "source_sha256": hashlib.sha256(Path(source).read_bytes()).hexdigest(),
        "blocks": entries,
        # Short aliases make the map convenient to consume without walking
        # the ordered block list while retaining that list as the canonical
        # representation.
        "text": {
            str(entry["source_paragraph_index"]): entry["source_index"]
            for entry in entries
            if entry.get("kind") == "text"
        },
        "opaque": {
            str(entry["id"]): {
                "kind": entry["kind"],
                "source_index": entry["source_index"],
                "signature": entry["signature"],
            }
            for entry in entries
            if entry.get("kind") != "text"
        },
    }
    if output is not None:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return result


# Descriptive alias used by integrations that call this a source map.
export_source_block_map = export_block_map


def _read_block_map(value: Path | dict[str, Any] | list[dict[str, Any]] | None) -> dict[str, Any] | None:
    if value is None:
        return None
    if isinstance(value, Path):
        value = json.loads(value.read_text(encoding="utf-8"))
    if isinstance(value, list):
        value = {"version": 1, "blocks": value}
    if not isinstance(value, dict) or not isinstance(value.get("blocks"), list):
        raise ValueError("block map must contain a blocks list")
    return value


def apply_markdown_delta(
    source: Path,
    baseline_markdown: Path,
    edited_markdown: Path,
    output: Path,
    *,
    overwrite: bool = False,
    allow_delete: bool = False,
    block_map: Path | dict[str, Any] | list[dict[str, Any]] | None = None,
) -> dict[str, int]:
    """Write ``output`` as ``source`` plus the paragraph delta of the Markdown.

    Paragraphs, drawings, tables, and equations that the delta does not name
    stay byte-for-byte in ``word/document.xml`` apart from the inserted
    siblings. A paragraph is editable only when its baseline Markdown block
    matches exactly one source paragraph.
    """

    validate_output_path(output)
    package = Package.load(source)
    source_map = _read_block_map(block_map)
    if source_map is not None and source_map.get("source_sha256"):
        actual_source_hash = hashlib.sha256(Path(source).read_bytes()).hexdigest()
        if source_map["source_sha256"] != actual_source_hash:
            raise ValueError("block map source package is stale")
    root = package.xml("word/document.xml")
    operations = plan_markdown_delta(
        root,
        baseline_markdown.read_text(encoding="utf-8"),
        edited_markdown.read_text(encoding="utf-8"),
        styles_xml=package.parts.get("word/styles.xml"),
        allow_delete=allow_delete,
        block_map=source_map,
    )
    # Reuse the source package byte-for-byte for a no-op.  Re-zipping an
    # untouched package changes ZIP metadata and can invalidate downstream
    # package hashes even though the document parts are equal.
    if not operations:
        if output.exists() and not overwrite:
            raise FileExistsError(output)
        if source.resolve() != output.resolve():
            output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, output)
        return {"replaced": 0, "inserted": 0}
    if operations:
        _apply(root, operations)
        package.set_xml("word/document.xml", root)
    package.write(output, overwrite=overwrite)
    summary = {"replaced": 0, "inserted": 0}
    for operation in operations:
        if operation[0] == "replace":
            summary["replaced"] += 1
        elif operation[0] == "insert":
            summary["inserted"] += 1
        elif operation[0] == "delete":
            summary["deleted"] = summary.get("deleted", 0) + len(operation[1])
    return summary


def plan_markdown_delta(
    root: etree._Element,
    baseline: str,
    edited: str,
    *,
    styles_xml: bytes | None = None,
    allow_delete: bool = False,
    block_map: Path | dict[str, Any] | list[dict[str, Any]] | None = None,
) -> list[tuple]:
    """Return replace and insert operations against ``root``'s body children."""

    body = root.find(f"{{{W}}}body")
    if body is None:
        raise ValueError("DOCX has no document body")
    all_children = list(body)
    children = [child for child in all_children if etree.QName(child).localname != "sectPr"]
    docx_keys = [(index, _paragraph_key(child)) for index, child in enumerate(children)]
    baseline_blocks = _blocks(baseline)
    edited_blocks = _blocks(edited)
    source_map = _read_block_map(block_map)
    mapping, opaque_mapping = _map_blocks_with_map(
        baseline_blocks, docx_keys, children, source_map
    )
    _validate_opaque_map(children, all_children, source_map)
    heading_ids = _heading_style_ids(styles_xml)
    prototypes = _prototypes(children, heading_ids)
    operations: list[tuple] = []
    matcher = SequenceMatcher(
        a=[_opcode_key(block) for block in baseline_blocks],
        b=[_opcode_key(block) for block in edited_blocks],
        autojunk=False,
    )
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "delete":
            if any(index in opaque_mapping for index in range(i1, i2)):
                raise ValueError("markdown delta edits an opaque block")
            if not allow_delete:
                raise ValueError("markdown delta deletes a block; pass allow_delete=True to permit deletion")
            mapped = [mapping.get(index) for index in range(i1, i2)]
            if any(index is None for index in mapped):
                raise ValueError("deleted Markdown block does not match a unique source paragraph")
            operations.append(("delete", mapped))
            continue
        new_blocks = edited_blocks[j1:j2]
        if any(block.kind != "text" for block in new_blocks):
            raise ValueError("markdown delta changes a table or image; those stay in the source DOCX")
        for block in new_blocks:
            _reject_inline_atoms(block.raw)
        if any(index in opaque_mapping for index in range(i1, i2)):
            raise ValueError("markdown delta edits an opaque block")
        if tag == "insert":
            next_anchor = _next_mapped(mapping, i1)
            style_anchor = max(
                (value for key, value in mapping.items() if key < i1),
                default=None,
            )
            operations.append(("insert", next_anchor, new_blocks, prototypes, style_anchor))
            continue
        mapped = [mapping.get(index) for index in range(i1, i2)]
        if not mapped or any(index is None for index in mapped):
            missing = baseline_blocks[i1].raw[:80]
            raise ValueError(f"edited paragraph does not match a unique source paragraph: {missing}")
        next_anchor = _next_mapped(mapping, i2)
        operations.append(("replace", mapped, new_blocks, prototypes, next_anchor))
    return operations


class _Block:
    def __init__(self, kind: str, raw: str) -> None:
        self.kind = kind
        self.raw = raw

    @property
    def key(self) -> str:
        return _key(self.raw)


def _map_blocks_with_map(
    blocks: list[_Block],
    docx_keys: list[tuple[int, str]],
    children: list[etree._Element],
    source_map: dict[str, Any] | None,
) -> tuple[dict[int, int], set[int]]:
    """Map Markdown blocks and mark baseline blocks backed by opaque nodes."""
    if source_map is None:
        return _map_blocks(blocks, docx_keys), set()
    entries = list(source_map.get("blocks", []))
    source_for_entry: dict[int, int] = {}
    for number, entry in enumerate(entries):
        try:
            source_value = entry.get("source_index", entry.get("index"))
            if source_value is None and entry.get("source_paragraph_index") is not None:
                paragraph_number = int(entry["source_paragraph_index"])
                paragraph_nodes = [
                    index for index, child in enumerate(children) if _local(child) == "p"
                ]
                source_value = paragraph_nodes[paragraph_number]
            source_index = int(source_value)
        except (AttributeError, IndexError, TypeError, ValueError):
            continue
        if 0 <= source_index < len(children):
            source_for_entry[number] = source_index
    text_entries = [
        (number, entry, source_for_entry[number])
        for number, entry in enumerate(entries)
        if number in source_for_entry and str(entry.get("kind", "text")) == "text"
    ]
    # A text-map entry is trusted only when it still points at the same
    # ordinary paragraph that was exported.  Otherwise a stale or tampered
    # map could redirect editable text into a table/equation and emit invalid
    # OOXML while appearing to succeed.
    seen_sources: set[int] = set()
    for number, entry, child_index in text_entries:
        child = children[child_index]
        if _local(child) != "p" or _opaque_kind(child) is not None:
            raise ValueError(f"text block map entry {number} does not point to an editable paragraph")
        expected = entry.get("text", entry.get("raw", ""))
        if _key(str(expected)) != _paragraph_key(child):
            raise ValueError(f"text block map entry {number} is stale for source index {child_index}")
        if child_index in seen_sources:
            raise ValueError(f"text block map reuses source index {child_index}")
        seen_sources.add(child_index)
    mapping: dict[int, int] = {}
    used: set[int] = set()
    text_blocks = [(index, block) for index, block in enumerate(blocks) if block.kind == "text"]
    # Pair exact source-map text in forward order.  A map can therefore carry
    # duplicate paragraphs without relying on fuzzy global matching.
    cursor = -1
    for block_index, block in text_blocks:
        candidates = [
            (number, entry, child_index)
            for number, entry, child_index in text_entries
            if number > cursor and child_index not in used
            and _key(str(entry.get("text", entry.get("raw", "")))) == block.key
        ]
        if len(candidates) == 1:
            _, _, child_index = candidates[0]
            mapping[block_index] = child_index
            used.add(child_index)
            cursor = candidates[0][0]
    # If the map does not carry text, retain the original matcher.  Otherwise
    # use it only for still-unmapped entries and preserve its conservative
    # ambiguity behavior.
    if not mapping:
        mapping = _map_blocks(blocks, docx_keys)
    else:
        remaining_blocks = [index for index, block in text_blocks if index not in mapping]
        remaining_children = [(index, key) for index, key in docx_keys if index not in used]
        fallback = _map_blocks([blocks[index] for index in remaining_blocks], remaining_children)
        for local, child_index in fallback.items():
            mapping[remaining_blocks[local]] = child_index

    # Baseline blocks are opaque when their explicit map index or id points at
    # an opaque source entry.  An opaque Markdown block is also paired with
    # opaque map entries by kind/order when no explicit marker was supplied.
    opaque_baseline: set[int] = set()
    opaque_entries = [
        (number, entry, source_for_entry[number])
        for number, entry in enumerate(entries)
        if number in source_for_entry and str(entry.get("kind", "text")) != "text"
    ]
    opaque_blocks = [index for index, block in enumerate(blocks) if block.kind != "text"]
    for block_index, (_, entry, child_index) in zip(opaque_blocks, opaque_entries):
        opaque_baseline.add(block_index)
        mapping[block_index] = child_index
    for number, entry in enumerate(entries):
        if str(entry.get("kind", "text")) == "text":
            continue
        marker = entry.get("markdown_index", entry.get("baseline_index", entry.get("block_index")))
        if marker is None:
            continue
        try:
            marker_index = int(marker)
        except (TypeError, ValueError):
            continue
        if 0 <= marker_index < len(blocks):
            opaque_baseline.add(marker_index)
            if number in source_for_entry:
                mapping[marker_index] = source_for_entry[number]
    return mapping, opaque_baseline


def _validate_opaque_map(
    children: list[etree._Element],
    all_children: list[etree._Element],
    source_map: dict[str, Any] | None,
) -> None:
    if source_map is None:
        return
    for entry in source_map.get("blocks", []):
        if not isinstance(entry, dict) or str(entry.get("kind", "text")) == "text":
            continue
        try:
            source_value = entry.get("source_index", entry.get("index"))
            if source_value is None:
                paragraph_number = int(entry["source_paragraph_index"])
                paragraph_nodes = [
                    index for index, child in enumerate(all_children) if _local(child) == "p"
                ]
                source_value = paragraph_nodes[paragraph_number]
            source_index = int(source_value)
        except (IndexError, KeyError, TypeError, ValueError):
            raise ValueError("opaque block map entry has no source index") from None
        if source_index < 0 or source_index >= len(all_children):
            raise ValueError(f"opaque block map source index is out of range: {source_index}")
        child = all_children[source_index]
        expected_kind = str(entry.get("kind"))
        actual_kind = _opaque_kind(child)
        if actual_kind != expected_kind:
            raise ValueError(f"opaque block map does not match source at index {source_index}")
        expected_signature = entry.get("signature")
        if expected_signature and expected_signature != _element_signature(child):
            raise ValueError(f"opaque block changed since block map export: {entry.get('id', expected_kind)}")


def _next_mapped(mapping: dict[int, int], start: int) -> int | None:
    return min((value for key, value in mapping.items() if key >= start), default=None)


def _blocks(text: str) -> list[_Block]:
    chunks = re.split(r"\n\s*\n", text.replace("\r\n", "\n"))
    blocks: list[_Block] = []
    for chunk in chunks:
        lines = [line.strip() for line in chunk.splitlines() if line.strip()]
        if not lines:
            continue
        if all(line.startswith("|") for line in lines) or (len(lines) == 1 and lines[0].startswith("![")):
            blocks.append(_Block("opaque", "\n".join(lines)))
            continue
        blocks.append(_Block("text", " ".join(lines)))
    return blocks


def _opcode_key(block: _Block) -> str:
    return block.kind + "\0" + block.key


def _key(raw: str) -> str:
    value = inline_key(raw)
    heading = HEADING_RE.match(raw.strip())
    if heading:
        value = inline_key(heading.group(2))
    return value


def _paragraph_key(element: etree._Element) -> str:
    if etree.QName(element).localname != "p":
        return ""
    if _opaque_kind(element) is not None:
        return ""
    text = "".join(node.text or "" for node in element.iter(f"{{{W}}}t"))
    return _key(text)


def _map_blocks(blocks: list[_Block], docx_keys: list[tuple[int, str]]) -> dict[int, int]:
    """Map Markdown blocks to source children without greedy duplicate capture."""
    source = [(index, key) for index, key in docx_keys if key]
    if not source:
        return {}

    from functools import lru_cache

    def pair_score(block: _Block, key: str, *, exact: bool) -> tuple[int, int, float] | None:
        if block.kind != "text" or not block.key:
            return None
        ratio = SequenceMatcher(a=block.key, b=key, autojunk=False).ratio()
        if not exact and ratio < 0.92:
            # Short headings and labels have too little text for the normal
            # ratio gate to survive one typo. Keep this exception bounded to a
            # single edit so short prose is not matched arbitrarily.
            if max(len(block.key), len(key)) > 8 or _edit_distance(block.key, key) > 1:
                return None
        return (1, 1 if exact else 0, ratio)

    def add_score(left: tuple[int, int, float], right: tuple[int, int, float]) -> tuple[int, int, float]:
        return left[0] + right[0], left[1] + right[1], left[2] + right[2]

    zero = (0, 0, 0.0)

    @lru_cache(maxsize=None)
    def solve(block_index: int, source_index: int, fuzzy_allowed: bool):
        if block_index >= len(blocks) or source_index >= len(source):
            return zero, ()
        block = blocks[block_index]
        child, key = source[source_index]
        options: list[tuple[tuple[int, int, float], tuple]] = []

        exact = block.kind == "text" and block.key == key
        score = pair_score(block, key, exact=exact)
        if score is not None and (exact or fuzzy_allowed):
            if not exact and source_index + 1 < len(source):
                next_key = source[source_index + 1][1]
                next_ratio = SequenceMatcher(a=block.key, b=next_key, autojunk=False).ratio()
                if score[2] - next_ratio < 0.04:
                    duplicate_run = (
                        source_index + 1 < len(source)
                        and block_index + 1 < len(blocks)
                        and blocks[block_index + 1].kind == "text"
                        and blocks[block_index + 1].key == source[source_index + 1][1]
                        and source[source_index + 1][1] == key
                    )
                    if not duplicate_run:
                        score = None
            if score is not None:
                tail_score, tail_path = solve(block_index + 1, source_index + 1, True)
                options.append((add_score(score, tail_score), ((block_index, child),) + tail_path))

        # Leave the current Markdown block unmatched without consuming source.
        tail_score, tail_path = solve(block_index + 1, source_index, fuzzy_allowed)
        options.append((tail_score, tail_path))

        # Skipping a source paragraph disables fuzzy matching for the current
        # block; a later exact anchor can still cross the gap.
        tail_score, tail_path = solve(block_index, source_index + 1, False)
        options.append((tail_score, tail_path))

        best_score, best_path = options[0]
        for candidate_score, candidate_path in options[1:]:
            if candidate_score > best_score:
                best_score, best_path = candidate_score, candidate_path
        return best_score, best_path

    _score, path = solve(0, 0, True)
    return {block_index: child_index for block_index, child_index in path}


def _edit_distance(left: str, right: str) -> int:
    """Return Levenshtein distance for bounded short-key fuzzy matching."""
    if len(left) < len(right):
        left, right = right, left
    previous = list(range(len(right) + 1))
    for i, left_char in enumerate(left, 1):
        current = [i]
        for j, right_char in enumerate(right, 1):
            current.append(min(
                current[-1] + 1,
                previous[j] + 1,
                previous[j - 1] + (left_char != right_char),
            ))
        previous = current
    return previous[-1]


def _anchor_before(blocks: list[_Block], mapping: dict[int, int], start: int) -> int:
    for index in range(start - 1, -1, -1):
        if index in mapping:
            return mapping[index]
    raise ValueError("inserted Markdown has no matched paragraph before it")


def _heading_style_ids(styles_xml: bytes | None) -> dict[int, str]:
    if not styles_xml:
        return {}
    root = etree.fromstring(styles_xml)
    found: dict[int, str] = {}
    for style in root.findall(f"{{{W}}}style"):
        name = style.find(f"{{{W}}}name")
        if name is None:
            continue
        match = re.fullmatch(r"heading ([1-6])", (name.get(f"{{{W}}}val") or "").strip().lower())
        if match:
            found[int(match.group(1))] = style.get(f"{{{W}}}styleId") or ""
    return {level: style_id for level, style_id in found.items() if style_id}


def _prototypes(children: list[etree._Element], heading_ids: dict[int, str]) -> dict:
    body_candidates = [
        child for child in children
        if _paragraph_key(child) and not _style_id(child)
    ]
    # Prefer an unstyled, non-bold paragraph after the first heading. This
    # avoids borrowing title/front-matter typography for body insertions when
    # both roles are represented as direct formatting rather than named
    # styles. Fall back to the longest ordinary paragraph only when no
    # neutral body sample exists.
    first_heading = next(
        (index for index, child in enumerate(children) if _is_heading_paragraph(child)),
        None,
    )
    preferred: list[etree._Element] = []
    for index, child in enumerate(children):
        if child in body_candidates and not _has_direct_bold(child) and (
            first_heading is not None and index > first_heading
        ):
            preferred.append(child)
    neutral = [child for child in body_candidates if not _has_direct_bold(child)]
    if not preferred and first_heading is not None:
        before_heading = [
            child for index, child in enumerate(children)
            if index < first_heading and child in neutral
        ]
        preferred = before_heading[-1:]
    body = (
        preferred[0]
        if preferred
        else (
            neutral[-1]
            if neutral
            else max(body_candidates, key=lambda child: len(_paragraph_key(child)), default=None)
        )
    )
    if body is None:
        body = next((child for child in children if etree.QName(child).localname == "p"), None)
    headings = {}
    for level, style_id in heading_ids.items():
        match = next((child for child in children if _style_id(child) == style_id), None)
        if match is not None:
            headings[level] = match
    return {
        "body": body,
        "headings": headings,
        "heading_ids": heading_ids,
        "list_headings": _list_headings(children),
    }


def _is_heading_paragraph(element: etree._Element) -> bool:
    style = _style_id(element).casefold()
    return style.startswith("heading") or element.find(f"{{{W}}}pPr/{{{W}}}numPr") is not None


def _has_direct_bold(element: etree._Element) -> bool:
    return any(
        run.find(f"{{{W}}}rPr/{{{W}}}b") is not None
        or run.find(f"{{{W}}}rPr/{{{W}}}bCs") is not None
        for run in element.findall(f"{{{W}}}r")
    )


def _style_id(element: etree._Element) -> str:
    style = element.find(f"{{{W}}}pPr/{{{W}}}pStyle")
    return style.get(f"{{{W}}}val") if style is not None else ""


def _list_headings(children: list[etree._Element]) -> dict[int, etree._Element]:
    """Numbered bold paragraphs keyed by numbering level.

    Documents without Heading styles still encode section and subsection
    formatting in those paragraphs. The first example of each level is kept.
    """

    return numbered_bold_prototypes(children)


def _apply(root: etree._Element, operations: list[tuple]) -> None:
    body = root.find(f"{{{W}}}body")
    children = [child for child in list(body) if etree.QName(child).localname != "sectPr"]
    refs = {index: child for index, child in enumerate(children)}

    def opaque_after(start: int | None) -> etree._Element | None:
        threshold = -1 if start is None else start
        return next(
            (node for index, node in refs.items() if index > threshold and _opaque_kind(node) is not None),
            None,
        )

    insertion_cursors: dict[int | None, etree._Element | None] = {}
    for operation in operations:
        if operation[0] == "insert":
            next_anchor, blocks, prototypes, style_anchor = operation[1:]
            next_node = refs.get(next_anchor) if next_anchor is not None else opaque_after(style_anchor)
            style_node = refs.get(style_anchor) if style_anchor is not None else prototypes.get("body")
            previous = insertion_cursors.get(next_anchor)
            if previous is not None:
                insertion_cursors[next_anchor] = _insert_after(previous, blocks, prototypes, style_node=style_node)
            else:
                insertion_cursors[next_anchor] = _insert_before(next_node, blocks, prototypes, body, style_node=style_node)
            continue
        if operation[0] == "delete":
            for index in operation[1]:
                node = refs.get(index)
                if node is not None and node.getparent() is body:
                    body.remove(node)
            continue
        indexes, new_blocks, prototypes = operation[1], operation[2], operation[3]
        first = refs[indexes[0]]
        _write_paragraph(first, new_blocks[0].raw, prototypes, replacement=first)
        for extra in indexes[1:]:
            extra_node = refs[extra]
            parent = extra_node.getparent()
            if parent is not None:
                parent.remove(extra_node)
        if len(new_blocks) > 1:
            next_anchor = operation[4] if len(operation) > 4 else None
            next_node = refs.get(next_anchor) if next_anchor is not None else opaque_after(max(indexes))
            # Place replacement siblings immediately before the next matched
            # source node. This keeps tables, drawings, and other opaque
            # nodes in their original relative location.
            _insert_before(next_node, new_blocks[1:], prototypes, body, style_node=first)


def _insert_after(
    anchor: etree._Element,
    blocks: list[_Block],
    prototypes: dict,
    *,
    style_node: etree._Element | None = None,
) -> etree._Element:
    cursor = anchor
    for block in blocks:
        paragraph = etree.Element(f"{{{W}}}p")
        source_style = style_node if style_node is not None else anchor
        properties = _properties_for(block.raw, prototypes, source_style)
        if properties is not None:
            paragraph.append(properties)
        heading = _is_heading(block.raw)
        run_source = (
            _heading_prototype(block.raw, prototypes)
            if heading
            else (prototypes.get("body") if prototypes.get("body") is not None else source_style)
        )
        if run_source is None:
            run_source = source_style
        _add_runs(
            paragraph,
            _display_text(block.raw),
            _run_properties(run_source),
            extra_marks={"b"} if heading else frozenset(),
        )
        cursor.addnext(paragraph)
        cursor = paragraph
    return cursor


def _insert_before(
    anchor: etree._Element | None,
    blocks: list[_Block],
    prototypes: dict,
    body: etree._Element,
    *,
    style_node: etree._Element | None = None,
) -> etree._Element | None:
    """Insert blocks before ``anchor`` (or immediately before ``sectPr``)."""
    def style_source(cursor: etree._Element | None, fallback: etree._Element | None) -> etree._Element | None:
        if style_node is not None:
            return style_node
        if cursor is not None:
            return cursor
        return fallback

    if anchor is None:
        sect = body.find(f"{{{W}}}sectPr")
        cursor = body[-1] if sect is None and len(body) else None
        for block in blocks:
            paragraph = etree.Element(f"{{{W}}}p")
            properties = _properties_for(block.raw, prototypes, style_source(cursor, body))
            if properties is not None:
                paragraph.append(properties)
            heading = _is_heading(block.raw)
            run_source = (
                _heading_prototype(block.raw, prototypes)
                if heading
                else (prototypes.get("body") if prototypes.get("body") is not None else style_source(cursor, prototypes.get("body")))
            )
            if run_source is None:
                run_source = prototypes.get("body")
            _add_runs(paragraph, _display_text(block.raw), _run_properties(run_source), extra_marks={"b"} if heading else frozenset())
            if sect is not None:
                sect.addprevious(paragraph)
            else:
                body.append(paragraph)
            cursor = paragraph
        return cursor
    cursor = anchor.getprevious()
    last: etree._Element | None = None
    for block in blocks:
        paragraph = etree.Element(f"{{{W}}}p")
        properties = _properties_for(block.raw, prototypes, style_source(cursor, prototypes.get("body")))
        if properties is not None:
            paragraph.append(properties)
        heading = _is_heading(block.raw)
        run_source = (
            _heading_prototype(block.raw, prototypes)
            if heading
            else (prototypes.get("body") if prototypes.get("body") is not None else style_source(cursor, prototypes.get("body")))
        )
        _add_runs(paragraph, _display_text(block.raw), _run_properties(run_source), extra_marks={"b"} if heading else frozenset())
        anchor.addprevious(paragraph)
        cursor = paragraph
        last = paragraph
    return last


def _write_paragraph(paragraph: etree._Element, raw: str, prototypes: dict, *, replacement: etree._Element) -> None:
    for child in list(paragraph):
        if etree.QName(child).localname != "pPr":
            paragraph.remove(child)
    _add_runs(paragraph, _display_text(raw), _run_properties(replacement))


def _is_heading(raw: str) -> bool:
    return HEADING_RE.match(raw.strip()) is not None


def _heading_prototype(raw: str, prototypes: dict) -> etree._Element | None:
    match = HEADING_RE.match(raw.strip())
    if match is None:
        return None
    level = min(len(match.group(1)), 6)
    headings = prototypes["headings"]
    if level in headings:
        return headings[level]
    if 1 in headings:
        return headings[1]
    lists = prototypes.get("list_headings") or {}
    # Markdown ## is a section; ### is the subsection under it.
    target = 0 if level <= 2 else level - 2
    if target in lists:
        return lists[target]
    if lists:
        return lists[min(lists)]
    return None


def _properties_for(raw: str, prototypes: dict, anchor: etree._Element) -> etree._Element | None:
    heading = HEADING_RE.match(raw.strip())
    if heading:
        level = min(len(heading.group(1)), 6)
        prototype = _heading_prototype(raw, prototypes)
        if prototype is not None and level in prototypes["headings"]:
            return _copy_paragraph_properties(prototype)
        if prototype is not None:
            return _copy_paragraph_properties(prototype, drop_style=True)
        style_id = prototypes["heading_ids"].get(level) or prototypes["heading_ids"].get(1)
        if style_id:
            properties = etree.Element(f"{{{W}}}pPr")
            etree.SubElement(properties, f"{{{W}}}pStyle").set(f"{{{W}}}val", style_id)
            return properties
    # Placement and formatting are separate concerns.  A body insertion that
    # happens to be placed before a heading must still use the body role's
    # prototype instead of inheriting heading or caption formatting.
    source = prototypes.get("body") if prototypes.get("body") is not None else anchor
    if source is None:
        return None
    return _copy_paragraph_properties(source)


def _copy_paragraph_properties(
    paragraph: etree._Element,
    *,
    drop_style: bool = False,
) -> etree._Element | None:
    properties = paragraph.find(f"{{{W}}}pPr")
    if properties is None:
        return None
    copied = copy.deepcopy(properties)
    had_numbering = bool(copied.findall(f"{{{W}}}numPr"))
    for numbering in copied.findall(f"{{{W}}}numPr"):
        copied.remove(numbering)
    for style in copied.findall(f"{{{W}}}pStyle"):
        if (style.get(f"{{{W}}}val") or "").casefold() in {"listparagraph", "list paragraph"}:
            copied.remove(style)
    if drop_style:
        for style in list(copied.findall(f"{{{W}}}pStyle")):
            copied.remove(style)
    elif had_numbering:
        # Explicit numbering identifies list identity. Indentation alone is
        # ordinary semantic paragraph formatting and must remain intact.
        style = copied.find(f"{{{W}}}pStyle")
        if style is not None and re.search(r"(?:list|bullet|number)", style.get(f"{{{W}}}val", ""), re.IGNORECASE):
            copied.remove(style)
    else:
        # A list prototype must not make inserted body paragraphs inherit its
        # list role even when it has no explicit indentation.
        for style in list(copied.findall(f"{{{W}}}pStyle")):
            if (style.get(f"{{{W}}}val") or "").casefold() in {"listparagraph", "list paragraph"}:
                copied.remove(style)
    return copied


def _display_text(raw: str) -> str:
    heading = HEADING_RE.match(raw.strip())
    text = heading.group(2).strip() if heading else raw
    return text


def _reject_inline_atoms(raw: str) -> None:
    """Reject edits that would silently drop lossless inline OOXML atoms."""
    try:
        items = parse_inline(_display_text(raw))
    except ValueError as exc:
        raise ValueError(f"edited Markdown has invalid inline syntax: {exc}") from exc
    if any(isinstance(item, Atom) for item in items):
        raise ValueError("markdown delta edits an inline atom; preserve it in the source bundle first")


def _run_properties(paragraph: etree._Element | None) -> etree._Element | None:
    if paragraph is None:
        return None
    properties = paragraph.find(f"{{{W}}}r/{{{W}}}rPr")
    if properties is None:
        return None
    copied = copy.deepcopy(properties)
    # Highlight belongs to the sample marker (and can make inserted text look
    # like a tracked placeholder).  Typography and emphasis remain intact.
    for node in list(copied.findall(f"{{{W}}}highlight")):
        copied.remove(node)
    return copied


def _add_runs(
    paragraph: etree._Element,
    text: str,
    base_properties: etree._Element | None,
    extra_marks: frozenset[str] = frozenset(),
) -> None:
    for run in display(text):
        marks = set(extra_marks)
        if run.bold:
            marks.add("b")
        if run.italic:
            marks.add("i")
        if run.subscript:
            marks.add("sub")
        elif run.superscript:
            marks.add("sup")
        if run.code:
            marks.add("code")
        _append_run(paragraph, run.text, base_properties, marks)


def _append_run(paragraph: etree._Element, text: str, base_properties: etree._Element | None, marks: set[str]) -> None:
    if not text:
        return
    run = etree.SubElement(paragraph, f"{{{W}}}r")
    properties = copy.deepcopy(base_properties) if base_properties is not None else etree.Element(f"{{{W}}}rPr")
    for name in ("b", "i", "vertAlign"):
        for node in properties.findall(f"{{{W}}}{name}"):
            properties.remove(node)
    if "b" in marks:
        etree.SubElement(properties, f"{{{W}}}b")
    if "i" in marks:
        etree.SubElement(properties, f"{{{W}}}i")
    if "sub" in marks or "sup" in marks:
        align = etree.SubElement(properties, f"{{{W}}}vertAlign")
        align.set(f"{{{W}}}val", "subscript" if "sub" in marks else "superscript")
    if "code" in marks:
        fonts = properties.find(f"{{{W}}}rFonts")
        if fonts is None:
            fonts = etree.SubElement(properties, f"{{{W}}}rFonts")
        for attribute in ("ascii", "hAnsi", "cs"):
            fonts.set(f"{{{W}}}{attribute}", "Consolas")
    if len(properties):
        run.append(properties)
    node = etree.SubElement(run, f"{{{W}}}t")
    if text.startswith(" ") or text.endswith(" "):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text


def roundtrip_source_docx(
    source: Path,
    workdir: Path,
    output: Path,
    *,
    edited: Path | None = None,
    baseline: Path | None = None,
    redline: Path | None = None,
    block_map: Path | None = None,
    allow_delete: bool = False,
    force: bool = False,
    render_engine: str = "none",
    report: Path | None = None,
) -> dict[str, Any]:
    """Export, patch, and redline one source DOCX without template assembly.

    The first invocation writes ``baseline.md`` and ``block-map.json`` under
    ``workdir`` and copies the editable Markdown to ``full.md``. Subsequent
    invocations reuse those files, so callers can edit ``full.md`` and rerun
    the same command. All candidate DOCX/report files are staged until patch,
    redline, and optional rendering checks succeed.
    """
    source, workdir, output = Path(source), Path(workdir), Path(output)
    baseline = Path(baseline) if baseline is not None else source
    workdir.mkdir(parents=True, exist_ok=True)
    if not source.is_file():
        raise FileNotFoundError(source)
    if not baseline.is_file():
        raise FileNotFoundError(baseline)
    if render_engine not in {"none", "auto", "word", "libreoffice"}:
        raise ValueError("render_engine must be one of: none, auto, word, libreoffice")
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    redline = Path(redline) if redline is not None else output.with_name(output.stem + ".tracked" + output.suffix)
    report = Path(report) if report is not None else workdir / "roundtrip.manifest.json"
    if len({output.resolve(), redline.resolve(), report.resolve()}) != 3:
        raise ValueError("roundtrip output, redline, and report paths must be distinct")
    if output.resolve() == source.resolve() or redline.resolve() == source.resolve():
        raise ValueError("roundtrip outputs must not overwrite the source package")
    if not force:
        for target in (output, redline, report):
            if target.exists():
                raise FileExistsError(f"Output exists; pass --force to overwrite: {target}")
    validate_output_path(output)
    validate_output_path(redline)
    validate_output_path(report, label="report")

    baseline_markdown = workdir / "baseline.md"
    full_markdown = workdir / "full.md"
    block_map = Path(block_map) if block_map is not None else workdir / "block-map.json"
    exported = False
    if not baseline_markdown.is_file() or not block_map.is_file():
        from .docx_export import docx_to_markdown

        docx_to_markdown(
            source,
            output=baseline_markdown,
            media_dir=workdir,
            track_changes="accept",
            force=True,
        )
        if not full_markdown.exists():
            shutil.copyfile(baseline_markdown, full_markdown)
        export_block_map(source, block_map)
        exported = True
    if edited is None:
        edited = full_markdown
    edited = Path(edited)
    if not edited.is_file():
        raise FileNotFoundError(edited)

    with tempfile.TemporaryDirectory(prefix=".docforge-roundtrip-", dir=workdir) as temporary:
        stage = Path(temporary)
        staged_output = stage / output.name
        delta = apply_markdown_delta(
            source,
            baseline_markdown,
            edited,
            staged_output,
            overwrite=True,
            allow_delete=allow_delete,
            block_map=block_map,
        )
        from ..docxdiff import create_tracked_docx

        staged_redline = stage / redline.name
        redline_report = stage / (redline.stem + ".report.json")
        redline_summary = create_tracked_docx(
            baseline,
            staged_output,
            staged_redline,
            author="docforge-roundtrip",
            overwrite=True,
            accept_baseline=True,
            report_path=redline_report,
            revision_date="1970-01-01T00:00:00Z",
            workers=1,
        )
        render_receipt: dict[str, Any] = {"enabled": False}
        if render_engine != "none":
            from ..docxmd.verify import render_pdfs

            pdfs = render_pdfs([staged_output, staged_redline], stage / "render", render_engine)
            render_receipt = {
                "enabled": True,
                "engine": render_engine,
                "documents": [
                    {"path": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
                    for path in pdfs
                ],
            }
        redline_report_payload = json.loads(redline_report.read_text(encoding="utf-8"))
        redline_report_payload["baseline"] = str(baseline.resolve())
        redline_report_payload["current"] = str(output.resolve())
        redline_report_payload["output"] = str(redline.resolve())
        payload = {
            "schema": "docforge.source-roundtrip.v1",
            "source": {"path": str(source.resolve()), "sha256": hashlib.sha256(source.read_bytes()).hexdigest()},
            "baseline": {"path": str(baseline.resolve()), "sha256": hashlib.sha256(baseline.read_bytes()).hexdigest()},
            "export": {
                "baseline_markdown": str(baseline_markdown.resolve()),
                "edited_markdown": str(edited.resolve()),
                "block_map": str(block_map.resolve()),
                "performed": exported,
            },
            "patch": delta,
            "output": {"path": str(output.resolve()), "sha256": hashlib.sha256(staged_output.read_bytes()).hexdigest()},
            "redline": {
                "path": str(redline.resolve()),
                "sha256": hashlib.sha256(staged_redline.read_bytes()).hexdigest(),
                "summary": redline_summary,
                "accepted_view_equivalent": True,
                "report": redline_report_payload,
            },
            "render": render_receipt,
        }
        staged_report = stage / report.name
        staged_report.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        output.parent.mkdir(parents=True, exist_ok=True)
        redline.parent.mkdir(parents=True, exist_ok=True)
        report.parent.mkdir(parents=True, exist_ok=True)
        targets = ((staged_output, output), (staged_redline, redline), (staged_report, report))
        backups: dict[Path, Path] = {}
        installed: list[Path] = []
        try:
            for _staged, target in targets:
                if target.exists():
                    backup = stage / (target.name + ".previous")
                    shutil.copy2(target, backup)
                    backups[target] = backup
            for staged_path, target in targets:
                os.replace(staged_path, target)
                installed.append(target)
        except Exception:
            for target in installed:
                backup = backups.get(target)
                if backup is not None and backup.exists():
                    os.replace(backup, target)
                elif target.exists():
                    target.unlink()
            raise
    return payload


source_roundtrip = roundtrip_source_docx
