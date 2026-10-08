"""Citation rendering, separate from chemical subscripts and superscripts.

A citation is recognized only by the ``\\cite`` / ``\\citep`` token that produced
it. Superscript styles travel through the renderer as private-use sentinels.
Formula scripts are ordinary ``w:vertAlign`` runs and never carry those
sentinels, so materializing a citation cannot rewrite ``NH_4^+``.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import Iterable, Mapping

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .blocks import Block

CITATION_RE = re.compile(r"\\citep?\{([^{}]+)\}")
_SENTINEL_RE = re.compile(r"(\ue000(?:S?[0-9]+)(?:[–,](?:S?[0-9]+))*\ue001)")
_FORMATS = ("template", "bracketed", "superscript", "superscript-bracketed")


@dataclass(frozen=True)
class CitationStyle:
    """Where a citation sits and whether its label is wrapped in brackets.

    ``superscript`` and ``brackets`` are independent. A Chinese proposal uses
    both: the Word run is superscript and its text is ``[1]``. A journal
    template may use a bare superscript ``1``, or a baseline ``[1]``.
    """

    name: str
    superscript: bool
    brackets: bool

    def visible(self, label: str) -> str:
        if self.brackets:
            return f"[{label}]"
        return label

    def embed(self, label: str) -> str:
        """Text stored before Word runs exist.

        Superscript styles hide the label inside sentinels. Brackets are
        applied later, when that sentinel becomes a run, so they are not
        another superscript digit sitting next to a formula.
        """

        if self.superscript:
            return f"\ue000{label}\ue001"
        return self.visible(label)

    @classmethod
    def by_name(cls, name: str) -> CitationStyle:
        style = _STYLES.get(name)
        if style is None:
            allowed = ", ".join(_FORMATS)
            raise ValueError(f"citation_format must be one of: {allowed}")
        return style


_STYLES = {
    "bracketed": CitationStyle("bracketed", superscript=False, brackets=True),
    "superscript": CitationStyle("superscript", superscript=True, brackets=False),
    "superscript-bracketed": CitationStyle("superscript-bracketed", superscript=True, brackets=True),
}
CitationStyle.BRACKETED = _STYLES["bracketed"]
CitationStyle.SUPERSCRIPT = _STYLES["superscript"]
CitationStyle.SUPERSCRIPT_BRACKETED = _STYLES["superscript-bracketed"]


def format_citation_labels(labels: Iterable[str | int]) -> str:
    """Sort and collapse consecutive numeric or SI-prefixed citation labels."""

    grouped: dict[str, set[int]] = {"": set(), "S": set()}
    for label in labels:
        value = str(label)
        prefix = "S" if value.startswith("S") else ""
        grouped[prefix].add(int(value.removeprefix("S")))

    ranges: list[str] = []
    for prefix in ("", "S"):
        values = sorted(grouped[prefix])
        start = end = None
        for value in values + [None]:
            if start is None:
                start = end = value
            elif value is not None and value == end + 1:
                end = value
            else:
                ranges.append(f"{prefix}{start}" if start == end else f"{prefix}{start}–{prefix}{end}")
                start = end = value
    return ",".join(ranges)


def _citation_keys(cluster: str) -> list[str]:
    keys: list[str] = []
    for group in CITATION_RE.findall(cluster):
        keys.extend(key.strip() for key in group.split(",") if key.strip())
    return keys


def replace_citations(text: str, mapping: Mapping[str, str | int], style: CitationStyle) -> str:
    """Replace one citation or a run of adjacent citations.

    Adjacent ``\\cite`` commands are one cluster. Their numbers are sorted
    and consecutive values become an en-dash range, including when the
    Markdown wrote each key as its own command.
    """

    cluster = (
        r"(?:\s*\\citep?\{[^{}]+\})+"
        if style.superscript
        else r"\\citep?\{[^{}]+\}(?:\s*\\citep?\{[^{}]+\})*"
    )

    def replace(match: re.Match[str]) -> str:
        keys = _citation_keys(match.group(0))
        missing = [key for key in keys if key not in mapping]
        if missing:
            raise ValueError(f"Unknown citation key(s): {', '.join(missing)}")
        label = format_citation_labels(mapping[key] for key in keys)
        return style.embed(label)

    return re.sub(cluster, replace, text)


def replace_block_citations(
    blocks: Iterable[Block], mapping: Mapping[str, str | int], style: CitationStyle
) -> tuple[Block, ...]:
    return tuple(
        Block(
            kind=block.kind,
            text=replace_citations(block.text, mapping, style),
            level=block.level,
            rows=tuple(tuple(replace_citations(value, mapping, style) for value in row) for row in block.rows),
            language=block.language,
            path=block.path,
            options=block.options,
        )
        for block in blocks
    )


def materialize_citations(document, style: CitationStyle) -> int:
    """Replace citation sentinels with runs of this style.

    Baseline styles are already final text. Superscript styles become
    ``w:vertAlign=superscript`` runs whose text is ``style.visible(label)``.
    """

    if not style.superscript:
        return 0
    converted = 0
    for parent in list(document._element.body.iter(qn("w:r"))):
        text_nodes = [child for child in parent if child.tag == qn("w:t") and child.text and _SENTINEL_RE.search(child.text)]
        if not text_nodes:
            continue
        grandparent = parent.getparent()
        if grandparent is None:
            continue
        base_properties = parent.find(qn("w:rPr"))
        replacements = []
        for node in text_nodes:
            for fragment in _SENTINEL_RE.split(node.text):
                if not fragment:
                    continue
                run = OxmlElement("w:r")
                if base_properties is not None:
                    run.append(copy.deepcopy(base_properties))
                text = OxmlElement("w:t")
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                if fragment.startswith("\ue000"):
                    text.text = style.visible(fragment[1:-1])
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
