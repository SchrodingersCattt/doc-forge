"""Pieces shared by the DOCX reader and writer."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .ooxml import Token, format_tokens, format_value, merge, parse_token, split_words

SCHEMA = "docforge.docxmd.v1"
BIB_SCHEMA = "docforge.bibliography.v2"

REVISION_ELEMENTS = {
    "ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange", "sectPrChange",
    "tblPrChange", "trPrChange", "tcPrChange", "numberingChange", "cellIns", "cellDel",
    "cellMerge", "tblGridChange",
}

# Shortcut marks force one property on top of the paragraph's base run.
FORCED: dict[str, list[Token]] = {
    "b": [("-b", None), ("b", None)],
    "i": [("-i", None), ("i", None)],
    "sup": [("-vertAlign", None), ("vertAlign@val", "superscript")],
    "sub": [("-vertAlign", None), ("vertAlign@val", "subscript")],
    "u": [("-u", None), ("u@val", "single")],
    "strike": [("-strike", None), ("strike", None)],
    "hl": [("-highlight", None), ("highlight@val", "yellow")],
}
SHORTCUT_ELEMENT = {"b": "b", "i": "i", "sup": "vertAlign", "sub": "vertAlign", "u": "u", "strike": "strike", "hl": "highlight"}


def clean_tokens(tokens: list[Token]) -> list[Token]:
    """Drop editing-session metadata: rsids, paragraph ids, revision ids."""
    result = []
    for key, value in tokens:
        path, _, attribute = key.partition("@")
        if attribute.startswith("rsid") or attribute in {"w14:paraId", "w14:textId"}:
            continue
        if attribute == "id" and any(segment.split("#")[0] in REVISION_ELEMENTS for segment in path.split(".")):
            continue
        result.append((key, value))
    return result


def top_name(key: str) -> str:
    return key.lstrip("-").split("@", 1)[0].split(".", 1)[0].split("#", 1)[0]


def run_tokens_for(base: list[Token], marks) -> list[Token]:
    """Run properties for ``marks`` on top of ``base``."""
    tokens = list(base)
    for mark in marks:
        if mark[0] == "span":
            tokens = merge(tokens, parse_tokens_text(mark[1]))
    for mark in marks:
        if mark[0] in FORCED:
            tokens = merge(tokens, FORCED[mark[0]])
    return tokens


def parse_tokens_text(text: str) -> list[Token]:
    return [parse_token(word) for word in split_words(text)]


# ----------------------------------------------------------------------------
# Numeric citations: ``1–3,7`` <-> [1, 2, 3, 7]

_CITATION_TEXT = re.compile(r"^\d+(?:[–-]\d+)?(?:,\d+(?:[–-]\d+)?)*$")


def parse_citation_numbers(text: str) -> list[int] | None:
    if not _CITATION_TEXT.match(text):
        return None
    numbers: list[int] = []
    for piece in text.split(","):
        if "–" in piece or "-" in piece:
            start, end = re.split("[–-]", piece)
            if int(end) <= int(start):
                return None
            numbers.extend(range(int(start), int(end) + 1))
        else:
            numbers.append(int(piece))
    return numbers


def format_citation_numbers(numbers: list[int]) -> str:
    parts: list[str] = []
    index = 0
    while index < len(numbers):
        end = index
        while end + 1 < len(numbers) and numbers[end + 1] == numbers[end] + 1:
            end += 1
        if end - index >= 2:
            parts.append(f"{numbers[index]}–{numbers[end]}")
        else:
            parts.extend(str(number) for number in numbers[index:end + 1])
        index = end + 1
    return ",".join(parts)


# ----------------------------------------------------------------------------
# Paragraph attribute lines: ``{: .Class spacing@after=240 r:b tc:vAlign=center role=title}``

ATTR_PREFIXES = ("r:", "tc:", "tr:")


@dataclass
class Attrs:
    cls: str | None = None
    p: list[Token] = field(default_factory=list)
    r: list[Token] = field(default_factory=list)
    tc: list[Token] = field(default_factory=list)
    tr: list[Token] = field(default_factory=list)
    params: dict[str, str] = field(default_factory=dict)

    def empty(self) -> bool:
        return not (self.cls or self.p or self.r or self.tc or self.tr or self.params)


PARAM_KEYS = {
    "role", "default", "style", "label", "grid", "id", "author", "initials", "date", "done",
    "parent", "rels",
}


def format_attrs(attrs: Attrs) -> str:
    parts = []
    if attrs.cls:
        parts.append("." + attrs.cls)
    if attrs.p:
        parts.append(format_tokens(attrs.p))
    for prefix, tokens in (("r:", attrs.r), ("tc:", attrs.tc), ("tr:", attrs.tr)):
        if tokens:
            parts.append(format_tokens(tokens, prefix))
    for key, value in attrs.params.items():
        parts.append(f"{key}={format_value(value)}")
    return "{: " + " ".join(parts) + "}" if parts else "{:}"


def parse_attrs(text: str) -> Attrs:
    text = text.strip()
    if text.startswith("{:") and text.endswith("}"):
        text = text[2:-1]
    attrs = Attrs()
    for word in split_words(text):
        if word.startswith(".") and len(word) > 1:
            attrs.cls = word[1:]
            continue
        negative = word.startswith("-")
        bare = word[1:] if negative else word
        for prefix, target in (("r:", attrs.r), ("tc:", attrs.tc), ("tr:", attrs.tr)):
            if bare.startswith(prefix):
                target.append(parse_token(("-" if negative else "") + bare[len(prefix):]))
                break
        else:
            key = bare.split("=", 1)[0]
            if not negative and key in PARAM_KEYS and "=" in bare:
                attrs.params[key] = parse_token(bare)[1] or ""
            else:
                attrs.p.append(parse_token(word))
    return attrs


ATTR_LINE = re.compile(r"^\{:(?:\s.*)?\}$")


__all__ = [
    "ATTR_LINE", "Attrs", "BIB_SCHEMA", "FORCED", "SCHEMA", "SHORTCUT_ELEMENT", "clean_tokens",
    "format_attrs", "format_citation_numbers", "parse_attrs", "parse_citation_numbers",
    "parse_tokens_text", "run_tokens_for", "top_name",
]
