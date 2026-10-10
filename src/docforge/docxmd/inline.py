"""Inline Markdown syntax for the runs of one Word paragraph.

A paragraph is a sequence of :class:`Text` and :class:`Atom` items, each
carrying ``marks`` (formatting relative to the paragraph's base run).
:func:`write` renders items as one Markdown line and :func:`parse` reads it
back; ``parse(write(items)) == normalize(items)`` is checked on export.

Marks, outermost first::

    {++ins++}{author=.. date=..}  {--del--}{...}   tracked insertion/deletion
    [text](url)                                    hyperlink
    [text]{color=C00000 -b}                        any other run properties
    ==text==                                       yellow highlight
    <u>text</u>  ~~text~~  **text**  *text*        underline, strike, bold, italic
    <sup>text</sup>  <sub>text</sub>               vertical alignment

Atoms: ``<tab/>``, ``<br/>``, ``<br type="page"/>``, ``<cr/>``,
``<sym font=".." char=".."/>``, ``<nbhyphen/>``, ``<shy/>``,
``![alt](media/x.png){cx=.. cy=..}``, ``\\citep{key1,key2}``, field
characters ``<fld-begin/>`` ``<instr code=".."/>`` ``<fld-sep/>``
``<fld-end/>``, comment milestones ``<comment-start id=".."/>``
``<comment-end id=".."/>``, bookmarks ``<bookmark-start name=".."/>``
``<bookmark-end name=".."/>``, ``<annotation-ref/>`` and the raw fallback
``<ooxml level="run|node" rels="..">..</ooxml>``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

from .ooxml import format_tokens, format_value, parse_tokens, split_words

PRECEDENCE = {
    "rev": 0, "link": 1, "span": 2, "hl": 3, "u": 4, "strike": 5, "b": 6, "i": 7,
    "sup": 8, "sub": 8,
}
MILESTONES = {"comment-start", "comment-end", "comment-ref", "bookmark-start", "bookmark-end", "node"}
_ESCAPABLE = set("\\*_[]{}<>|#=~:!&`+-().")
_TAG = re.compile(r"<(/?)([a-z][a-z0-9-]*)((?:\s+[a-zA-Z_:][\w:.-]*=\"(?:[^\"\\]|\\.)*\")*)\s*(/?)>")
_ATTR = re.compile(r"([a-zA-Z_:][\w:.-]*)=\"((?:[^\"\\]|\\.)*)\"")
_ENTITY = re.compile(r"&#(x[0-9A-Fa-f]+|\d+);")
_TOKEN_WORD = re.compile(r"^-?\.?[A-Za-z_][\w:.#@-]*(=.*)?$")
_PAIRED = {"sup", "sub", "u"}
_VOID = {
    "tab", "br", "cr", "sym", "nbhyphen", "shy", "fld-begin", "fld-sep", "fld-end",
    "instr", "annotation-ref", "comment-start", "comment-end", "comment-ref", "footnote-ref",
    "endnote-ref", "bookmark-start", "bookmark-end",
}


@dataclass(frozen=True)
class Text:
    text: str
    marks: tuple = ()


@dataclass(frozen=True)
class Atom:
    kind: str
    attrs: tuple = ()
    marks: tuple = ()
    body: str | None = None

    def get(self, name: str, default: str | None = None) -> str | None:
        for key, value in self.attrs:
            if key == name:
                return value
        return default


def sort_marks(marks) -> tuple:
    return tuple(sorted(set(marks), key=lambda mark: (PRECEDENCE[mark[0]], repr(mark))))


def normalize(items) -> list:
    result: list = []
    for item in items:
        if isinstance(item, Text):
            if not item.text:
                continue
            item = Text(item.text, sort_marks(item.marks))
            if result and isinstance(result[-1], Text) and result[-1].marks == item.marks:
                result[-1] = Text(result[-1].text + item.text, item.marks)
                continue
        else:
            marks = () if item.kind in MILESTONES or (item.kind == "ooxml" and item.get("level") == "node") else sort_marks(item.marks)
            item = replace(item, marks=marks)
        result.append(item)
    return result


# ----------------------------------------------------------------------------
# Writing


def _quote_attr(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _escape(text: str, strict: bool) -> str:
    out: list[str] = []
    for index, char in enumerate(text):
        nxt = text[index + 1] if index + 1 < len(text) else ""
        prev = text[index - 1] if index else ""
        if char in "\\*<{}|":
            out.append("\\" + char)
        elif char in "[]" and strict:
            out.append("\\" + char)
        elif char in "=~" and (strict or nxt == char or prev == char):
            out.append("\\" + char)
        elif char == "&" and nxt == "#":
            out.append("\\&")
        elif char == "!" and (strict or nxt == "["):
            out.append("\\!")
        elif ord(char) < 0x20 or char in "\u2028\u2029":
            out.append(f"&#x{ord(char):X};")
        else:
            out.append(char)
    return "".join(out)


_OPENERS = {
    "link": "[", "span": "[", "hl": "==", "u": "<u>", "strike": "~~", "b": "**", "i": "*",
    "sup": "<sup>", "sub": "<sub>",
}


def _open(mark) -> str:
    if mark[0] == "rev":
        return "{++" if mark[1] == "ins" else "{--"
    return _OPENERS[mark[0]]


def _close(mark) -> str:
    kind = mark[0]
    if kind == "rev":
        closer = "++}" if mark[1] == "ins" else "--}"
        attrs = " ".join(f"{key}={_quote_attr(value)}" for key, value in mark[2])
        return closer + ("{" + attrs + "}" if attrs else "")
    if kind == "link":
        return "](" + mark[1].replace(")", "%29").replace(" ", "%20") + ")"
    if kind == "span":
        return "]{" + mark[1] + "}"
    return {
        "hl": "==", "u": "</u>", "strike": "~~", "b": "**", "i": "*", "sup": "</sup>",
        "sub": "</sub>",
    }[kind]


def _atom_text(atom: Atom, strict: bool) -> str:
    if atom.kind == "image":
        alt = _escape(atom.get("alt", "") or "", True)
        src = (atom.get("src") or "").replace(" ", "%20")
        rest = [(key, value) for key, value in atom.attrs if key not in {"alt", "src"}]
        attrs = " ".join(f"{key}={format_value(value)}" for key, value in rest)
        return f"![{alt}]({src})" + ("{" + attrs + "}" if attrs else "")
    if atom.kind == "cite":
        return "\\citep{" + (atom.get("keys") or "") + "}"
    attrs = "".join(f" {key}={_quote_attr(value)}" for key, value in atom.attrs)
    if atom.kind == "ooxml":
        return f"<ooxml{attrs}>{atom.body or ''}</ooxml>"
    return f"<{atom.kind}{attrs}/>"


def _render(items, strict: bool) -> str:
    out: list[str] = []
    stack: list = []
    for item in items:
        if isinstance(item, Atom) and (item.kind in MILESTONES or (item.kind == "ooxml" and item.get("level") == "node")):
            out.append(_atom_text(item, strict))
            continue
        target = list(item.marks)
        common = 0
        while common < len(stack) and common < len(target) and stack[common] == target[common]:
            common += 1
        for mark in reversed(stack[common:]):
            out.append(_close(mark))
        for mark in target[common:]:
            out.append(_open(mark))
        stack = target
        if isinstance(item, Text):
            out.append(_escape(item.text, strict))
        else:
            out.append(_atom_text(item, strict))
    for mark in reversed(stack):
        out.append(_close(mark))
    text = "".join(out)
    leading = len(text) - len(text.lstrip(" "))
    trailing = len(text) - len(text.rstrip(" "))
    if leading == len(text):
        return "&#32;" * len(text)
    return "&#32;" * leading + text[leading:len(text) - trailing] + "&#32;" * trailing


def write(items, *, block_start: bool = True) -> str:
    """Render ``items`` as one line; falls back to strict escaping if needed."""
    items = normalize(items)
    for strict in (False, True):
        text = _render(items, strict)
        if block_start and text[:1] in {"#", ":"}:
            text = "\\" + text
        try:
            if normalize(parse(text)) == items:
                return text
        except ValueError:
            pass
    raise ValueError(f"Inline content does not round-trip: {items!r}")


# ----------------------------------------------------------------------------
# Parsing


def _find_closing_bracket(text: str, start: int) -> int:
    depth = 0
    index = start
    while index < len(text):
        char = text[index]
        if char == "\\":
            index += 2
            continue
        if char == "<":
            if text.startswith("<ooxml", index):
                end = text.find("</ooxml>", index)
                if end < 0:
                    return -1
                index = end + len("</ooxml>")
                continue
            match = _TAG.match(text, index)
            if match:
                index = match.end()
                continue
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth == 0:
                return index
        index += 1
    return -1


def _find_brace_block(text: str, start: int) -> int:
    """Index of the ``}`` closing an attribute block opened at ``start``."""
    quoted = False
    index = start + 1
    while index < len(text):
        char = text[index]
        if quoted:
            if char == "\\":
                index += 1
            elif char == '"':
                quoted = False
        elif char == '"':
            quoted = True
        elif char == "}":
            return index
        elif char == "{":
            return -1
        index += 1
    return -1


def _valid_token_block(body: str) -> bool:
    try:
        words = list(split_words(body))
    except ValueError:
        return False
    return all(_TOKEN_WORD.match(word) for word in words)


def parse_attr_words(body: str) -> list[tuple[str, str]]:
    result = []
    for word in split_words(body):
        if "=" in word:
            key, value = word.split("=", 1)
            if len(value) >= 2 and value[0] == value[-1] == '"':
                value = re.sub(r"\\(.)", r"\1", value[1:-1])
            result.append((key, value))
        else:
            result.append((word, ""))
    return result


def _unescape_attr(value: str) -> str:
    return re.sub(r"\\(.)", r"\1", value)


def _find_unescaped(text: str, needle: str, start: int) -> int:
    index = start
    while index < len(text):
        if text[index] == "\\":
            index += 2
            continue
        if text.startswith(needle, index):
            return index
        index += 1
    return -1


def parse(text: str) -> list:
    items: list = []
    _parse_into(text, (), items)
    return normalize(items)


def _toggle(stack: list, mark) -> None:
    if mark in stack:
        stack.remove(mark)
    else:
        stack.append(mark)


def _parse_into(text: str, outer: tuple, items: list) -> None:
    stack: list = []
    buffer: list[str] = []

    def marks() -> tuple:
        return sort_marks(tuple(outer) + tuple(stack))

    def flush() -> None:
        if buffer:
            items.append(Text("".join(buffer), marks()))
            buffer.clear()

    index = 0
    while index < len(text):
        char = text[index]
        if char == "\\":
            if text.startswith("\\citep{", index):
                end = text.find("}", index)
                if end < 0:
                    raise ValueError(f"Unclosed citation in: {text!r}")
                flush()
                items.append(Atom("cite", (("keys", text[index + 7:end].replace(" ", "")),), marks()))
                index = end + 1
                continue
            if index + 1 < len(text) and text[index + 1] in _ESCAPABLE:
                buffer.append(text[index + 1])
                index += 2
                continue
            buffer.append(char)
            index += 1
            continue
        if char == "&":
            match = _ENTITY.match(text, index)
            if match:
                code = match.group(1)
                buffer.append(chr(int(code[1:], 16) if code.startswith("x") else int(code)))
                index = match.end()
                continue
        if char == "*":
            run = 0
            while index + run < len(text) and text[index + run] == "*":
                run += 1
            flush()
            remaining = run
            while remaining:
                if stack and stack[-1] == ("i",):
                    stack.pop()
                    remaining -= 1
                elif stack and stack[-1] == ("b",) and remaining >= 2:
                    stack.pop()
                    remaining -= 2
                else:
                    break
            if remaining >= 2 and ("b",) not in stack:
                stack.append(("b",))
                remaining -= 2
            if remaining >= 1 and ("i",) not in stack:
                stack.append(("i",))
                remaining -= 1
            while remaining >= 2 and ("b",) in stack:
                stack.remove(("b",))
                remaining -= 2
            if remaining and ("i",) in stack:
                stack.remove(("i",))
                remaining -= 1
            if remaining:
                raise ValueError(f"Unbalanced emphasis in: {text!r}")
            index += run
            continue
        if text.startswith("==", index) or text.startswith("~~", index):
            flush()
            _toggle(stack, ("hl",) if char == "=" else ("strike",))
            index += 2
            continue
        if text.startswith("{++", index) or text.startswith("{--", index):
            closer = "++}" if text[index + 1] == "+" else "--}"
            end = _find_unescaped(text, closer, index + 3)
            if end < 0:
                raise ValueError(f"Unclosed revision in: {text!r}")
            after = end + 3
            attrs: tuple = ()
            if after < len(text) and text[after] == "{":
                close = _find_brace_block(text, after)
                if close > 0:
                    attrs = tuple(parse_attr_words(text[after + 1:close]))
                    after = close + 1
            flush()
            mark = ("rev", "ins" if closer == "++}" else "del", attrs)
            _parse_into(text[index + 3:end], marks() + (mark,), items)
            index = after
            continue
        if char == "!" and text.startswith("![", index):
            close = _find_closing_bracket(text, index + 1)
            if close > 0 and close + 1 < len(text) and text[close + 1] == "(":
                end = text.find(")", close + 2)
                if end > 0:
                    alt = re.sub(r"\\(.)", r"\1", text[index + 2:close])
                    src = text[close + 2:end].replace("%20", " ")
                    attrs = [("alt", alt), ("src", src)]
                    after = end + 1
                    if after < len(text) and text[after] == "{":
                        brace = _find_brace_block(text, after)
                        if brace > 0:
                            attrs += parse_attr_words(text[after + 1:brace])
                            after = brace + 1
                    flush()
                    items.append(Atom("image", tuple(attrs), marks()))
                    index = after
                    continue
        if char == "[":
            close = _find_closing_bracket(text, index)
            if close > 0 and close + 1 < len(text):
                if text[close + 1] == "{":
                    brace = _find_brace_block(text, close + 1)
                    body = text[close + 2:brace] if brace > 0 else ""
                    if brace > 0 and _valid_token_block(body):
                        flush()
                        tokens = format_tokens(parse_tokens(body))
                        _parse_into(text[index + 1:close], marks() + (("span", tokens),), items)
                        index = brace + 1
                        continue
                elif text[close + 1] == "(":
                    end = text.find(")", close + 2)
                    if end > 0:
                        flush()
                        href = text[close + 2:end].replace("%29", ")").replace("%20", " ")
                        _parse_into(text[index + 1:close], marks() + (("link", href),), items)
                        index = end + 1
                        continue
        if char == "<":
            if text.startswith("<ooxml", index):
                head = _TAG.match(text, index)
                end = text.find("</ooxml>", index)
                if head is None or end < 0:
                    raise ValueError(f"Malformed <ooxml> in: {text!r}")
                flush()
                attrs = tuple((key, _unescape_attr(value)) for key, value in _ATTR.findall(head.group(3)))
                atom = Atom("ooxml", attrs, marks(), text[head.end():end])
                items.append(atom)
                index = end + len("</ooxml>")
                continue
            match = _TAG.match(text, index)
            if match:
                closing, name, raw_attrs, void = match.groups()
                if name in _PAIRED:
                    flush()
                    mark = (name,)
                    if closing:
                        if mark in stack:
                            stack.remove(mark)
                    else:
                        stack.append(mark)
                    index = match.end()
                    continue
                if name in _VOID and not closing:
                    flush()
                    attrs = tuple((key, _unescape_attr(value)) for key, value in _ATTR.findall(raw_attrs))
                    items.append(Atom(name, attrs, marks()))
                    index = match.end()
                    continue
        buffer.append(char)
        index += 1
    flush()


__all__ = ["Atom", "MILESTONES", "Text", "normalize", "parse", "parse_attr_words", "sort_marks", "write"]
