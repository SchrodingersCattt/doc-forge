"""Small inline codec for ordinary Markdown and the supported TeX subset.

The reversible docxmd codec remains the lossless bundle parser. This module
covers the semantic intersection used by Markdown ↔ LaTeX conversion and does
not import Word.
"""

from __future__ import annotations

import re

from ..semantic import InlineNode, text_node


_MD = re.compile(
    r"\\cite[p]?\{(?P<cite>[^}]+)\}"
    r"|`(?P<code>[^`]+)`"
    r"|\$\$?(?P<math>[^$]+)\$\$?"
    r"|\*\*\*(?P<both>.+?)\*\*\*"
    r"|\*\*(?P<strong>.+?)\*\*"
    r"|\*(?P<emph>.+?)\*"
    r"|\[(?P<link>[^\]]+)\]\((?P<url>[^)]+)\)"
    r"|<sub>(?P<sub>.*?)</sub>"
    r"|<sup>(?P<sup>.*?)</sup>"
    r"|<u>(?P<u>.*?)</u>",
    re.DOTALL,
)


def parse_markdown_inlines(text: str) -> tuple[InlineNode, ...]:
    nodes: list[InlineNode] = []
    cursor = 0
    for match in _MD.finditer(text):
        if match.start() > cursor:
            nodes.append(text_node(text[cursor:match.start()]))
        if match.group("cite") is not None:
            keys = tuple(part.strip() for part in match.group("cite").split(",") if part.strip())
            nodes.append(InlineNode("citation", attrs=(("keys", ",".join(keys)),)))
        elif match.group("code") is not None:
            nodes.append(InlineNode("code", text=match.group("code")))
        elif match.group("math") is not None:
            nodes.append(InlineNode("math", text=match.group("math").strip(), attrs=(("display", "false"),)))
        elif match.group("both") is not None:
            nodes.append(InlineNode("strong", children=(InlineNode("emph", children=parse_markdown_inlines(match.group("both"))),)))
        elif match.group("strong") is not None:
            nodes.append(InlineNode("strong", children=parse_markdown_inlines(match.group("strong"))))
        elif match.group("emph") is not None:
            nodes.append(InlineNode("emph", children=parse_markdown_inlines(match.group("emph"))))
        elif match.group("link") is not None:
            nodes.append(InlineNode("link", attrs=(("href", match.group("url")),), children=parse_markdown_inlines(match.group("link"))))
        elif match.group("sub") is not None:
            nodes.append(InlineNode("subscript", children=parse_markdown_inlines(match.group("sub"))))
        elif match.group("sup") is not None:
            nodes.append(InlineNode("superscript", children=parse_markdown_inlines(match.group("sup"))))
        elif match.group("u") is not None:
            nodes.append(InlineNode("underline", children=parse_markdown_inlines(match.group("u"))))
        cursor = match.end()
    if cursor < len(text):
        nodes.append(text_node(text[cursor:]))
    return tuple(node for node in nodes if node.kind != "text" or node.text)


def write_markdown_inlines(nodes: tuple[InlineNode, ...]) -> str:
    return "".join(_write_markdown(node) for node in nodes)


def _write_markdown(node: InlineNode) -> str:
    inner = write_markdown_inlines(node.children)
    if node.kind == "text":
        return node.text
    if node.kind == "strong":
        return f"**{inner or node.text}**"
    if node.kind == "emph":
        return f"*{inner or node.text}*"
    if node.kind == "code":
        return f"`{node.text}`"
    if node.kind == "math":
        return f"${node.text}$"
    if node.kind == "link":
        return f"[{inner or node.text}]({node.attr('href')})"
    if node.kind == "subscript":
        return f"<sub>{inner or node.text}</sub>"
    if node.kind == "superscript":
        return f"<sup>{inner or node.text}</sup>"
    if node.kind == "underline":
        return f"<u>{inner or node.text}</u>"
    if node.kind == "citation":
        return "\\cite{" + node.attr("keys") + "}"
    if node.kind == "break":
        return "\n"
    if node.kind == "strike":
        return f"~~{inner or node.text}~~"
    return inner or node.text


def write_latex_inlines(nodes: tuple[InlineNode, ...]) -> str:
    return "".join(_write_latex(node) for node in nodes)


def _write_latex(node: InlineNode) -> str:
    inner = write_latex_inlines(node.children)
    if node.kind == "text":
        return _escape_latex(node.text)
    if node.kind == "strong":
        return "\\textbf{" + (inner or _escape_latex(node.text)) + "}"
    if node.kind == "emph":
        return "\\emph{" + (inner or _escape_latex(node.text)) + "}"
    if node.kind == "code":
        return "\\texttt{" + _escape_latex(node.text) + "}"
    if node.kind == "math":
        body = node.text
        return f"$${body}$$" if node.attr("display") == "true" else f"${body}$"
    if node.kind == "link":
        return "\\href{" + node.attr("href") + "}{" + (inner or _escape_latex(node.text)) + "}"
    if node.kind == "subscript":
        return "\\textsubscript{" + (inner or _escape_latex(node.text)) + "}"
    if node.kind == "superscript":
        return "\\textsuperscript{" + (inner or _escape_latex(node.text)) + "}"
    if node.kind == "underline":
        return "\\underline{" + (inner or _escape_latex(node.text)) + "}"
    if node.kind == "citation":
        return "\\cite{" + node.attr("keys") + "}"
    if node.kind == "break":
        return "\n"
    return inner or _escape_latex(node.text)


def _escape_latex(text: str) -> str:
    return (
        text.replace("\\", "\\textbackslash{}")
        .replace("&", "\\&")
        .replace("%", "\\%")
        .replace("#", "\\#")
        .replace("_", "\\_")
    )


_TEX_CMD = re.compile(r"\\([A-Za-z]+)\*?")


def parse_latex_inlines(text: str, *, diagnostics: list | None = None) -> tuple[InlineNode, ...]:
    """Parse a supported TeX inline subset. Unknown commands keep their argument text."""

    nodes: list[InlineNode] = []
    index = 0
    while index < len(text):
        if text[index] == "\\":
            match = _TEX_CMD.match(text, index)
            if match is None:
                nodes.append(text_node(text[index]))
                index += 1
                continue
            command = match.group(1)
            index = match.end()
            argument, index = _braced(text, index)
            if command in {"textbf", "mathbf"}:
                nodes.append(InlineNode("strong", children=parse_latex_inlines(argument, diagnostics=diagnostics)))
            elif command in {"emph", "textit", "mathit"}:
                nodes.append(InlineNode("emph", children=parse_latex_inlines(argument, diagnostics=diagnostics)))
            elif command == "texttt":
                nodes.append(InlineNode("code", text=argument))
            elif command in {"cite", "citep", "mainref"}:
                keys = ",".join(part.strip() for part in argument.split(",") if part.strip())
                nodes.append(InlineNode("citation", attrs=(("keys", keys),)))
            elif command == "href":
                label, index = _braced(text, index)
                nodes.append(InlineNode("link", attrs=(("href", argument),), children=parse_latex_inlines(label, diagnostics=diagnostics)))
            elif command in {"textsubscript"}:
                nodes.append(InlineNode("subscript", children=parse_latex_inlines(argument, diagnostics=diagnostics)))
            elif command in {"textsuperscript"}:
                nodes.append(InlineNode("superscript", children=parse_latex_inlines(argument, diagnostics=diagnostics)))
            elif command == "underline":
                nodes.append(InlineNode("underline", children=parse_latex_inlines(argument, diagnostics=diagnostics)))
            elif command in {"label", "ref", "autoref", "eqref"}:
                if command == "label":
                    nodes.append(InlineNode("span", text=argument, attrs=(("label", argument),)))
                else:
                    nodes.append(InlineNode("link", text=argument, attrs=(("href", "#" + argument),)))
            else:
                if diagnostics is not None:
                    diagnostics.append(command)
                nodes.extend(parse_latex_inlines(argument, diagnostics=diagnostics))
            continue
        if text[index] == "$":
            end = text.find("$", index + 1)
            if end == -1:
                nodes.append(text_node(text[index:]))
                break
            nodes.append(InlineNode("math", text=text[index + 1:end].strip(), attrs=(("display", "false"),)))
            index = end + 1
            continue
        next_special = _next_special(text, index + 1)
        nodes.append(text_node(text[index:next_special]))
        index = next_special
    return tuple(node for node in nodes if node.kind != "text" or node.text)


def _next_special(text: str, start: int) -> int:
    positions = [len(text)]
    for mark in ("\\", "$"):
        found = text.find(mark, start)
        if found != -1:
            positions.append(found)
    return min(positions)


def _braced(text: str, index: int) -> tuple[str, int]:
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != "{":
        return "", index
    depth = 0
    start = index + 1
    for cursor in range(index, len(text)):
        if text[cursor] == "{":
            depth += 1
        elif text[cursor] == "}":
            depth -= 1
            if depth == 0:
                return text[start:cursor], cursor + 1
    return text[start:], len(text)


def citation_keys(nodes: tuple[InlineNode, ...]) -> tuple[str, ...]:
    found: list[str] = []
    for node in nodes:
        if node.kind == "citation":
            found.extend(part for part in node.attr("keys").split(",") if part)
        found.extend(citation_keys(node.children))
    return tuple(found)


__all__ = [
    "citation_keys", "parse_latex_inlines", "parse_markdown_inlines",
    "write_latex_inlines", "write_markdown_inlines",
]
