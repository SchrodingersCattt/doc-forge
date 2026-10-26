"""Supported-subset LaTeX reader and writer.

This is not a TeX engine. Authoring macros, preamble layout, and custom
environments are reported and are not reconstructed. The native TeX → DOCX
converter remains the renderer for ``tex2docx``.
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

from ..pipeline.models import Diagnostic
from ..semantic import BlockNode, DocumentMetadata, Provenance, Reference, SemanticDocument
from .inlines import citation_keys, parse_latex_inlines
from .markdown import render_latex
from .protocol import (
    FormatCapabilities,
    LossReport,
    ReadOptions,
    ReadResult,
    SourceContext,
    WriteOptions,
    WritePlan,
    WriteResult,
    finding,
)


_CAPABILITIES = FormatCapabilities(
    format_id="latex",
    node_kinds=("paragraph", "heading", "list", "table", "figure", "math", "code", "quote"),
    citations=True,
    media=True,
    annotations=False,
    preservation=("REBUILD",),
    notes=("Custom macros and preamble layout are outside the supported subset.",),
)

_BLOCK_ENVS = {
    "itemize", "enumerate", "quote", "quotation", "equation", "equation*",
    "align", "align*", "figure", "table", "tabular", "verbatim", "lstlisting",
}
_HEADINGS = {
    "section": "1",
    "subsection": "2",
    "subsubsection": "3",
    "paragraph": "4",
}


class LatexReader:
    format_id = "latex"

    def probe(self, source: Path) -> FormatCapabilities:
        return _CAPABILITIES

    def read(self, source: Path, *, options: ReadOptions) -> ReadResult:
        source = Path(source)
        text = source.read_text(encoding="utf-8-sig")
        diagnostics: list[Diagnostic] = []
        title = _command_arg(text, "title")
        body = _document_body(text)
        counter = _Counter()
        blocks, references = _parse_body(body, source.name, counter, diagnostics)
        document = SemanticDocument(
            metadata=DocumentMetadata(title=_plain(title)),
            blocks=tuple(blocks),
            references=tuple(references),
        )
        return ReadResult(
            document=document,
            source=SourceContext("latex", hashlib.sha256(text.encode("utf-8")).hexdigest(), source.name),
            diagnostics=tuple(diagnostics),
            capabilities=_CAPABILITIES,
            origin_format="latex",
            origin_name=source.name,
            content_hash=document.semantic_hash(),
        )


class LatexWriter:
    format_id = "latex"

    def plan(self, document: SemanticDocument, *, source, options: WriteOptions) -> WritePlan:
        dropped = tuple(block.id for block in document.walk() if block.kind == "opaque")
        diagnostics = tuple(
            finding("warning", "latex.opaque_omitted", "opaque block omitted from LaTeX export", block=block.id)
            for block in document.walk() if block.kind == "opaque"
        )
        return WritePlan("latex", "REBUILD", document, source, options, LossReport(options.loss_policy, diagnostics, dropped))

    def write(self, plan: WritePlan, output: Path) -> WriteResult:
        from ..pipeline.models import ValidationReport

        output = Path(output)
        if output.exists() and not plan.options.force:
            raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(render_latex(plan.document), encoding="utf-8")
        errors = plan.document.validate()
        status = "pass" if not errors else "fail"
        return WriteResult(
            output, "REBUILD", plan.loss_report,
            ValidationReport(status, tuple(finding("error", "semantic.invalid", item) for item in errors)),
        )


class _Counter:
    def __init__(self) -> None:
        self.value = 0

    def next(self, prefix: str) -> str:
        self.value += 1
        return f"{prefix}{self.value}"


def _plain(nodes_text: str) -> str:
    return "".join(node.plain() for node in parse_latex_inlines(nodes_text)).strip()


def _strip_comments(text: str) -> str:
    return re.sub(r"(?<!\\)%.*", "", text)


def _document_body(text: str) -> str:
    text = _strip_comments(text)
    match = re.search(r"\\begin\{document\}", text)
    if not match:
        return text
    end = re.search(r"\\end\{document\}", text)
    return text[match.end(): end.start() if end else len(text)]


def _command_arg(text: str, command: str) -> str:
    cleaned = _strip_comments(text)
    match = re.search(rf"\\{command}\*?\{{", cleaned)
    if not match:
        return ""
    return _balanced(cleaned, match.end() - 1)[0]


def _balanced(text: str, brace_index: int) -> tuple[str, int]:
    if brace_index >= len(text) or text[brace_index] != "{":
        return "", brace_index
    depth = 0
    for cursor in range(brace_index, len(text)):
        char = text[cursor]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[brace_index + 1:cursor], cursor + 1
    return text[brace_index + 1:], len(text)


def _parse_body(text: str, filename: str, counter: _Counter, diagnostics: list[Diagnostic]) -> tuple[list[BlockNode], list[Reference]]:
    blocks: list[BlockNode] = []
    references: list[Reference] = []
    paragraph: list[str] = []

    def flush() -> None:
        raw = "".join(paragraph).strip()
        paragraph.clear()
        if not raw:
            return
        unknown: list[str] = []
        inlines = parse_latex_inlines(raw, diagnostics=unknown)
        for command in sorted(set(unknown)):
            diagnostics.append(finding(
                "warning", "latex.unsupported_command",
                f"unsupported command \\{command} kept as text",
                command=command,
            ))
        block = BlockNode(counter.next("b"), "paragraph", inlines=inlines, provenance=Provenance(filename))
        blocks.append(block)
        _refs(inlines, references, counter)

    index = 0
    while index < len(text):
        if text[index].isspace():
            if text[index] == "\n" and index + 1 < len(text) and text[index + 1] == "\n":
                flush()
            else:
                paragraph.append(text[index])
            index += 1
            continue
        if text.startswith("$$", index):
            flush()
            end = text.find("$$", index + 2)
            if end == -1:
                diagnostics.append(finding("error", "latex.unclosed_math", "unclosed display math"))
                break
            blocks.append(BlockNode(
                counter.next("b"), "math",
                attrs=(("tex", text[index + 2:end].strip()), ("display", "true")),
                provenance=Provenance(filename),
            ))
            index = end + 2
            continue
        if text.startswith("\\[", index):
            flush()
            end = text.find("\\]", index + 2)
            if end == -1:
                diagnostics.append(finding("error", "latex.unclosed_math", "unclosed display math"))
                break
            blocks.append(BlockNode(
                counter.next("b"), "math",
                attrs=(("tex", text[index + 2:end].strip()), ("display", "true")),
                provenance=Provenance(filename),
            ))
            index = end + 2
            continue
        if text[index] == "\\":
            command = _command_name(text, index)
            if command in _HEADINGS:
                flush()
                arg, index = _argument_after(text, index + 1 + len(command))
                unknown: list[str] = []
                inlines = parse_latex_inlines(arg, diagnostics=unknown)
                blocks.append(BlockNode(
                    counter.next("b"), "heading", inlines=inlines,
                    attrs=(("level", _HEADINGS[command]),),
                    provenance=Provenance(filename),
                ))
                _refs(inlines, references, counter)
                continue
            if command == "begin":
                flush()
                env, after = _argument_after(text, index + len("\\begin"))
                body, end = _environment_body(text, after, env)
                index = end
                blocks.extend(_environment(env, body, filename, counter, diagnostics, references))
                continue
            if command in {"end", "maketitle", "tableofcontents"}:
                flush()
                if command == "end":
                    _, index = _argument_after(text, index + len("\\end"))
                else:
                    index += 1 + len(command)
                continue
            if command in {"item"}:
                paragraph.append(text[index])
                index += 1 + len(command)
                continue
        paragraph.append(text[index])
        index += 1
    flush()
    return blocks, references


def _command_name(text: str, index: int) -> str:
    match = re.match(r"\\([A-Za-z]+)\*?", text[index:])
    return match.group(1) if match else ""


def _argument_after(text: str, index: int) -> tuple[str, int]:
    while index < len(text) and (text[index].isspace() or text[index] == "*"):
        index += 1
    if index < len(text) and text[index] == "{":
        return _balanced(text, index)
    return "", index


def _environment_body(text: str, index: int, env: str) -> tuple[str, int]:
    pattern = re.compile(rf"\\(begin|end)\{{{re.escape(env)}\}}")
    depth = 1
    cursor = index
    while depth and cursor < len(text):
        match = pattern.search(text, cursor)
        if match is None:
            return text[index:], len(text)
        if match.group(1) == "begin":
            depth += 1
            cursor = match.end()
        else:
            depth -= 1
            if depth == 0:
                return text[index:match.start()], match.end()
            cursor = match.end()
    return text[index:], len(text)


def _environment(env: str, body: str, filename: str, counter: _Counter, diagnostics: list[Diagnostic], references: list[Reference]) -> list[BlockNode]:
    provenance = Provenance(filename)
    name = env.rstrip("*")
    if name in {"itemize", "enumerate"}:
        items = []
        for raw in re.split(r"\\item\b", body)[1:]:
            unknown: list[str] = []
            inlines = parse_latex_inlines(raw.strip(), diagnostics=unknown)
            items.append(BlockNode(counter.next("b"), "list_item", inlines=inlines, provenance=provenance))
            _refs(inlines, references, counter)
        return [BlockNode(
            counter.next("b"), "list", children=tuple(items),
            attrs=(("ordered", "true" if name == "enumerate" else "false"), ("start", "1")),
            provenance=provenance,
        )]
    if name in {"quote", "quotation"}:
        unknown = []
        inlines = parse_latex_inlines(" ".join(body.split()), diagnostics=unknown)
        return [BlockNode(counter.next("b"), "quote", inlines=inlines, provenance=provenance)]
    if name in {"equation", "align"}:
        return [BlockNode(counter.next("b"), "math", attrs=(("tex", body.strip()), ("display", "true")), provenance=provenance)]
    if name == "verbatim" or name == "lstlisting":
        return [BlockNode(counter.next("b"), "code", attrs=(("language", ""), ("code", body.strip("\n"))), provenance=provenance)]
    if name == "tabular":
        return [_tabular(body, provenance, counter)]
    if name == "figure":
        return [_figure(body, provenance, counter, diagnostics)]
    if name == "table":
        opened = re.search(r"\\begin\{tabular\}\{[^}]*\}", body)
        inner = _environment_body(body, opened.end(), "tabular")[0] if opened else body
        tabular = _tabular(inner, provenance, counter)
        caption = _command_arg(body, "caption")
        children = [tabular]
        if caption:
            children.append(BlockNode(counter.next("b"), "caption", inlines=parse_latex_inlines(caption), attrs=(("kind", "table"),), provenance=provenance))
        return children
    diagnostics.append(finding("warning", "latex.unsupported_environment", f"unsupported environment {env}", environment=env))
    return [BlockNode(counter.next("b"), "opaque", attrs=(("environment", env), ("tex", body.strip())), provenance=provenance)]


def _tabular(body: str, provenance: Provenance, counter: _Counter) -> BlockNode:
    rows = []
    for raw in body.split("\\\\"):
        if not raw.strip() or raw.strip().startswith("\\hline"):
            raw = raw.replace("\\hline", "")
        if not raw.strip():
            continue
        cells = []
        for cell in raw.split("&"):
            cells.append(BlockNode(counter.next("b"), "table_cell", inlines=parse_latex_inlines(cell.strip()), provenance=provenance))
        rows.append(BlockNode(counter.next("b"), "table_row", children=tuple(cells), provenance=provenance))
    return BlockNode(counter.next("b"), "table", children=tuple(rows), provenance=provenance)


def _figure(body: str, provenance: Provenance, counter: _Counter, diagnostics: list[Diagnostic]) -> BlockNode:
    graphic = re.search(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}", body)
    path = graphic.group(1).strip() if graphic else ""
    if not graphic:
        diagnostics.append(finding("warning", "latex.figure_missing_graphic", "figure has no includegraphics"))
    caption = _command_arg(body, "caption")
    children = []
    if caption:
        children.append(BlockNode(counter.next("b"), "caption", inlines=parse_latex_inlines(caption), provenance=provenance))
    return BlockNode(
        counter.next("b"), "figure", children=tuple(children),
        attrs=(("path", path), ("resource", "")),
        provenance=provenance,
    )


def _refs(nodes, references: list[Reference], counter: _Counter) -> None:
    for key in citation_keys(nodes):
        if any(item.kind == "citation" and item.key == key for item in references):
            continue
        references.append(Reference(counter.next("r"), "citation", key))


__all__ = ["LatexReader", "LatexWriter"]
