"""Ordinary Markdown reader and writer over SemanticDocument.

Template assembly and the standalone DOCX renderer stay in ``docforge.markdown``.
This adapter projects the block parser into the canonical tree and writes the
supported subset back out.
"""

from __future__ import annotations

import hashlib
import re
import shutil
from pathlib import Path

from ..pipeline.models import Diagnostic
from ..semantic import (
    BlockNode,
    DocumentMetadata,
    InlineNode,
    Provenance,
    Reference,
    ResourceDescriptor,
    SemanticDocument,
    text_node,
)
from .inlines import citation_keys, parse_markdown_inlines, write_latex_inlines, write_markdown_inlines
from .protocol import (
    FormatCapabilities,
    ReadOptions,
    ReadResult,
    SourceContext,
    WriteOptions,
    WritePlan,
    WriteResult,
    ensure_publishable,
    export_losses,
    finding,
    publish_bytes,
    require_valid_document,
)


_CAPABILITIES = FormatCapabilities(
    format_id="markdown",
    node_kinds=("paragraph", "heading", "list", "table", "figure", "math", "code", "quote", "caption"),
    citations=True,
    media=True,
    annotations=False,
    preservation=("REBUILD",),
    notes=("Ordinary Markdown is a lossy semantic format.",),
)


class MarkdownReader:
    format_id = "markdown"

    def probe(self, source: Path) -> FormatCapabilities:
        return _CAPABILITIES

    def read(self, source: Path, *, options: ReadOptions) -> ReadResult:
        from ..markdown import parse_markdown

        source = Path(source)
        blocks = parse_markdown(source, strip_comments=True)
        diagnostics: list[Diagnostic] = []
        counter = _Counter()
        nodes: list[BlockNode] = []
        references: list[Reference] = []
        resources: list[ResourceDescriptor] = []
        index = 0
        while index < len(blocks):
            block = blocks[index]
            if block.kind in {"bullet", "ordered"}:
                ordered = block.kind == "ordered"
                items: list[BlockNode] = []
                while index < len(blocks) and blocks[index].kind == block.kind:
                    current = blocks[index]
                    item_text = re.sub(r"^\d+\.\s+", "", current.text) if ordered else current.text
                    inlines = parse_markdown_inlines(item_text)
                    items.append(BlockNode(counter.block(), "list_item", inlines=inlines, provenance=_prov(source)))
                    _collect_citations(inlines, references, counter)
                    index += 1
                nodes.append(BlockNode(
                    counter.block(),
                    "list",
                    children=tuple(items),
                    attrs=(("ordered", "true" if ordered else "false"), ("start", "1")),
                    provenance=_prov(source),
                ))
                continue
            node, extra_resources = _one_block(block, source, counter, diagnostics)
            nodes.append(node)
            resources.extend(extra_resources)
            _collect_citations(node.inlines, references, counter)
            for child in node.walk():
                _collect_citations(child.inlines, references, counter)
            index += 1
        document = SemanticDocument(
            metadata=DocumentMetadata(),
            blocks=tuple(nodes),
            references=tuple(references),
            resources=tuple(resources),
        )
        payload = source.read_bytes()
        return ReadResult(
            document=document,
            source=SourceContext("markdown", hashlib.sha256(payload).hexdigest(), source.name),
            diagnostics=tuple(diagnostics),
            capabilities=_CAPABILITIES,
            origin_format="markdown",
            origin_name=source.name,
            content_hash=document.semantic_hash(),
        )


class MarkdownWriter:
    format_id = "markdown"

    def plan(self, document: SemanticDocument, *, source, options: WriteOptions) -> WritePlan:
        from .protocol import LossReport

        diagnostics, dropped = export_losses(document, _CAPABILITIES)
        return WritePlan(
            "markdown", "REBUILD", document, source, options,
            LossReport(options.loss_policy, diagnostics, dropped),
        )

    def write(self, plan: WritePlan, output: Path) -> WriteResult:
        from ..pipeline.models import ValidationReport

        output = Path(output)
        require_valid_document(plan.document)
        ensure_publishable(plan.loss_report)
        if output.exists() and not plan.options.force:
            raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
        text = render_markdown(plan.document)
        copy_figure_files(plan.document, plan.origin, output)
        publish_bytes(output, text.encode("utf-8"), force=plan.options.force)
        return WriteResult(output, "REBUILD", plan.loss_report, ValidationReport("pass"))


def render_markdown(document: SemanticDocument) -> str:
    lines: list[str] = []
    for block in document.blocks:
        rendered = _render_block(block)
        if rendered:
            lines.append(rendered)
    return "\n\n".join(lines).rstrip() + "\n"


def render_latex(document: SemanticDocument) -> str:
    body: list[str] = []
    if document.metadata.title:
        body.append("\\title{" + write_latex_inlines((text_node(document.metadata.title),)) + "}")
        body.append("\\maketitle")
    for block in document.blocks:
        rendered = _render_latex_block(block)
        if rendered:
            body.append(rendered)
    rendered_body = "\n\n".join(body)
    preamble = ["\\documentclass{article}"]
    if "textasciitilde" in rendered_body or "textasciicircum" in rendered_body:
        preamble.append("\\usepackage{textcomp}")
    if "\\includegraphics" in rendered_body:
        preamble.append("\\usepackage{graphicx}")
    if "\\href{" in rendered_body:
        preamble.append("\\usepackage{hyperref}")
    return "\n".join([*preamble, "\\begin{document}", "", rendered_body, "\\end{document}", ""])


def copy_figure_files(document: SemanticDocument, origin: Path | None, output: Path) -> None:
    """Copy figure files next to ``output``, keeping the stored relative path."""

    if origin is None:
        return
    source_root = Path(origin).resolve().parent
    output_root = Path(output).resolve().parent
    for block in document.walk():
        if block.kind != "figure":
            continue
        relative = block.attr("path")
        if not relative:
            resource = document.resource(block.attr("resource"))
            relative = resource.path if resource is not None else ""
        if not relative or relative.startswith(("/", "\\")) or ".." in Path(relative).parts:
            continue
        source = (source_root / relative).resolve()
        try:
            source.relative_to(source_root)
        except ValueError:
            continue
        if not source.is_file():
            continue
        destination = (output_root / relative).resolve()
        try:
            destination.relative_to(output_root)
        except ValueError:
            continue
        if destination == source:
            continue
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, destination)


def _render_block(block: BlockNode) -> str:
    if block.kind == "heading":
        level = int(block.attr("level") or "1")
        return "#" * level + " " + write_markdown_inlines(block.inlines)
    if block.kind == "paragraph":
        return write_markdown_inlines(block.inlines)
    if block.kind == "quote":
        return "> " + write_markdown_inlines(block.inlines)
    if block.kind == "code":
        language = block.attr("language")
        return f"```{language}\n{block.attr('code')}\n```"
    if block.kind == "math":
        return "$$\n" + block.attr("tex") + "\n$$"
    if block.kind == "thematic_break":
        return "---"
    if block.kind == "list":
        ordered = block.attr("ordered") == "true"
        rows = []
        for number, item in enumerate(block.children, 1):
            prefix = f"{number}. " if ordered else "- "
            rows.append(prefix + write_markdown_inlines(item.inlines))
        return "\n".join(rows)
    if block.kind == "table":
        rows = []
        for row in block.children:
            cells = [write_markdown_inlines(cell.inlines) or cell.plain() for cell in row.children]
            rows.append("| " + " | ".join(cells) + " |")
        if not rows:
            return ""
        width = rows[0].count("|") - 1
        separator = "| " + " | ".join("---" for _ in range(max(width, 1))) + " |"
        return "\n".join([rows[0], separator, *rows[1:]])
    if block.kind == "figure":
        resource = block.attr("resource")
        path = block.attr("path")
        caption = ""
        for child in block.children:
            if child.kind == "caption":
                caption = write_markdown_inlines(child.inlines)
        return f"![{caption}]({path or resource})"
    if block.kind == "caption":
        return write_markdown_inlines(block.inlines)
    if block.kind == "reference_entry":
        return block.attr("text") or write_markdown_inlines(block.inlines)
    if block.kind == "opaque":
        return ""
    return write_markdown_inlines(block.inlines)


def _render_latex_block(block: BlockNode) -> str:
    if block.kind == "heading":
        level = block.attr("level") or "1"
        command = {"1": "section", "2": "subsection", "3": "subsubsection"}.get(level, "paragraph")
        return "\\" + command + "{" + write_latex_inlines(block.inlines) + "}"
    if block.kind == "paragraph":
        return write_latex_inlines(block.inlines)
    if block.kind == "quote":
        return "\\begin{quote}\n" + write_latex_inlines(block.inlines) + "\n\\end{quote}"
    if block.kind == "code":
        return "\\begin{verbatim}\n" + block.attr("code") + "\n\\end{verbatim}"
    if block.kind == "math":
        return "\\begin{equation}\n" + block.attr("tex") + "\n\\end{equation}"
    if block.kind == "list":
        env = "enumerate" if block.attr("ordered") == "true" else "itemize"
        items = "\n".join("\\item " + write_latex_inlines(item.inlines) for item in block.children)
        return f"\\begin{{{env}}}\n{items}\n\\end{{{env}}}"
    if block.kind == "table":
        body = []
        for row in block.children:
            cells = [write_latex_inlines(cell.inlines) or cell.plain() for cell in row.children]
            body.append(" & ".join(cells) + r" \\")
        cols = "l" * max((len(row.children) for row in block.children), default=1)
        return "\\begin{tabular}{" + cols + "}\n" + "\n".join(body) + "\n\\end{tabular}"
    if block.kind == "figure":
        caption = ""
        for child in block.children:
            if child.kind == "caption":
                caption = write_latex_inlines(child.inlines)
        path = block.attr("path")
        lines = ["\\begin{figure}", f"\\includegraphics{{{path}}}", f"\\caption{{{caption}}}", "\\end{figure}"]
        return "\n".join(lines)
    if block.kind == "thematic_break":
        return r"\noindent\rule{\textwidth}{0.4pt}"
    return write_latex_inlines(block.inlines)


class _Counter:
    def __init__(self) -> None:
        self._value = 0

    def block(self) -> str:
        self._value += 1
        return f"b{self._value}"

    def reference(self) -> str:
        self._value += 1
        return f"r{self._value}"


def _prov(source: Path) -> Provenance:
    return Provenance(source=source.name)


def _collect_citations(nodes: tuple[InlineNode, ...], references: list[Reference], counter: _Counter) -> None:
    for key in citation_keys(nodes):
        if any(item.kind == "citation" and item.key == key for item in references):
            continue
        references.append(Reference(counter.reference(), "citation", key))


def _one_block(block, source: Path, counter: _Counter, diagnostics: list[Diagnostic]) -> tuple[BlockNode, list[ResourceDescriptor]]:
    provenance = _prov(source)
    if block.kind == "heading":
        return BlockNode(
            counter.block(), "heading",
            inlines=parse_markdown_inlines(block.text),
            attrs=(("level", str(block.level)),),
            provenance=provenance,
        ), []
    if block.kind == "paragraph":
        return BlockNode(counter.block(), "paragraph", inlines=parse_markdown_inlines(block.text), provenance=provenance), []
    if block.kind == "quote":
        return BlockNode(counter.block(), "quote", inlines=parse_markdown_inlines(block.text), provenance=provenance), []
    if block.kind == "code":
        return BlockNode(
            counter.block(), "code",
            attrs=(("language", block.language), ("code", block.text)),
            provenance=provenance,
        ), []
    if block.kind == "equation":
        return BlockNode(counter.block(), "math", attrs=(("tex", block.text), ("display", "true")), provenance=provenance), []
    if block.kind == "separator":
        return BlockNode(counter.block(), "thematic_break", provenance=provenance), []
    if block.kind == "table_caption":
        return BlockNode(counter.block(), "caption", inlines=parse_markdown_inlines(block.text), attrs=(("kind", "table"),), provenance=provenance), []
    if block.kind == "reference":
        return BlockNode(counter.block(), "reference_entry", attrs=(("text", block.text),), provenance=provenance), []
    if block.kind == "table":
        rows = []
        for row in block.rows:
            cells = tuple(
                BlockNode(counter.block(), "table_cell", inlines=parse_markdown_inlines(cell), provenance=provenance)
                for cell in row
            )
            rows.append(BlockNode(counter.block(), "table_row", children=cells, provenance=provenance))
        return BlockNode(counter.block(), "table", children=tuple(rows), provenance=provenance), []
    if block.kind == "image":
        resource_id = "res" + counter.block()
        path = Path(block.path) if block.path else None
        relative = ""
        sha = ""
        media = "application/octet-stream"
        if path is not None and path.is_file():
            try:
                relative = path.resolve().relative_to(source.parent.resolve()).as_posix()
            except ValueError:
                relative = path.name
                diagnostics.append(finding(
                    "warning", "resource.path_reduced", "image path was reduced to its file name",
                    category="loss", name=path.name,
                ))
            sha = hashlib.sha256(path.read_bytes()).hexdigest()
            suffix = path.suffix.casefold()
            media = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif"}.get(suffix, media)
        else:
            diagnostics.append(finding("error", "resource.missing", "image file is missing", name=Path(block.path).name if block.path else ""))
        caption = BlockNode(counter.block(), "caption", inlines=parse_markdown_inlines(block.text), provenance=provenance)
        figure = BlockNode(
            counter.block(), "figure",
            children=(caption,),
            attrs=(("resource", resource_id), ("path", relative)),
            provenance=provenance,
        )
        return figure, [ResourceDescriptor(resource_id, media, sha, relative)]
    diagnostics.append(finding(
        "warning", "markdown.unsupported_block", f"unsupported markdown block kind: {block.kind}",
        category="loss",
    ))
    return BlockNode(counter.block(), "opaque", attrs=(("origin", block.kind),), provenance=provenance), []


__all__ = ["MarkdownReader", "MarkdownWriter", "copy_figure_files", "render_latex", "render_markdown"]
