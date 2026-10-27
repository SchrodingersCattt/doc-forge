"""Ordinary Markdown reader and writer over SemanticDocument.

Template assembly and the standalone DOCX renderer stay in ``docforge.markdown``.
This adapter projects the block parser into the canonical tree and writes the
supported subset back out.
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

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
    ConversionError,
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
        output = Path(output)
        require_valid_document(plan.document)
        ensure_publishable(plan.loss_report)
        if output.exists() and not plan.options.force:
            raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
        text = render_markdown(plan.document, origin=plan.origin)
        report = publish_text_export(
            output, text, plan.document, plan.origin, force=plan.options.force, kind="markdown",
        )
        return WriteResult(output, "REBUILD", plan.loss_report, report)


def render_markdown(document: SemanticDocument, *, origin: Path | None = None) -> str:
    paths = _figure_output_paths(document, origin)
    lines: list[str] = []
    for block in document.blocks:
        rendered = _render_block(block, paths)
        if rendered:
            lines.append(rendered)
    return "\n\n".join(lines).rstrip() + "\n"


_OUTER_DISPLAY = re.compile(
    r"\\begin\{(?P<env>equation\*?|align\*?|gather\*?|multline\*?|flalign\*?|eqnarray\*?|displaymath)\}"
    r"(?P<body>.*)\\end\{(?P=env)\}\s*\Z",
    re.DOTALL,
)
_AMSMATH_MARKERS = (
    r"\begin{align",
    r"\begin{gather",
    r"\begin{multline",
    r"\begin{flalign",
    r"\begin{eqnarray",
    r"\begin{aligned",
    r"\begin{cases",
    r"\begin{split",
)


def display_math_tex(tex: str) -> str:
    """Return compilable display math.

    An outer ``align`` (or other display environment) is kept as itself.
    ``equation`` is not wrapped around another display environment.
    """

    stripped = tex.strip()
    match = _OUTER_DISPLAY.match(stripped)
    if match is None:
        return "\\begin{equation}\n" + stripped + "\n\\end{equation}"
    if match.group("env").startswith("equation"):
        inner = _OUTER_DISPLAY.match(match.group("body").strip())
        if inner is not None and not inner.group("env").startswith("equation"):
            return match.group("body").strip()
    return stripped


def semantic_display_source(env: str, body: str) -> str:
    """Tex stored for one display environment.

    Plain equations stay bare. ``align`` keeps its environment so a later write
    does not nest it inside ``equation``.
    """

    inner = body.strip()
    nested = _OUTER_DISPLAY.match(inner)
    if nested is not None and not nested.group("env").startswith("equation"):
        return inner
    if env.rstrip("*") == "equation":
        return inner
    return f"\\begin{{{env}}}\n{inner}\n\\end{{{env}}}"


def render_latex(document: SemanticDocument, *, origin: Path | None = None) -> str:
    paths = _figure_output_paths(document, origin)
    body: list[str] = []
    if document.metadata.title:
        body.append("\\title{" + write_latex_inlines((text_node(document.metadata.title),)) + "}")
        body.append("\\maketitle")
    for block in document.blocks:
        rendered = _render_latex_block(block, paths)
        if rendered:
            body.append(rendered)
    rendered_body = "\n\n".join(body)
    preamble = ["\\documentclass{article}"]
    if any(token in rendered_body for token in ("textasciitilde", "textasciicircum", "textbackslash")):
        preamble.append("\\usepackage{textcomp}")
    if "\\includegraphics" in rendered_body:
        preamble.append("\\usepackage{graphicx}")
    if any(token in rendered_body for token in _AMSMATH_MARKERS):
        preamble.append("\\usepackage{amsmath}")
    if "\\href{" in rendered_body:
        preamble.append("\\usepackage{hyperref}")
    return "\n".join([*preamble, "\\begin{document}", "", rendered_body, "\\end{document}", ""])


@dataclass(frozen=True)
class _MediaPlacement:
    block_id: str
    stored: str
    output_rel: str
    source: Path | None


def plan_media(document: SemanticDocument, origin: Path | None) -> tuple[_MediaPlacement, ...]:
    """Map each figure to a portable output path and the file that should be copied."""

    used: set[str] = set()
    placements: list[_MediaPlacement] = []
    for block in document.walk():
        if block.kind != "figure":
            continue
        stored = _stored_figure_path(document, block)
        portable = _dedupe_rel(_portable_media_rel(stored), used)
        placements.append(_MediaPlacement(block.id, stored, portable, _media_source(origin, stored)))
    return tuple(placements)


def publish_text_export(
    output: Path,
    text: str,
    document: SemanticDocument,
    origin: Path | None,
    *,
    force: bool,
    kind: str,
):
    """Stage text and figures, audit the candidate, and publish only on success."""

    output = Path(output)
    placements = plan_media(document, origin)
    files: dict[str, bytes] = {output.name: text.encode("utf-8")}
    for item in placements:
        if item.source is None or not item.output_rel:
            continue
        files[item.output_rel] = item.source.read_bytes()
    report = _audit_text_export(files, output.name, kind=kind, placements=placements, origin=origin)
    if not report.ok:
        raise ConversionError("export audit failed", code="AUDIT_FAILED", diagnostics=report.diagnostics)
    _publish_files(output, files, force=force)
    return report


def copy_figure_files(document: SemanticDocument, origin: Path | None, output: Path) -> None:
    """Copy figure files next to ``output`` using portable relative paths."""

    if origin is None:
        return
    output_root = Path(output).resolve().parent
    for item in plan_media(document, origin):
        if item.source is None or not item.output_rel:
            continue
        destination = (output_root / item.output_rel).resolve()
        try:
            destination.relative_to(output_root)
        except ValueError:
            continue
        if destination == item.source.resolve():
            continue
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(item.source, destination)


def _figure_output_paths(document: SemanticDocument, origin: Path | None) -> dict[str, str]:
    return {item.block_id: item.output_rel for item in plan_media(document, origin)}


def _stored_figure_path(document: SemanticDocument, block: BlockNode) -> str:
    path = (block.attr("path") or "").replace("\\", "/").strip()
    if path:
        return path
    resource = document.resource(block.attr("resource"))
    if resource is None:
        return ""
    return (resource.path or "").replace("\\", "/").strip()


def _portable_media_rel(stored: str) -> str:
    if not stored or "://" in stored:
        return ""
    raw = stored.strip()
    path = PurePosixPath(raw)
    if raw.startswith("/") or (len(raw) > 1 and raw[1] == ":"):
        return path.name
    parts: list[str] = []
    escaped = False
    for piece in path.parts:
        if piece in {"", "."}:
            continue
        if piece == "..":
            if parts:
                parts.pop()
            else:
                escaped = True
            continue
        parts.append(piece)
    if escaped or not parts:
        return path.name
    return "/".join(parts)


def _dedupe_rel(rel: str, used: set[str]) -> str:
    if not rel:
        return ""
    if rel not in used:
        used.add(rel)
        return rel
    path = PurePosixPath(rel)
    parent = "" if str(path.parent) in {"", "."} else path.parent.as_posix()
    suffix = path.suffix
    stem = path.stem
    number = 2
    while True:
        name = f"{stem}-{number}{suffix}"
        candidate = f"{parent}/{name}" if parent else name
        if candidate not in used:
            used.add(candidate)
            return candidate
        number += 1


def _media_source(origin: Path | None, stored: str) -> Path | None:
    if origin is None or not stored or "://" in stored:
        return None
    raw = stored.strip()
    if raw.startswith("/") or (len(raw) > 1 and raw[1] == ":"):
        candidate = Path(raw)
    else:
        candidate = (Path(origin).parent / raw).resolve()
    try:
        return candidate if candidate.is_file() else None
    except OSError:
        return None


def _latex_filename(path: str) -> str:
    if any(char in path for char in " %#&{}$~^\\"):
        return "\\detokenize{" + path.replace("\\", "/") + "}"
    return path


def _audit_text_export(files: dict[str, bytes], output_name: str, *, kind: str, placements: tuple[_MediaPlacement, ...], origin: Path | None):
    from ..pipeline.models import ValidationReport

    diagnostics = []
    try:
        text = files.get(output_name, b"").decode("utf-8")
    except UnicodeError:
        return ValidationReport("fail", (finding("error", "export.encoding", "export is not valid UTF-8"),))
    if kind == "latex":
        if "\\documentclass" not in text or "\\begin{document}" not in text or "\\end{document}" not in text:
            diagnostics.append(finding("error", "latex.structure", "generated LaTeX is missing a document wrapper"))
        else:
            preamble, _, body = text.partition("\\begin{document}")
            if "\\includegraphics" in body and "\\usepackage{graphicx}" not in preamble:
                diagnostics.append(finding("error", "latex.package", "graphicx is required for includegraphics"))
            if "\\href{" in body and "\\usepackage{hyperref}" not in preamble:
                diagnostics.append(finding("error", "latex.package", "hyperref is required for href"))
            if any(token in body for token in ("\\textasciitilde", "\\textasciicircum", "\\textbackslash")) and "\\usepackage{textcomp}" not in preamble:
                diagnostics.append(finding("error", "latex.package", "textcomp is required for escaped text symbols"))
    if origin is not None:
        for item in placements:
            if not item.stored or "://" in item.stored:
                continue
            if item.source is None or not item.output_rel or item.output_rel not in files:
                diagnostics.append(finding(
                    "error", "export.missing_media",
                    "figure file was not staged with the export",
                    name=PurePosixPath(item.output_rel or item.stored).name,
                ))
    status = "fail" if any(item.severity == "error" for item in diagnostics) else "pass"
    return ValidationReport(status, tuple(diagnostics), {"export": kind})


def _publish_files(output: Path, files: dict[str, bytes], *, force: bool) -> None:
    output = Path(output)
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
    root = output.parent
    planned: list[tuple[Path, bytes]] = []
    for rel, payload in files.items():
        rel_path = PurePosixPath(rel)
        if rel_path.is_absolute() or ".." in rel_path.parts:
            raise ConversionError("export path escapes the output directory", code="PATH_ESCAPE")
        destination = root.joinpath(*rel_path.parts)
        if destination.exists() and not force:
            raise FileExistsError(f"Output exists; pass --force to overwrite: {destination}")
        planned.append((destination, payload))
    root.mkdir(parents=True, exist_ok=True)
    staged: list[tuple[Path, Path]] = []
    try:
        for destination, payload in planned:
            destination.parent.mkdir(parents=True, exist_ok=True)
            descriptor, temporary_name = tempfile.mkstemp(
                prefix=f".{destination.stem}-", suffix=destination.suffix, dir=destination.parent,
            )
            os.close(descriptor)
            temporary = Path(temporary_name)
            staged.append((temporary, destination))
            temporary.write_bytes(payload)
        for temporary, destination in staged:
            os.replace(temporary, destination)
    finally:
        for temporary, _destination in staged:
            temporary.unlink(missing_ok=True)


def _render_block(block: BlockNode, paths: dict[str, str]) -> str:
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
        path = paths.get(block.id, block.attr("path"))
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


def _render_latex_block(block: BlockNode, paths: dict[str, str]) -> str:
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
        return display_math_tex(block.attr("tex"))
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
        path = paths.get(block.id) or block.attr("path")
        lines = ["\\begin{figure}"]
        if path:
            lines.append("\\includegraphics{" + _latex_filename(path) + "}")
        lines.extend((f"\\caption{{{caption}}}", "\\end{figure}"))
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
        portable = ""
        sha = ""
        media = "application/octet-stream"
        if path is not None and path.is_file():
            relative = os.path.relpath(path.resolve(), source.parent.resolve()).replace("\\", "/")
            portable = _portable_media_rel(relative) or path.name
            if relative.startswith(("/", "\\")) or (len(relative) > 1 and relative[1] == ":"):
                relative = portable
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
        return figure, [ResourceDescriptor(resource_id, media, sha, portable)]
    diagnostics.append(finding(
        "warning", "markdown.unsupported_block", f"unsupported markdown block kind: {block.kind}",
        category="loss",
    ))
    return BlockNode(counter.block(), "opaque", attrs=(("origin", block.kind),), provenance=provenance), []


__all__ = ["MarkdownReader", "MarkdownWriter", "copy_figure_files", "render_latex", "render_markdown"]
