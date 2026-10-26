"""Public conversion API.

``read``, ``write``, ``convert``, ``restyle``, and ``audit`` are the stable
entry points. Existing ``md2docx`` and ``tex2docx`` commands call the same
legacy builders, so those outputs stay on the historical renderer.
"""

from __future__ import annotations

import hashlib
from dataclasses import replace
from pathlib import Path
from typing import Mapping

from .adapters.protocol import (
    ConversionError,
    LossReport,
    ReadOptions,
    ReadResult,
    WriteOptions,
    WriteResult,
    default_loss_policy,
    ensure_publishable,
    finding,
    is_loss,
    require_valid_document,
)
from .pipeline.models import Diagnostic, ValidationReport
from .semantic import SemanticDocument


def read(source: Path | str, *, fidelity: str = "semantic", loss_policy: str | None = None) -> ReadResult:
    """Read a Markdown, LaTeX, or DOCX file into a SemanticDocument."""

    path = Path(source)
    if fidelity not in {"semantic", "preserve"}:
        raise ConversionError("fidelity must be semantic or preserve", code="FIDELITY")
    policy = default_loss_policy(fidelity, loss_policy)  # type: ignore[arg-type]
    fmt = sniff_source(path)
    reader = _readers()[fmt]
    result = reader.read(path, options=ReadOptions(fidelity=fidelity, loss_policy=policy))  # type: ignore[arg-type]
    diagnostics = _escalate(result.diagnostics, fidelity)
    return replace(result, diagnostics=diagnostics, origin=path.resolve())


def write(
    loaded: ReadResult | SemanticDocument,
    output: Path | str,
    *,
    fidelity: str = "semantic",
    loss_policy: str | None = None,
    force: bool = False,
    style: object | None = None,
) -> WriteResult:
    """Write a document.

    An unchanged DOCX is copied only after its on-disk SHA-256 still matches
    the snapshot from ``read``. An edited DOCX is refused at ``preserve``
    fidelity until source patching exists. Unchanged Markdown or TeX is
    rebuilt by the legacy renderer.
    """

    destination = Path(output)
    if fidelity not in {"semantic", "preserve"}:
        raise ConversionError("fidelity must be semantic or preserve", code="FIDELITY")
    policy = default_loss_policy(fidelity, loss_policy)  # type: ignore[arg-type]
    if isinstance(loaded, ReadResult):
        if loaded.origin_format == "docx" and fidelity == "preserve":
            if not loaded.unchanged():
                raise ConversionError(
                    "edited DOCX cannot be written with preserve fidelity until source patching exists",
                    code="PRESERVE_EDIT",
                )
            _require_fresh_source(loaded)
            return convert(
                loaded.origin, destination, fidelity=fidelity, loss_policy=policy, force=force, style=style,
            )
        if loaded.origin is not None and loaded.unchanged():
            _require_fresh_source(loaded)
            return convert(
                loaded.origin, destination, fidelity=fidelity, loss_policy=policy, force=force, style=style,
            )
        document = loaded.document
        origin = loaded.origin
        diagnostics = loaded.diagnostics
    else:
        document = loaded
        origin = None
        diagnostics = ()
    return _write_semantic(
        document, destination, fidelity=fidelity, loss_policy=policy, force=force, style=style,
        diagnostics=diagnostics, origin=origin,
    )


def convert(
    source: Path | str,
    output: Path | str,
    *,
    fidelity: str = "semantic",
    loss_policy: str | None = None,
    style: object | None = None,
    force: bool = False,
    title: str | None = None,
    si: Path | str | None = None,
    bib: Path | str | None = None,
    target: str = "main",
) -> WriteResult:
    """Convert between Markdown, LaTeX, and DOCX."""

    src = Path(source)
    dest = Path(output)
    if fidelity not in {"semantic", "preserve"}:
        raise ConversionError("fidelity must be semantic or preserve", code="FIDELITY")
    policy = default_loss_policy(fidelity, loss_policy)  # type: ignore[arg-type]
    source_format = sniff_source(src)
    target_format = sniff_target(dest)
    if source_format == "docx" and target_format != "docx" and fidelity == "preserve":
        raise ConversionError(
            "preserve fidelity does not rebuild DOCX as another format",
            code="LOSS_POLICY",
            diagnostics=(finding("error", "fidelity.preserve", "cross-format output would drop native DOCX constructs"),),
        )
    if source_format == target_format == "docx" and style is None:
        from .roundtrip import SourceSnapshot, exact_reuse

        exact_reuse(src, dest, snapshot=SourceSnapshot.from_file(src), force=force)
        report = LossReport(policy)
        return WriteResult(dest, "EXACT_REUSE", report, audit(dest))
    if source_format == target_format == "docx" and style is not None:
        return restyle(src, dest, profile=style, loss_policy=policy, force=force)
    if {source_format, target_format} == {"markdown", "latex"} or (
        target_format in {"markdown", "latex"} and source_format in {"markdown", "latex"}
    ):
        loaded = read(src, fidelity=fidelity, loss_policy=policy)
        ensure_publishable(LossReport(policy, loaded.diagnostics))
        return _write_semantic(
            loaded.document, dest, fidelity=fidelity, loss_policy=policy, force=force,
            diagnostics=loaded.diagnostics, origin=src,
        )
    if target_format == "docx" and source_format in {"markdown", "latex"}:
        return _legacy_docx(
            src, dest, source_format=source_format, policy=policy, force=force, style=style,
            title=title, si=Path(si) if si else None, bib=Path(bib) if bib else None, target=target,
        )
    if source_format == "docx" and target_format in {"markdown", "latex"}:
        loaded = read(src, fidelity="semantic", loss_policy=policy)
        ensure_publishable(LossReport(policy, loaded.diagnostics))
        if target_format == "markdown":
            from .markdown import docx_to_markdown

            docx_to_markdown(src, output=dest, force=force)
            message = "DOCX to Markdown is a semantic export"
        else:
            from .tex import docx_to_tex

            docx_to_tex(src, output=dest, force=force)
            message = "DOCX to LaTeX is a semantic export"
        notice = finding("warning", "docx.semantic_export", message)
        return WriteResult(dest, "REBUILD", LossReport(policy, loaded.diagnostics + (notice,)), ValidationReport("pass"))
    raise ConversionError(f"unsupported conversion: {source_format} -> {target_format}", code="UNSUPPORTED_CONVERSION")


def restyle(
    source: Path | str,
    output: Path | str,
    *,
    profile: object,
    preserve_content: bool = True,
    loss_policy: str | None = None,
    force: bool = False,
) -> WriteResult:
    """Restyle a DOCX from a StyleProfile without rewriting its text."""

    from .style.engine import restyle_docx

    policy = default_loss_policy("preserve", loss_policy)  # type: ignore[arg-type]
    summary = restyle_docx(Path(source), Path(output), profile, force=force, preserve_content=preserve_content)  # type: ignore[arg-type]
    report = LossReport(policy, (finding("info", "style.applied", "style profile applied to the source package"),))
    validation = audit(output)
    if not validation.ok:
        raise ConversionError("restyle audit failed", code="AUDIT_FAILED", diagnostics=validation.diagnostics)
    return WriteResult(Path(output), str(summary["execution_mode"]), report, validation, None)  # type: ignore[arg-type]


def audit(path: Path | str) -> ValidationReport:
    """Audit a DOCX package. Non-DOCX paths are reported as unavailable."""

    target = Path(path)
    if target.suffix.casefold() not in {".docx", ".docm", ".dotx", ".dotm"}:
        return ValidationReport("unavailable", (finding("info", "audit.skipped", "audit currently checks DOCX packages"),))
    from .adapters.docx import audit_docx

    return audit_docx(target)


def sniff_source(path: Path) -> str:
    path = Path(path)
    suffix = path.suffix.casefold()
    if suffix in {".md", ".markdown"}:
        if not path.is_file():
            raise FileNotFoundError(path)
        return "markdown"
    if suffix in {".tex", ".latex"}:
        if not path.is_file():
            raise FileNotFoundError(path)
        return "latex"
    if suffix in {".docx", ".docm", ".dotx", ".dotm"}:
        from .output import detect_format

        detect_format(path)
        return "docx"
    raise ConversionError("unsupported source format", code="UNSUPPORTED_FORMAT")


def sniff_target(path: Path) -> str:
    suffix = Path(path).suffix.casefold()
    if suffix in {".md", ".markdown"}:
        return "markdown"
    if suffix in {".tex", ".latex"}:
        return "latex"
    if suffix == ".docx":
        return "docx"
    raise ConversionError("unsupported target format", code="UNSUPPORTED_FORMAT")


def _readers() -> Mapping[str, object]:
    from .adapters.docx import DocxReader
    from .adapters.latex import LatexReader
    from .adapters.markdown import MarkdownReader

    return {"markdown": MarkdownReader(), "latex": LatexReader(), "docx": DocxReader()}


def _legacy_docx(source: Path, output: Path, *, source_format: str, policy: str, force: bool, style, title, si, bib, target: str) -> WriteResult:
    from .adapters.legacy import md2docx_namespace, run_md2docx, run_tex2docx, tex2docx_namespace

    loaded = read(source, fidelity="semantic", loss_policy=policy)
    ensure_publishable(LossReport(policy, loaded.diagnostics))
    if style is None:
        if source_format == "markdown":
            run_md2docx(md2docx_namespace(source, output, force=force, title=title), echo=False)
        else:
            run_tex2docx(tex2docx_namespace(source, output, si=si, bib=bib, target=target, force=force), echo=False)
        return WriteResult(output, "REBUILD", LossReport(policy, loaded.diagnostics), audit(output))
    import tempfile

    with tempfile.TemporaryDirectory(prefix="docforge-style-") as directory:
        staged = Path(directory) / "staged.docx"
        if source_format == "markdown":
            run_md2docx(md2docx_namespace(source, staged, force=True, title=title), echo=False)
        else:
            run_tex2docx(tex2docx_namespace(source, staged, si=si, bib=bib, target=target, force=True), echo=False)
        return restyle(staged, output, profile=style, loss_policy=policy, force=force)


def _write_semantic(
    document: SemanticDocument,
    output: Path,
    *,
    fidelity: str,
    loss_policy: str | None,
    force: bool,
    style=None,
    diagnostics: tuple[Diagnostic, ...] = (),
    origin: Path | None = None,
) -> WriteResult:
    from .adapters.latex import LatexWriter
    from .adapters.markdown import MarkdownWriter

    policy = default_loss_policy(fidelity, loss_policy)  # type: ignore[arg-type]
    target = sniff_target(output)
    options = WriteOptions(fidelity=fidelity, loss_policy=policy, force=force, style=style)  # type: ignore[arg-type]
    require_valid_document(document)
    if target == "latex":
        plan = LatexWriter().plan(document, source=None, options=options)
    elif target in {"markdown", "docx"}:
        plan = MarkdownWriter().plan(document, source=None, options=options)
    else:
        raise ConversionError("unsupported target format", code="UNSUPPORTED_FORMAT")
    plan = replace(plan, origin=origin)
    report = LossReport(policy, diagnostics + plan.loss_report.diagnostics, plan.loss_report.dropped)
    ensure_publishable(report)
    if target == "docx":
        import tempfile

        from .adapters.legacy import md2docx_namespace, run_md2docx
        from .adapters.markdown import render_markdown

        with tempfile.TemporaryDirectory(prefix="docforge-semantic-") as directory:
            staged = Path(directory) / "document.md"
            staged.write_text(render_markdown(document), encoding="utf-8")
            if style is None:
                run_md2docx(md2docx_namespace(staged, output, force=force), echo=False)
                return WriteResult(output, "REBUILD", report, audit(output))
            interim = Path(directory) / "interim.docx"
            run_md2docx(md2docx_namespace(staged, interim, force=True), echo=False)
            return restyle(interim, output, profile=style, loss_policy=policy, force=force)
    writer = MarkdownWriter() if target == "markdown" else LatexWriter()
    result = writer.write(replace_plan_report(plan, report), output)
    return WriteResult(result.output, "REBUILD", report, result.validation)


def replace_plan_report(plan, report: LossReport):
    from dataclasses import replace as replace_dataclass

    return replace_dataclass(plan, loss_report=report)


def _require_fresh_source(loaded: ReadResult) -> None:
    if loaded.origin is None or loaded.source is None:
        raise ConversionError("reuse needs the source snapshot", code="STALE_SOURCE")
    digest = hashlib.sha256(Path(loaded.origin).read_bytes()).hexdigest()
    if digest != loaded.source.sha256:
        raise ConversionError("source file changed after it was read", code="STALE_SOURCE")


def _escalate(diagnostics: tuple[Diagnostic, ...], fidelity: str) -> tuple[Diagnostic, ...]:
    if fidelity != "preserve":
        return diagnostics
    escalated = []
    for item in diagnostics:
        if item.severity == "warning" and is_loss(item):
            escalated.append(Diagnostic("error", item.code, item.message, details=dict(item.details)))
        else:
            escalated.append(item)
    return tuple(escalated)


__all__ = ["audit", "convert", "read", "restyle", "sniff_source", "sniff_target", "write"]
