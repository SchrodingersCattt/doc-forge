"""Thin command-line entry points for docforge.

Each subcommand lazily imports its feature module so the minimum install stays
functional even when optional dependencies (matplotlib, openai, google-genai)
are absent.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docforge",
        description=__doc__,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    md = sub.add_parser("md2docx", help="Markdown -> DOCX (block parser + renderer)")
    md.add_argument("inputs", nargs="*", type=Path, help="Markdown files to merge in order")
    md.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    md.add_argument("--keep-comments", action="store_true", help="Keep Markdown HTML comments as red TODO notes")
    md.add_argument("--title", help="Optional title for a standalone document")
    md.add_argument("--template", type=Path, help="DOCX template to preserve while assembling Markdown")
    md.add_argument("--metadata", type=Path, help="Markdown metadata/front-matter file for template assembly")
    md.add_argument("--bibliography", type=Path, help="Ordered JSON bibliography for citation expansion")
    md.add_argument("--citation-base", type=Path, help="Existing manifest whose citation numbers a supplement should reuse")
    md.add_argument("--citation-format", choices=("template", "bracketed", "superscript", "superscript-bracketed"), default="template", help="Citation rendering policy")
    md.add_argument("--columns", choices=("template", "one", "two"), default="template", help="Body column layout for template assembly")
    md.add_argument("--figure-span", choices=("column", "page"), default="column", help="Default figure span: current text column or printable page width; per-image span= overrides this value")
    md.add_argument("--font", dest="font_family", help="Explicit Latin font override for generated text")
    md.add_argument("--east-asia-font", dest="east_asia_font", help="Explicit East Asian font override for generated text")
    md.add_argument("--body-font-size", type=float, help="Generated body text size in points")
    md.add_argument("--abstract-font-size", type=float, help="Generated abstract text size in points")
    md.add_argument("--caption-font-size", type=float, help="Generated figure and table caption size in points")
    md.add_argument("--reference-font-size", type=float, help="Generated bibliography entry size in points")
    md.add_argument("--style", dest="style_profile", default="template", help="Template style profile or JSON semantic-style map")
    md.add_argument("--line-numbers", choices=("template", "on", "off"), default="template", help="Line-number policy for generated sections")
    md.add_argument("--heading-before", type=float, default=None, help="Explicit heading spacing before in points")
    md.add_argument("--native-toc", action="store_true", help="Insert a native Word table of contents after template front matter")
    md.add_argument("--restart-heading-numbering", action="store_true", help="Restart H2 decimal numbering after each H1")
    md.add_argument("--body-first-line-chars", type=float, default=None, help="Body paragraph first-line indent in character units")
    md.add_argument("--page-break-before-h1", action="store_true", help="Start each level-one Markdown heading on a new page")
    md.add_argument("--numbering-prefix", default="", help="Prefix for figure, table-caption, and equation numbers (e.g. S for SI)")
    md.add_argument("--bibliography-scope", choices=("auto", "all", "new-only"), default="auto", help="Reference entries to render: auto uses new-only with --citation-base and all otherwise")
    md.add_argument("--citation-numbering", choices=("first-citation", "source-order"), default="first-citation", help="Citation numbering policy")
    md.add_argument("--bibliography-profile", choices=("markdown", "plain"), default="markdown", help="Bibliography formatter profile")
    md.add_argument("--omit-metadata-back-matter", action="store_true", help="Omit acknowledgment, author-contribution, and code-availability sections")
    md.add_argument("--no-title", action="store_true", help="Remove level-one Markdown headings; retain the metadata title block")
    md.add_argument("--skip-images", action="store_true", help="Skip Markdown body images")
    md.add_argument("--force", action="store_true", help="Overwrite an existing output file")
    md.add_argument(
        "--roundtrip-manifest",
        type=Path,
        help="Manifest from docx2md; reuse its exact source DOCX when all sections and media are unchanged",
    )

    contract = sub.add_parser(
        "contract",
        help="Atomically build, validate, and audit a semantic-role DOCX",
    )
    contract.add_argument("inputs", nargs="+", type=Path, help="Markdown files to merge in order")
    contract.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    contract.add_argument("--template", type=Path, required=True, help="DOCX template with semantic paragraph roles")
    contract.add_argument("--metadata", type=Path, required=True, help="Metadata Markdown file")
    contract.add_argument("--bibliography", type=Path)
    contract.add_argument("--citation-base", type=Path)
    contract.add_argument("--citation-format", choices=("template", "bracketed", "superscript", "superscript-bracketed"), default="template")
    contract.add_argument("--title", default="")
    contract.add_argument("--style", dest="style_profile", default="template")
    contract.add_argument("--columns", choices=("template", "one", "two"), default="template")
    contract.add_argument("--figure-span", choices=("column", "page"), default="column")
    contract.add_argument("--line-numbers", choices=("template", "on", "off"), default="template")
    contract.add_argument("--font", dest="font_family")
    contract.add_argument("--east-asia-font", dest="east_asia_font")
    contract.add_argument("--keep-comments", action="store_true")
    contract.add_argument("--no-title", action="store_true")
    contract.add_argument("--native-toc", action="store_true")
    contract.add_argument("--force", action="store_true", help="Replace existing output and sidecars")

    docx_md = sub.add_parser("docx2md", help="DOCX -> GitHub-Flavored Markdown via Pandoc")
    docx_md.add_argument("input", type=Path, help="Input DOCX path")
    docx_md.add_argument("-o", "--output", type=Path, help="Output Markdown path")
    docx_md.add_argument("--split-dir", type=Path, help="Directory for section Markdown files")
    docx_md.add_argument("--section-map", type=Path, help="JSON section map used with --split-dir")
    docx_md.add_argument("--media-dir", type=Path, help="Pandoc extraction root (contains media/)")
    docx_md.add_argument("--track-changes", choices=("accept", "reject", "all"), default="accept")
    docx_md.add_argument("--force", action="store_true", help="Overwrite existing output files")

    mdx = sub.add_parser("md-export", help="DOCX -> reversible Markdown bundle (docs/docxmd-syntax.md)")
    mdx.add_argument("input", type=Path, help="Input DOCX path")
    mdx.add_argument("bundle", type=Path, help="Bundle directory (bundle.json, Markdown parts, media, references)")
    mdx.add_argument("--id", dest="doc_id", required=True, help="Document id inside the bundle, e.g. main or si")
    mdx.add_argument("--split", type=Path, help='JSON list of {"file": NAME, "start": REGEX} part boundaries')
    mdx.add_argument("--bibliography", default="references.json", help="References JSON file name inside the bundle")
    mdx.add_argument("--bib-style", help="Style key for this document's references (default: the document id)")
    mdx.add_argument("--no-citations", action="store_true", help="Keep superscript citation numbers as plain text")
    mdx.add_argument("--force", action="store_true", help="Re-export a document that is already in the bundle")

    mdb = sub.add_parser("md-build", help="Reversible Markdown bundle -> DOCX")
    mdb.add_argument("bundle", type=Path, help="Bundle directory")
    mdb.add_argument("--id", dest="doc_id", required=True, help="Document id inside the bundle")
    mdb.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    mdb.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    mdr = sub.add_parser("md-roundtrip", help="Export, rebuild and compare a DOCX with its Markdown bundle")
    mdr.add_argument("input", type=Path, help="Input DOCX path")
    mdr.add_argument("bundle", type=Path, help="Bundle directory to (re)write")
    mdr.add_argument("--id", dest="doc_id", required=True, help="Document id inside the bundle")
    mdr.add_argument("--split", type=Path, help="Part boundaries JSON (see md-export)")
    mdr.add_argument("--bibliography", default="references.json", help="References JSON file name inside the bundle")
    mdr.add_argument("--bib-style", help="Style key for this document's references")
    mdr.add_argument("-o", "--output", type=Path, help="Rebuilt DOCX path (default: BUNDLE/_roundtrip/ID.docx)")
    mdr.add_argument("--render", choices=("none", "auto", "word", "libreoffice"), default="none", help="Also render both files and compare pages")
    mdr.add_argument("--report", type=Path, help="Write the JSON report here")
    mdr.add_argument("--force", action="store_true", help="Re-export a document that is already in the bundle")

    delta = sub.add_parser(
        "apply-delta",
        help="Apply a Markdown paragraph delta onto a source DOCX without redrawing it",
    )
    delta.add_argument("source", type=Path, help="Source DOCX whose layout is preserved")
    delta.add_argument("baseline", type=Path, help="Markdown exported from the source DOCX")
    delta.add_argument("edited", type=Path, help="Edited Markdown; only matched paragraph changes are applied")
    delta.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    delta.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    delivery = sub.add_parser("deliver", help="Build named article/supplement DOCX files from a JSON profile")
    delivery.add_argument("--profile", type=Path, required=True, help="Delivery profile JSON")
    delivery.add_argument("--accept-revisions", action="store_true", help="Apply reviewed DOCX revisions to mapped Markdown sources before building")
    delivery.add_argument("--force", action="store_true", help="Overwrite delivery outputs and converted sources")

    md2 = sub.add_parser("redline", help="Create a DOCX with native Word revisions against a reviewed DOCX")
    # Positional paths remain supported for existing scripts.  The explicit
    # flags make the review operation self-documenting and are convenient in
    # CI: ``redline --baseline REVIEWED --current CURRENT --output TRACKED``.
    md2.add_argument("base", nargs="?", type=Path, help="Reviewed DOCX baseline (or use --baseline)")
    md2.add_argument("current", nargs="?", type=Path, help="Freshly generated DOCX (or use --current)")
    md2.add_argument("-o", "--output", dest="output", type=Path, help="Output tracked DOCX path")
    md2.add_argument("--baseline", dest="baseline", type=Path, help="Reviewed DOCX baseline")
    md2.add_argument("--current", dest="current_option", type=Path, help="Freshly generated DOCX")
    md2.add_argument("--revision-author", default="M.Y.G.", help="Author recorded for revisions")
    md2.add_argument("--accept-baseline", action="store_true", help="Accept earlier baseline revisions before creating this redline")
    md2.add_argument("--report", dest="report_path", type=Path, help="Write a JSON action/reason/ratio alignment report")
    md2.add_argument("--force", action="store_true", help="Overwrite an existing output file")
    md2.add_argument("--ratio-cache", type=Path, help="JSON file that keeps paragraph similarity ratios between runs")
    md2.add_argument("--workers", type=int, help="Processes for uncached ratios; 1 disables the process pool")

    tex = sub.add_parser("tex2docx", help="LaTeX -> DOCX (main + optional SI)")
    tex.add_argument("main", type=Path, help="main.tex path")
    tex.add_argument("--si", type=Path, help="si.tex path")
    tex.add_argument("--bib", type=Path, help="ref.bib path")
    tex.add_argument("--target", choices=("main", "si"), default="main", help="TeX source to render; the other source is used for cross-references")
    tex.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    tex.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    docx_tex = sub.add_parser("docx2tex", help="DOCX -> standalone LaTeX via Pandoc")
    docx_tex.add_argument("input", type=Path, help="Input DOCX path")
    docx_tex.add_argument("-o", "--output", type=Path, required=True, help="Output TeX path")
    docx_tex.add_argument("--media-dir", type=Path, help="Pandoc extraction root (contains media/)")
    docx_tex.add_argument("--track-changes", choices=("accept", "reject", "all"), default="accept")
    docx_tex.add_argument("--body-only", action="store_true", help="Write Pandoc body without a preamble")
    docx_tex.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    aigc = sub.add_parser("aigc", help="Generate figures from Markdown prompt files")
    aigc.add_argument("prompt_dir", type=Path, help="Directory containing *_prompt.md files")
    aigc.add_argument("--output", type=Path, default=None, help="Output directory (default: next to prompts)")
    aigc.add_argument("--backend", choices=["litellm", "google"], default="litellm")
    aigc.add_argument("--prompt", action="append", default=[], help="Prompt stem or filename; repeatable")
    aigc.add_argument("--all", action="store_true", help="Generate every Markdown prompt")
    aigc.add_argument("--list", action="store_true", help="List discovered prompt files")
    aigc.add_argument("--dry-run", action="store_true", help="Validate prompts and configuration without API calls")
    aigc.add_argument("--env", type=Path, default=None, help=".env file with BASE_URL / API keys")
    aigc.add_argument("--overwrite", action="store_true", help="Replace canonical image and sidecar")

    pack = sub.add_parser("sourcepack", help="Build an auditable TeX source ZIP from an allowlist")
    pack.add_argument("root", type=Path, help="Repository root")
    pack.add_argument("--allow", action="append", default=[], help="Relative path to include; repeatable")
    pack.add_argument("--glob-dir", action="append", default=[], help="Directory to include recursively; repeatable")
    pack.add_argument("-o", "--output", type=Path, default=None, help="Output ZIP path")

    check = sub.add_parser("check", help="Validate a generated DOCX (headings, images, clean review state)")
    check.add_argument("docx", type=Path, help="DOCX to validate")
    check.add_argument("--verify-clean", action="store_true", help="Require accepted revisions and no comments")
    check.add_argument("--verify-template", action="store_true", help="Check template placeholders, citations, sections, and image parts")

    gate = sub.add_parser("gate", help="Run project-neutral source quality gates")
    gate_sub = gate.add_subparsers(dest="gate_kind", required=True)
    abbr = gate_sub.add_parser("abbr", help="Check abbreviations and project whitelist")
    abbr.add_argument("paths", nargs="+", type=Path)
    abbr.add_argument("--config", type=Path, help="JSON config containing abbr-whitelist")
    abbr.add_argument("--project", help="Project key under config.projects")
    abbr.add_argument("--allow-undefined", action="store_true")
    abbr.add_argument("--require-reuse", action="store_true")
    refs = gate_sub.add_parser("references", help="Check TeX figure/table references")
    refs.add_argument("--root", type=Path, default=Path.cwd())
    refs.add_argument("--config", type=Path, required=True)
    style = gate_sub.add_parser("style", help="Check configurable prose style rules")
    style.add_argument("paths", nargs="*", type=Path)
    style.add_argument("--config", type=Path, required=True)
    style.add_argument("--baseline", type=Path)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    raw_argv = list(argv) if argv is not None else sys.argv[1:]
    args = parser.parse_args(raw_argv)
    args._command_argv = ["docforge", *raw_argv]
    try:
        if args.command == "md2docx":
            return _cmd_md2docx(args)
        if args.command == "contract":
            return _cmd_contract(args)
        if args.command == "docx2md":
            return _cmd_docx2md(args)
        if args.command == "md-export":
            return _cmd_md_export(args)
        if args.command == "md-build":
            return _cmd_md_build(args)
        if args.command == "md-roundtrip":
            return _cmd_md_roundtrip(args)
        if args.command == "apply-delta":
            return _cmd_apply_delta(args)
        if args.command == "redline":
            return _cmd_redline(args)
        if args.command == "deliver":
            return _cmd_deliver(args)
        if args.command == "tex2docx":
            return _cmd_tex2docx(args)
        if args.command == "docx2tex":
            return _cmd_docx2tex(args)
        if args.command == "aigc":
            return _cmd_aigc(args)
        if args.command == "sourcepack":
            return _cmd_sourcepack(args)
        if args.command == "check":
            return _cmd_check(args)
        if args.command == "gate":
            return _cmd_gate(args)
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    parser.error(f"Unknown command: {args.command}")
    return 2


def _cmd_md2docx(args: argparse.Namespace) -> int:
    from ..markdown import parse_markdown, render_blocks_to_doc, convert_unicode_scripts_in_docx
    from ..markdown import reuse_unchanged_roundtrip_source
    from ..output import validate_output_path

    validate_output_path(args.output)
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    if args.roundtrip_manifest is not None:
        if reuse_unchanged_roundtrip_source(args.roundtrip_manifest, args.output, force=args.force):
            print(f"roundtrip exact source reused: {args.output}")
            return 0
        if not args.inputs:
            raise ValueError("Markdown inputs are required when the roundtrip bundle has changed")
    if not args.inputs and args.roundtrip_manifest is None:
        raise ValueError("md2docx requires Markdown inputs or --roundtrip-manifest")
    if args.template is not None:
        from ..markdown import assemble_markdown_template, write_assembly_sidecars

        if args.metadata is None:
            raise ValueError("Template assembly requires --metadata")
        result = assemble_markdown_template(
            args.inputs,
            template_path=args.template,
            output=args.output,
            metadata_path=args.metadata,
            bibliography_path=args.bibliography,
            citation_base_path=args.citation_base,
            citation_format=args.citation_format,
            title=args.title or "",
            keep_comments=args.keep_comments,
            skip_images=args.skip_images,
            columns=args.columns,
            figure_span=args.figure_span,
            font_family=args.font_family,
            east_asia_font=args.east_asia_font,
            style_profile=args.style_profile,
            line_numbers=args.line_numbers,
            include_title=True,
            strip_level_one_headings=args.no_title,
            heading_before=args.heading_before,
            native_toc=args.native_toc,
            restart_heading_numbering=args.restart_heading_numbering,
            body_first_line_chars=args.body_first_line_chars,
            page_break_before_h1=args.page_break_before_h1,
            body_font_size=args.body_font_size,
            abstract_font_size=args.abstract_font_size,
            caption_font_size=args.caption_font_size,
            reference_font_size=args.reference_font_size,
            numbering_prefix=args.numbering_prefix,
            bibliography_scope=args.bibliography_scope,
            citation_numbering=args.citation_numbering,
            bibliography_profile=args.bibliography_profile,
            include_metadata_back_matter=not args.omit_metadata_back_matter,
            force=args.force,
        )
        manifest, checksum = write_assembly_sidecars(
            result,
            inputs=args.inputs,
            template_path=args.template,
            metadata_path=args.metadata,
            bibliography_path=args.bibliography,
            citation_base_path=args.citation_base,
            command=args._command_argv,
        )
        print(result.output)
        print(manifest)
        print(checksum)
        return 0
    if (
        args.metadata is not None
        or args.bibliography is not None
        or args.citation_base is not None
        or args.style_profile != "template"
        or args.line_numbers != "template"
        or args.figure_span != "column"
        or args.body_font_size is not None
        or args.abstract_font_size is not None
        or args.caption_font_size is not None
        or args.reference_font_size is not None
    ):
        raise ValueError("Metadata, bibliography, style, line-number, figure-span, and typography options require --template")
    blocks = [
        block
        for path in args.inputs
        for block in parse_markdown(path, strip_comments=not args.keep_comments)
    ]
    if args.skip_images:
        blocks = [block for block in blocks if block.kind != "image"]
    if args.no_title:
        blocks = [
            block for block in blocks
            if not (block.kind == "heading" and block.level == 1)
        ]
    doc = render_blocks_to_doc(
        blocks,
        title=args.title,
        number_prefix=args.numbering_prefix,
        heading_before=args.heading_before,
        font_family=args.font_family,
        east_asia_font=args.east_asia_font,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.output))
    convert_unicode_scripts_in_docx(args.output)
    print(args.output)
    return 0


def _load_split(path: Path | None) -> list[dict] | None:
    return json.loads(path.read_text(encoding="utf-8-sig")) if path else None


def _cmd_contract(args: argparse.Namespace) -> int:
    from ..markdown import build_docx_contract

    result = build_docx_contract(
        args.inputs,
        template_path=args.template,
        metadata_path=args.metadata,
        bibliography_path=args.bibliography,
        citation_base_path=args.citation_base,
        output=args.output,
        force=args.force,
        command=args._command_argv,
        title=args.title,
        citation_format=args.citation_format,
        style_profile=args.style_profile,
        columns=args.columns,
        figure_span=args.figure_span,
        line_numbers=args.line_numbers,
        font_family=args.font_family,
        east_asia_font=args.east_asia_font,
        keep_comments=args.keep_comments,
        strip_level_one_headings=args.no_title,
        native_toc=args.native_toc,
    )
    print(result.output)
    print(result.manifest)
    print(result.checksum)
    return 0


def _guard_bundle(args: argparse.Namespace) -> None:
    manifest = args.bundle / "bundle.json"
    if manifest.exists() and not args.force:
        documents = json.loads(manifest.read_text(encoding="utf-8")).get("documents", [])
        if any(doc.get("id") == args.doc_id for doc in documents):
            raise FileExistsError(f"Bundle already has document {args.doc_id!r}; pass --force to re-export over its Markdown")


def _cmd_md_export(args: argparse.Namespace) -> int:
    from ..docxmd import ExportOptions, export_docx

    _guard_bundle(args)
    report = export_docx(args.input, args.bundle, ExportOptions(
        args.doc_id, _load_split(args.split), args.bibliography, args.bib_style, citations=not args.no_citations,
    ))
    print(json.dumps(report, indent=1, ensure_ascii=False, default=str))
    return 0


def _cmd_md_build(args: argparse.Namespace) -> int:
    from ..docxmd import build_docx

    build_docx(args.bundle, args.doc_id, args.output, overwrite=args.force)
    print(args.output)
    return 0


def _cmd_md_roundtrip(args: argparse.Namespace) -> int:
    from ..docxmd import compare_pdfs, render_pdfs, roundtrip

    _guard_bundle(args)
    report = roundtrip(
        args.input, args.bundle, args.doc_id, split=_load_split(args.split),
        bibliography=args.bibliography, bib_style=args.bib_style, output=args.output,
    )
    if args.render != "none":
        out = Path(report["rebuilt"]).parent / "pdf"
        source_copy = out / f"{args.doc_id}.source.docx"
        out.mkdir(parents=True, exist_ok=True)
        source_copy.write_bytes(args.input.read_bytes())
        left, right = render_pdfs([source_copy, Path(report["rebuilt"])], out, args.render)
        pages = compare_pdfs(left, right)
        report["render"] = {"engine": args.render, "pages": pages}
        report["ok"] = report["ok"] and all(page["same"] for page in pages)
    text = json.dumps(report, indent=1, ensure_ascii=False, default=str)
    if args.report:
        args.report.write_text(text, encoding="utf-8")
    print(text)
    return 0 if report["ok"] else 1


def _cmd_apply_delta(args: argparse.Namespace) -> int:
    from ..markdown.source_edit import apply_markdown_delta

    summary = apply_markdown_delta(
        args.source,
        args.baseline,
        args.edited,
        args.output,
        overwrite=args.force,
    )
    print("markdown delta: " + ", ".join(f"{key}={value}" for key, value in summary.items()))
    print(args.output)
    return 0


def _cmd_deliver(args: argparse.Namespace) -> int:
    from ..markdown import deliver

    result = deliver(args.profile, accept_revisions=args.accept_revisions, force=args.force)
    for artifact in result.artifacts:
        print(artifact.output)
        print(artifact.manifest)
        print(artifact.checksum)
    print(result.manifest)
    return 0


def _cmd_redline(args: argparse.Namespace) -> int:
    from ..docxdiff import create_tracked_docx
    from ..output import validate_output_path

    baseline = args.baseline or args.base
    current = args.current_option or args.current
    if baseline is None or current is None:
        raise ValueError("redline requires a baseline and current DOCX (use positional paths or --baseline/--current)")
    if args.output is None:
        raise ValueError("redline requires --output/-o")
    validate_output_path(args.output)
    if args.report_path is not None:
        validate_output_path(args.report_path, label="report")
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    summary = create_tracked_docx(
        baseline,
        current,
        args.output,
        author=args.revision_author,
        overwrite=args.force,
        preserve_base_revisions=not args.accept_baseline,
        accept_baseline=args.accept_baseline,
        report_path=args.report_path,
        ratio_cache=args.ratio_cache,
        workers=args.workers,
    )
    print("tracked revisions: " + ", ".join(f"{name}={count}" for name, count in summary.items()))
    if args.report_path is not None:
        print(args.report_path)
    print(args.output)
    return 0


def _cmd_docx2md(args: argparse.Namespace) -> int:
    from ..markdown import docx_to_markdown
    from ..output import validate_output_path

    validate_output_path(args.output)
    validate_output_path(args.split_dir, label="split output directory")
    if args.output is None and args.split_dir is None:
        raise ValueError("docx2md requires --output, --split-dir, or both")
    result = docx_to_markdown(
        args.input,
        output=args.output,
        split_dir=args.split_dir,
        section_map_path=args.section_map,
        media_dir=args.media_dir,
        track_changes=args.track_changes,
        force=args.force,
    )
    if result.output is not None:
        print(result.output)
    if result.split_dir is not None:
        for path in result.sections:
            print(path)
        print(result.split_dir / "manifest.json")
    return 0


def _cmd_tex2docx(args: argparse.Namespace) -> int:
    from ..tex import build_validated_docx, convert_files
    from ..output import validate_output_path

    validate_output_path(args.output)
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    source = args.main if args.target == "main" else args.si
    if source is None:
        raise ValueError("--target si requires --si")

    def build(path: Path) -> None:
        convert_files(
            args.main,
            si_path=args.si,
            bib_path=args.bib,
            output=path,
            target=args.target,
            citation_prefix="S" if args.target == "si" else "",
            float_prefix="S" if args.target == "si" else "",
        )

    build_validated_docx(
        args.output,
        source,
        build,
        atomic=True,
        auxiliary_tex_path=args.si if args.target == "main" else args.main,
    )
    print(args.output)
    return 0


def _cmd_docx2tex(args: argparse.Namespace) -> int:
    from ..tex import docx_to_tex
    from ..output import validate_output_path

    validate_output_path(args.output)
    result = docx_to_tex(
        args.input,
        output=args.output,
        media_dir=args.media_dir,
        track_changes=args.track_changes,
        body_only=args.body_only,
        force=args.force,
    )
    print(result.output)
    return 0


def _cmd_aigc(args: argparse.Namespace) -> int:
    from ..aigc import (
        discover_prompts,
        generate_one,
        load_configuration,
        load_prompt,
        select_prompts,
    )

    output_dir = args.output or (args.prompt_dir.parent / "aigc")
    paths = discover_prompts(args.prompt_dir)
    if args.list:
        for path in paths:
            spec = load_prompt(path, root=args.prompt_dir)
            print(
                f"{path.stem}: output={spec.output_name}, status={spec.status}, "
                f"ratio={spec.aspect_ratio}"
            )
        return 0
    selected = select_prompts(paths, args.prompt, args.all)
    if not selected:
        print("No prompts selected. Use --prompt NAME or --all.", file=sys.stderr)
        return 2
    config = load_configuration(args.env)
    if args.dry_run:
        if args.backend == "litellm":
            missing = [name for name in ("base_url", "litellm_key", "image_model") if not config[name]]
        else:
            missing = ["gemini_key"] if not config["gemini_key"] else []
        for path in selected:
            spec = load_prompt(path, root=args.prompt_dir)
            print(
                f"OK {spec.source}: output={spec.output_name}, "
                f"prompt_sha256={spec.source_sha256[:12]}, "
                f"aspect_ratio={spec.aspect_ratio}, image_size={spec.image_size}"
            )
        if missing:
            print(f"Configuration missing required variable names: {', '.join(missing)}", file=sys.stderr)
            return 1
        print(f"Configuration is present for backend={args.backend}; values were not displayed.")
        return 0
    failures = 0
    for path in selected:
        spec = load_prompt(path, root=args.prompt_dir)
        print(f"Generating {spec.output_name} from {spec.source} using {args.backend}...")
        try:
            result = generate_one(
                spec,
                args.backend,
                config,
                output_dir=output_dir,
                overwrite=args.overwrite,
                root=args.prompt_dir.parent,
            )
            print(f"Saved {result.output}")
        except Exception as exc:  # noqa: BLE001
            failures += 1
            print(f"FAILED {spec.output_name}: {type(exc).__name__}: {exc}", file=sys.stderr)
    return 1 if failures else 0


def _cmd_sourcepack(args: argparse.Namespace) -> int:
    from ..sourcepack import package_files

    if not args.allow and not args.glob_dir:
        print("error: provide at least one --allow path or --glob-dir", file=sys.stderr)
        return 2
    output = package_files(
        args.root,
        allowlist=args.allow,
        glob_dirs=[Path(value) for value in args.glob_dir],
        output=args.output,
    )
    print(f"Created {output} ({output.stat().st_size / 1024 / 1024:.2f} MB)")
    return 0


def _cmd_check(args: argparse.Namespace) -> int:
    from ..markdown import verify_navigation_headings, verify_image_relationships, verify_clean_review_state

    verify_navigation_headings(args.docx)
    verify_image_relationships(args.docx)
    if args.verify_clean:
        verify_clean_review_state(args.docx)
    if args.verify_template:
        from ..markdown import verify_template_output

        verify_template_output(args.docx)
    print(f"OK {args.docx}")
    return 0


def _cmd_gate(args: argparse.Namespace) -> int:
    if args.gate_kind == "abbr":
        import json
        from ..gates import check_abbreviations, load_abbreviation_whitelist

        whitelist = load_abbreviation_whitelist(args.config, project=args.project)
        findings = check_abbreviations(
            args.paths,
            whitelist=whitelist,
            require_definition=not args.allow_undefined,
            require_reuse=args.require_reuse,
        )
        for finding in findings:
            print(f"{finding.path}:{finding.line}: {finding.rule}: {finding.message}", file=sys.stderr)
        return 1 if findings else 0
    if args.gate_kind == "references":
        import json
        from ..gates import audit_references

        config = json.loads(args.config.read_text(encoding="utf-8-sig"))
        issues = audit_references(args.root, config)
        for issue in issues:
            print(f"{issue.kind} {issue.label}: {issue.message}", file=sys.stderr)
        return 1 if issues else 0
    if args.gate_kind == "style":
        import json
        from ..gates import check_style, load_style_config, unresolved

        config = load_style_config(args.config)
        paths = args.paths
        if not paths:
            configured = config.get("documents", [])
            if not isinstance(configured, list) or not configured:
                raise ValueError("style gate requires paths or a documents list in config")
            paths = [Path(item) for item in configured]
        baseline = json.loads(args.baseline.read_text(encoding="utf-8-sig")) if args.baseline else None
        findings = unresolved(check_style(paths, config=config), baseline)
        for finding in findings:
            print(f"{finding.path}:{finding.line}: {finding.rule}: {finding.message}", file=sys.stderr)
        return 1 if findings else 0
    raise ValueError(f"Unknown gate: {args.gate_kind}")


if __name__ == "__main__":
    raise SystemExit(main())
