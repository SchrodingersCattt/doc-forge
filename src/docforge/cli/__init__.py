"""Thin command-line entry points for docforge.

Each subcommand lazily imports its feature module so the minimum install stays
functional even when optional dependencies (matplotlib, openai, google-genai)
are absent.
"""

from __future__ import annotations

import argparse
import importlib
import inspect
import json
import sys
from pathlib import Path

from ..output import FormatError


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
    md.add_argument("--end-matter-placement", choices=("source", "end"), default="source", help="Place source TOC/graphical abstract sections after references on a new page (requires --template)")
    md.add_argument("--inline-math", choices=("runs", "omml"), default="runs", help="Inline math representation; omml requires Pandoc and --template")
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
    contract.add_argument("--baseline", type=Path, help="Reviewed baseline DOCX to redline against")
    contract.add_argument("--redline-output", type=Path, help="Tracked DOCX path (defaults beside the clean output)")
    contract.add_argument("--render", choices=("auto", "word", "libreoffice", "none"), default="auto", help="Render audit engine")
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

    rt = sub.add_parser(
        "roundtrip",
        help="Exact DOCX roundtrip: reuse unchanged source or patch its package",
    )
    rt.add_argument("input", type=Path, help="DOCX source")
    rt.add_argument("--workdir", type=Path, required=True, help="Markdown bundle directory")
    rt.add_argument("--output", "-o", type=Path, required=True, help="Round-tripped DOCX")
    rt.add_argument("--section-map", type=Path, help="JSON section map for the legacy sectioned DOCX workflow")
    rt.add_argument("--track-changes", choices=("accept", "reject", "all"), default="accept")
    rt.add_argument("--baseline", type=Path, help="Baseline DOCX for redline, or baseline Markdown")
    rt.add_argument("--redline", type=Path, help="Optional tracked-changes DOCX output")
    rt.add_argument("--edited", type=Path, help="Edited Markdown (source-package workflow; default: WORKDIR/full.md)")
    rt.add_argument("--block-map", type=Path, help="Explicit exported source block map")
    rt.add_argument("--allow-delete", action="store_true", help="Permit deletion of matched text paragraphs")
    rt.add_argument("--render", choices=("none", "auto", "word", "libreoffice"), default="none", help="Render patched and tracked DOCX files")
    rt.add_argument("--report", type=Path, help="Roundtrip JSON report")
    rt.add_argument("--force", action="store_true", help="Overwrite output and bundle files")

    delta = sub.add_parser(
        "apply-delta",
        help="Apply a Markdown paragraph delta onto a source DOCX without redrawing it",
    )
    delta.add_argument("source", type=Path, help="Source DOCX whose layout is preserved")
    delta.add_argument("baseline", type=Path, help="Markdown exported from the source DOCX")
    delta.add_argument("edited", type=Path, help="Edited Markdown; only matched paragraph changes are applied")
    delta.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    delta.add_argument("--block-map", type=Path, help="Exported source block map used to identify opaque blocks")
    delta.add_argument("--allow-delete", action="store_true", help="Permit deletion of matched text paragraphs")
    delta.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    block_map = sub.add_parser(
        "export-block-map",
        help="Export ordered text paragraph indices and opaque source block ids from a DOCX",
    )
    block_map.add_argument("source", type=Path, help="Source DOCX")
    block_map.add_argument("-o", "--output", type=Path, required=True, help="JSON block map path")
    block_map.add_argument("--force", action="store_true", help="Overwrite an existing map")

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
    md2.add_argument(
        "--allow-block-replace",
        action="store_true",
        help="Allow a sparse redline to insert the current document and delete the baseline",
    )

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

    pipeline = sub.add_parser("pipeline", help="Run a versioned, auditable document pipeline spec")
    pipeline.add_argument("spec", type=Path, help="JSON (or optional YAML) docforge.pipeline.v1 spec")
    pipeline.add_argument("--workspace", type=Path, default=Path.cwd(), help="Input workspace root")
    pipeline.add_argument("--output-dir", type=Path, default=None, help="Atomic publication directory (default: WORKSPACE/dist)")
    pipeline.add_argument("--run-id", default=None, help="Stable run id for reproducible receipts")
    pipeline.add_argument("--no-publish", action="store_true", help="Build and validate without replacing the publication directory")
    pipeline.add_argument("--plugin", action="append", default=[], metavar="MODULE:FUNCTION", help="Load a registry plugin (repeatable)")

    convert = sub.add_parser("convert", help="Convert Markdown, LaTeX, and DOCX through the semantic hub")
    convert.add_argument("source", type=Path)
    convert.add_argument("output", type=Path)
    convert.add_argument("--fidelity", choices=("semantic", "preserve"), default="semantic")
    convert.add_argument("--loss-policy", choices=("error", "warn", "drop"), default=None)
    convert.add_argument("--style", type=Path, help="StyleProfile JSON or YAML applied when the output is DOCX")
    convert.add_argument("--title", default=None)
    convert.add_argument("--si", type=Path, help="Second TeX file used by tex2docx")
    convert.add_argument("--bib", type=Path, help="BibTeX file used by tex2docx")
    convert.add_argument("--target", choices=("main", "si"), default="main")
    convert.add_argument("--force", action="store_true")

    restyle = sub.add_parser("restyle", help="Apply a StyleProfile to a DOCX without changing its text")
    restyle.add_argument("source", type=Path)
    restyle.add_argument("output", type=Path)
    restyle.add_argument("--profile", type=Path, required=True)
    restyle.add_argument("--loss-policy", choices=("error", "warn", "drop"), default=None)
    restyle.add_argument("--force", action="store_true")

    audit = sub.add_parser("audit", help="Audit a DOCX package and its relationships")
    audit.add_argument("docx", type=Path)

    export = sub.add_parser("export", help="Export a DOCX to a reversible bundle or semantic Markdown")
    export.add_argument("input", type=Path)
    export.add_argument("--bundle", type=Path, required=True)
    export.add_argument("--fidelity", choices=("preserve", "semantic"), default="preserve")
    export.add_argument("--id", dest="doc_id", default="main")
    export.add_argument("--force", action="store_true")

    build = sub.add_parser("build", help="Build a DOCX from an exported bundle")
    build.add_argument("bundle", type=Path)
    build.add_argument("output", type=Path)
    build.add_argument("--fidelity", choices=("preserve", "semantic"), default="preserve")
    build.add_argument("--id", dest="doc_id", default="main")
    build.add_argument("--force", action="store_true")

    return parser


def _preflight_docx(*paths: Path | None) -> None:
    """Reject non-DOCX sources before an exporter creates any directories."""

    from ..output import _detect_format

    for path in paths:
        if path is not None:
            _detect_format(path)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    raw_argv = list(argv) if argv is not None else sys.argv[1:]
    args = parser.parse_args(raw_argv)
    args._command_argv = ["docforge", *raw_argv]
    try:
        if args.command == "md2docx":
            return _cmd_md2docx(args)
        if args.command == "convert":
            return _cmd_convert(args)
        if args.command == "restyle":
            return _cmd_restyle(args)
        if args.command == "audit":
            return _cmd_audit(args)
        if args.command == "export":
            return _cmd_export(args)
        if args.command == "build":
            return _cmd_build(args)
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
        if args.command == "roundtrip":
            return _cmd_roundtrip(args)
        if args.command == "apply-delta":
            return _cmd_apply_delta(args)
        if args.command == "export-block-map":
            return _cmd_export_block_map(args)
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
        if args.command == "pipeline":
            return _cmd_pipeline(args)
    except FormatError as exc:
        # Format failures are intentionally stable and do not expose ZIP or
        # Pandoc parser text.  They still use the CLI's conventional exit 2.
        print(f"error [{exc.code}]: {exc}", file=sys.stderr)
        return 2
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    parser.error(f"Unknown command: {args.command}")
    return 2


def _cmd_pipeline(args: argparse.Namespace) -> int:
    from ..pipeline import ArtifactPresenceValidator, CopyArtifactStage, PipelineRunner, PipelineSpec, Registry

    registry = Registry()

    def copy_artifact(context):
        parameters = context.parameters
        required = ("source", "artifact_id", "role")
        missing = [key for key in required if not parameters.get(key)]
        if missing:
            raise ValueError(f"copy-artifact parameters missing: {', '.join(missing)}")
        return CopyArtifactStage(
            name=context.stage.id,
            source=str(parameters["source"]),
            artifact_id=str(parameters["artifact_id"]),
            role=str(parameters["role"]),
            output=parameters.get("output"),
            media_type=str(parameters.get("media_type", "application/octet-stream")),
            schema=parameters.get("schema"),
        ).run(context)

    def artifact_presence(context):
        required = context.parameters.get("required", context.parameters.get("artifacts", ()))
        return ArtifactPresenceValidator(tuple(str(item) for item in required)).validate(context)

    registry.register_stage("copy-artifact", copy_artifact)
    registry.register_gate("artifact-presence", artifact_presence)
    registry.register_validator("artifact-presence", artifact_presence)
    for plugin in args.plugin:
        if ":" not in plugin:
            raise ValueError(f"pipeline plugin must use MODULE:FUNCTION: {plugin}")
        module_name, function_name = plugin.split(":", 1)
        hook = getattr(importlib.import_module(module_name), function_name)
        try:
            parameters = inspect.signature(hook).parameters
        except (TypeError, ValueError):
            parameters = {"registry": object()}
        loaded = hook(registry) if parameters else hook()
        if loaded is not None:
            if not isinstance(loaded, Registry):
                raise TypeError(f"pipeline plugin {plugin} must return Registry or None")
            registry = loaded

    spec = PipelineSpec.from_file(args.spec)
    result = PipelineRunner(registry).run(
        spec,
        args.workspace,
        output_dir=args.output_dir,
        publish=not args.no_publish,
        run_id=args.run_id,
    )
    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2, default=str))
    return 0 if result.success else 1


def _cmd_md2docx(args: argparse.Namespace) -> int:
    from ..adapters.legacy import run_md2docx

    return run_md2docx(args)


def _load_split(path: Path | None) -> list[dict] | None:
    return json.loads(path.read_text(encoding="utf-8-sig")) if path else None


def _cmd_contract(args: argparse.Namespace) -> int:
    from ..markdown import build_docx_contract

    _preflight_docx(args.template, args.baseline)
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
        baseline_path=args.baseline,
        redline_output=args.redline_output,
        render_engine=args.render,
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

    _preflight_docx(args.input)
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

    _preflight_docx(args.input)
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


def _cmd_roundtrip(args: argparse.Namespace) -> int:
    _preflight_docx(args.input)
    if args.section_map is None:
        from ..markdown import roundtrip_source_docx

        result = roundtrip_source_docx(
            args.input,
            args.workdir,
            args.output,
            edited=args.edited,
            baseline=args.baseline,
            redline=args.redline,
            block_map=args.block_map,
            allow_delete=args.allow_delete,
            force=args.force,
            render_engine=args.render,
            report=args.report,
        )
        print(json.dumps(result, ensure_ascii=False, indent=1, default=str))
        return 0
    from ..docxmd import roundtrip

    report = roundtrip(
        args.input,
        args.workdir,
        section_map=args.section_map,
        track_changes=args.track_changes,
        baseline=args.baseline,
        redline=args.redline,
        output=args.output,
        force=args.force,
    )
    print(json.dumps(report, indent=1, ensure_ascii=False, default=str))
    return 0 if report.get("ok") else 1


def _cmd_apply_delta(args: argparse.Namespace) -> int:
    from ..markdown.source_edit import apply_markdown_delta

    _preflight_docx(args.source)
    summary = apply_markdown_delta(
        args.source,
        args.baseline,
        args.edited,
        args.output,
        overwrite=args.force,
        allow_delete=args.allow_delete,
        block_map=args.block_map,
    )
    print("markdown delta: " + ", ".join(f"{key}={value}" for key, value in summary.items()))
    print(args.output)
    return 0


def _cmd_export_block_map(args: argparse.Namespace) -> int:
    from ..markdown.source_edit import export_block_map

    _preflight_docx(args.source)
    if args.output.exists() and not args.force:
        raise FileExistsError(args.output)
    export_block_map(args.source, args.output)
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
    _preflight_docx(baseline, current)
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
        allow_block_replace=args.allow_block_replace,
    )
    print("tracked revisions: " + ", ".join(f"{name}={count}" for name, count in summary.items()))
    if args.report_path is not None:
        print(args.report_path)
    print(args.output)
    return 0


def _cmd_docx2md(args: argparse.Namespace) -> int:
    from ..markdown import docx_to_markdown
    from ..output import validate_output_path

    _preflight_docx(args.input)
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
    from ..adapters.legacy import run_tex2docx

    return run_tex2docx(args)


def _cmd_convert(args: argparse.Namespace) -> int:
    from ..api import convert

    result = convert(
        args.source,
        args.output,
        fidelity=args.fidelity,
        loss_policy=args.loss_policy,
        style=args.style,
        force=args.force,
        title=args.title,
        si=args.si,
        bib=args.bib,
        target=args.target,
    )
    print(result.output)
    print(result.execution_mode)
    return 0


def _cmd_restyle(args: argparse.Namespace) -> int:
    from ..api import restyle

    result = restyle(args.source, args.output, profile=args.profile, loss_policy=args.loss_policy, force=args.force)
    print(result.output)
    print(result.execution_mode)
    return 0


def _cmd_audit(args: argparse.Namespace) -> int:
    from ..api import audit

    report = audit(args.docx)
    print(json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
    return 0 if report.ok else 1


def _cmd_export(args: argparse.Namespace) -> int:
    if args.fidelity == "preserve":
        export_args = argparse.Namespace(
            input=args.input,
            bundle=args.bundle,
            doc_id=args.doc_id,
            split=None,
            bibliography="references.json",
            bib_style=None,
            no_citations=False,
            force=args.force,
        )
        return _cmd_md_export(export_args)
    from ..markdown import docx_to_markdown

    _preflight_docx(args.input)
    output = args.bundle / f"{args.doc_id}.md"
    docx_to_markdown(args.input, output=output, force=args.force)
    print(output)
    return 0


def _cmd_build(args: argparse.Namespace) -> int:
    if args.fidelity == "preserve":
        return _cmd_md_build(args)
    from ..adapters.legacy import md2docx_namespace, run_md2docx

    source = args.bundle / f"{args.doc_id}.md"
    if not source.is_file():
        raise ValueError("semantic build expects a Markdown file named after the document id")
    run_md2docx(md2docx_namespace(source, args.output, force=args.force))
    return 0


def _cmd_docx2tex(args: argparse.Namespace) -> int:
    from ..tex import docx_to_tex
    from ..output import validate_output_path

    _preflight_docx(args.input)
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

    _preflight_docx(args.docx)
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
