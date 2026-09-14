"""Thin command-line entry points for docforge.

Each subcommand lazily imports its feature module so the minimum install stays
functional even when optional dependencies (matplotlib, openai, google-genai)
are absent.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docforge",
        description=__doc__,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    md = sub.add_parser("md2docx", help="Markdown -> DOCX (block parser + renderer)")
    md.add_argument("inputs", nargs="+", type=Path, help="Markdown files to merge in order")
    md.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX path")
    md.add_argument("--keep-comments", action="store_true", help="Keep Markdown HTML comments as red TODO notes")
    md.add_argument("--title", help="Optional title for a standalone document")
    md.add_argument("--template", type=Path, help="DOCX template to preserve while assembling Markdown")
    md.add_argument("--metadata", type=Path, help="Markdown metadata/front-matter file for template assembly")
    md.add_argument("--bibliography", type=Path, help="Ordered JSON bibliography for citation expansion")
    md.add_argument("--citation-base", type=Path, help="Existing main-manuscript manifest whose citation numbers SI should reuse")
    md.add_argument("--citation-format", choices=("template", "superscript", "bracketed"), default="template", help="Citation rendering policy")
    md.add_argument("--columns", choices=("template", "one", "two"), default="template", help="Body column layout for template assembly")
    md.add_argument("--font", dest="font_family", help="Explicit Latin font override for generated text")
    md.add_argument("--east-asia-font", dest="east_asia_font", help="Explicit East Asian font override for generated text")
    md.add_argument("--style", dest="style_profile", default="template", help="Template style profile or JSON semantic-style map")
    md.add_argument("--line-numbers", choices=("template", "on", "off"), default="template", help="Line-number policy for generated sections")
    md.add_argument("--heading-before", type=float, default=None, help="Explicit heading spacing before in points")
    md.add_argument("--native-toc", action="store_true", help="Insert a native Word table of contents after template front matter")
    md.add_argument("--restart-heading-numbering", action="store_true", help="Restart H2 decimal numbering after each H1")
    md.add_argument("--body-first-line-chars", type=float, default=None, help="Body paragraph first-line indent in character units")
    md.add_argument("--page-break-before-h1", action="store_true", help="Start each level-one Markdown heading on a new page")
    md.add_argument("--numbering-prefix", default="", help="Prefix for figure, table-caption, and equation numbers (e.g. S for SI)")
    md.add_argument("--bibliography-scope", choices=("all", "new-only"), default="all", help="Reference entries to render: all used citations or only citations absent from --citation-base")
    md.add_argument("--omit-metadata-back-matter", action="store_true", help="Omit acknowledgment, author-contribution, and code-availability sections")
    md.add_argument("--no-title", action="store_true", help="Remove level-one Markdown headings; retain the metadata title block")
    md.add_argument("--skip-images", action="store_true", help="Skip Markdown body images")
    md.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    docx_md = sub.add_parser("docx2md", help="DOCX -> GitHub-Flavored Markdown via Pandoc")
    docx_md.add_argument("input", type=Path, help="Input DOCX path")
    docx_md.add_argument("-o", "--output", type=Path, help="Output Markdown path")
    docx_md.add_argument("--split-dir", type=Path, help="Directory for section Markdown files")
    docx_md.add_argument("--section-map", type=Path, help="JSON section map used with --split-dir")
    docx_md.add_argument("--media-dir", type=Path, help="Pandoc extraction root (contains media/)")
    docx_md.add_argument("--track-changes", choices=("accept", "reject", "all"), default="accept")
    docx_md.add_argument("--force", action="store_true", help="Overwrite existing output files")

    md2 = sub.add_parser("redline", help="Create a DOCX with native Word revisions against a reviewed DOCX")
    md2.add_argument("base", type=Path, help="Reviewed DOCX baseline")
    md2.add_argument("current", type=Path, help="Freshly generated DOCX")
    md2.add_argument("-o", "--output", type=Path, required=True, help="Output tracked DOCX path")
    md2.add_argument("--revision-author", default="M.Y.G.", help="Author recorded for revisions")
    md2.add_argument("--force", action="store_true", help="Overwrite an existing output file")

    tex = sub.add_parser("tex2docx", help="LaTeX -> DOCX (main + optional SI)")
    tex.add_argument("main", type=Path, help="main.tex path")
    tex.add_argument("--si", type=Path, help="si.tex path")
    tex.add_argument("--bib", type=Path, help="ref.bib path")
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

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    raw_argv = list(argv) if argv is not None else sys.argv[1:]
    args = parser.parse_args(raw_argv)
    args._command_argv = ["docforge", *raw_argv]
    try:
        if args.command == "md2docx":
            return _cmd_md2docx(args)
        if args.command == "docx2md":
            return _cmd_docx2md(args)
        if args.command == "redline":
            return _cmd_redline(args)
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
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    parser.error(f"Unknown command: {args.command}")
    return 2


def _cmd_md2docx(args: argparse.Namespace) -> int:
    from ..markdown import parse_markdown, render_blocks_to_doc, convert_unicode_scripts_in_docx
    from ..output import validate_output_path

    validate_output_path(args.output)
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
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
            numbering_prefix=args.numbering_prefix,
            bibliography_scope=args.bibliography_scope,
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
        or args.font_family is not None
        or args.east_asia_font is not None
        or args.style_profile != "template"
        or args.line_numbers != "template"
    ):
        raise ValueError("Metadata, bibliography, style, font, and line-number options require --template")
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
    doc = render_blocks_to_doc(blocks, title=args.title, number_prefix=args.numbering_prefix, heading_before=args.heading_before)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.output))
    convert_unicode_scripts_in_docx(args.output)
    print(args.output)
    return 0


def _cmd_redline(args: argparse.Namespace) -> int:
    from ..docxdiff import create_tracked_docx
    from ..output import validate_output_path

    validate_output_path(args.output)
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    summary = create_tracked_docx(
        args.base,
        args.current,
        args.output,
        author=args.revision_author,
        overwrite=args.force,
    )
    print("tracked revisions: " + ", ".join(f"{name}={count}" for name, count in summary.items()))
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
    from ..tex import convert_files
    from ..output import validate_output_path

    validate_output_path(args.output)
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    convert_files(
        args.main,
        si_path=args.si,
        bib_path=args.bib,
        output=args.output,
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


if __name__ == "__main__":
    raise SystemExit(main())
