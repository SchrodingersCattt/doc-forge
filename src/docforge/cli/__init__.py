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
    md.add_argument("--skip-images", action="store_true", help="Skip Markdown body images")
    md.add_argument("--force", action="store_true", help="Overwrite an existing output file")

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

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "md2docx":
            return _cmd_md2docx(args)
        if args.command == "redline":
            return _cmd_redline(args)
        if args.command == "tex2docx":
            return _cmd_tex2docx(args)
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

    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    blocks = [
        block
        for path in args.inputs
        for block in parse_markdown(path, strip_comments=not args.keep_comments)
    ]
    if args.skip_images:
        blocks = [block for block in blocks if block.kind != "image"]
    doc = render_blocks_to_doc(blocks, title=args.title)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.output))
    convert_unicode_scripts_in_docx(args.output)
    print(args.output)
    return 0


def _cmd_redline(args: argparse.Namespace) -> int:
    from ..docxdiff import create_tracked_docx

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


def _cmd_tex2docx(args: argparse.Namespace) -> int:
    from ..tex import convert_files

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
    print(f"OK {args.docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())