"""Legacy md2docx and tex2docx orchestration.

The CLI and the new convert() entry point both call these functions, so the
historical command behavior stays on one implementation.
"""

from __future__ import annotations

import argparse
from pathlib import Path


def run_md2docx(args: argparse.Namespace, *, echo: bool = True) -> int:
    from ..markdown import parse_markdown, render_blocks_to_doc, convert_unicode_scripts_in_docx
    from ..markdown import reuse_unchanged_roundtrip_source
    from ..output import _detect_format, validate_output_path

    validate_output_path(args.output)
    if args.template is not None:
        _detect_format(args.template)
    if args.output.exists() and not args.force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {args.output}")
    if args.roundtrip_manifest is not None:
        if reuse_unchanged_roundtrip_source(
            args.roundtrip_manifest, args.output, force=args.force, validate_format=True,
        ):
            _echo(echo, f"roundtrip exact source reused: {args.output}")
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
            command=getattr(args, "_command_argv", ["docforge", "md2docx"]),
        )
        _echo(echo, result.output)
        _echo(echo, manifest)
        _echo(echo, checksum)
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
    _echo(echo, args.output)
    return 0


def run_tex2docx(args: argparse.Namespace, *, echo: bool = True) -> int:
    from ..output import validate_output_path
    from ..tex import build_validated_docx, convert_files

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
    _echo(echo, args.output)
    return 0


def _echo(enabled: bool, line: object) -> None:
    if enabled:
        print(line)


def md2docx_namespace(source: Path, output: Path, *, force: bool = False, title: str | None = None) -> argparse.Namespace:
    """Defaults that match ``docforge md2docx SOURCE -o OUTPUT`` without a template."""

    return argparse.Namespace(
        inputs=[Path(source)],
        output=Path(output),
        keep_comments=False,
        title=title,
        template=None,
        metadata=None,
        bibliography=None,
        citation_base=None,
        citation_format="template",
        columns="template",
        figure_span="column",
        font_family=None,
        east_asia_font=None,
        body_font_size=None,
        abstract_font_size=None,
        caption_font_size=None,
        reference_font_size=None,
        style_profile="template",
        line_numbers="template",
        heading_before=None,
        native_toc=False,
        restart_heading_numbering=False,
        body_first_line_chars=None,
        page_break_before_h1=False,
        numbering_prefix="",
        bibliography_scope="auto",
        citation_numbering="first-citation",
        bibliography_profile="markdown",
        omit_metadata_back_matter=False,
        no_title=False,
        skip_images=False,
        force=force,
        roundtrip_manifest=None,
        _command_argv=["docforge", "convert", str(source), str(output)],
    )


def tex2docx_namespace(
    main: Path,
    output: Path,
    *,
    si: Path | None = None,
    bib: Path | None = None,
    target: str = "main",
    force: bool = False,
) -> argparse.Namespace:
    return argparse.Namespace(
        main=Path(main),
        si=Path(si) if si else None,
        bib=Path(bib) if bib else None,
        target=target,
        output=Path(output),
        force=force,
    )


__all__ = [
    "md2docx_namespace", "run_md2docx", "run_tex2docx", "tex2docx_namespace",
]
