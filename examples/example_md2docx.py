#!/usr/bin/env python3
"""Convert a simple Markdown file into a DOCX with docforge (library API).

This is the minimal standalone example: parse a Markdown document with the
block parser, render it into a fresh styled DOCX, and validate the package.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docforge.markdown import parse_markdown, render_blocks_to_doc
from docforge.markdown import verify_navigation_headings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=Path("output.docx"))
    parser.add_argument("--title", default="")
    args = parser.parse_args()

    blocks = parse_markdown(args.input)
    doc = render_blocks_to_doc(blocks, title=args.title)
    doc.save(str(args.output))
    verify_navigation_headings(args.output)
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())