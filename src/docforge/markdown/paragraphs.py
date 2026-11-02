"""Generic Markdown paragraphs with native Word math."""

from __future__ import annotations

import re
import subprocess
import tempfile
from pathlib import Path
from typing import Sequence
from zipfile import ZipFile

from lxml import etree


def _portable_math(line: str) -> str:
    """Use equivalent math spellings that render in both Word and LibreOffice."""
    line = line.replace("\\setminus", "\\smallsetminus")
    line = re.sub(r"\\lvert\s*(.*?)\s*\\rvert", r"\\left| \1 \\right|", line)
    return re.sub(r"(?<![\w$])([A-Za-z][\w.]*) \$\s*([<>])", r"$\\text{\1} \2", line)


def render_omml_paragraphs(lines: Sequence[str]) -> list[etree._Element]:
    """Convert independent Markdown lines into Word paragraphs with OMML math.

    Each input line becomes one paragraph; text styles and inline math remain
    editable. Pandoc is an internal math conversion dependency of docforge.
    """
    if not lines:
        return []
    with tempfile.TemporaryDirectory(prefix="docforge-paragraphs-") as directory:
        source = Path(directory) / "source.md"
        output = Path(directory) / "output.docx"
        source.write_text("\n\n".join(_portable_math(line) for line in lines) + "\n", encoding="utf-8")
        try:
            subprocess.run(
                ["pandoc", "-f", "markdown", str(source), "-o", str(output)],
                check=True, capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
        except FileNotFoundError as exc:
            raise RuntimeError("pandoc is required for native Word math; install it on PATH") from exc
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"Word math conversion failed: {exc.stderr.strip()}") from exc
        with ZipFile(output) as package:
            root = etree.fromstring(package.read("word/document.xml"))
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = root.xpath("./w:body/w:p", namespaces=namespace)
    if len(paragraphs) != len(lines):
        raise ValueError(f"Expected {len(lines)} paragraphs but generated {len(paragraphs)}")
    return paragraphs
