"""Compile mathematical LaTeX to native Word OMML with Pandoc.

Pandoc is used only as a math compiler.  Docforge remains responsible for
document structure, templates, equation numbering, and paragraph layout.
"""

from __future__ import annotations

import copy
import re
import subprocess
import tempfile
import zipfile
from functools import lru_cache
from pathlib import Path

from lxml import etree

MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NAMESPACES = {"m": MATH_NAMESPACE, "w": WORD_NAMESPACE}
PANDOC_VERSION = "2.9.2.1"
_NAMED_OPERATORS = ("clip", "tanh")


class MathConversionError(RuntimeError):
    """Raised when mathematical LaTeX cannot be converted without loss."""


@lru_cache(maxsize=1)
def _pandoc_version() -> str:
    try:
        completed = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            timeout=10,
        )
    except FileNotFoundError as exc:
        raise MathConversionError(
            "pandoc is required for DOCX math conversion; install Pandoc "
            f"{PANDOC_VERSION} and make `pandoc` available on PATH"
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise MathConversionError("pandoc version check timed out") from exc
    if completed.returncode:
        diagnostic = completed.stderr.strip() or completed.stdout.strip()
        raise MathConversionError(f"pandoc version check failed: {diagnostic}")
    first_line = completed.stdout.splitlines()[0] if completed.stdout else ""
    match = re.fullmatch(r"pandoc\s+(\S+)", first_line.strip())
    if not match:
        raise MathConversionError(f"could not determine pandoc version from: {first_line!r}")
    version = match.group(1)
    if version != PANDOC_VERSION:
        raise MathConversionError(
            f"docforge requires pandoc {PANDOC_VERSION} for reproducible OMML; found {version}"
        )
    return version


def _source_excerpt(source: str, limit: int = 100) -> str:
    excerpt = " ".join(source.split())
    return excerpt if len(excerpt) <= limit else excerpt[: limit - 1] + "…"


def _normalize_backend_source(source: str) -> str:
    """Express named functions portably without changing their semantics."""
    for name in _NAMED_OPERATORS:
        source = re.sub(rf"\\{name}(?![A-Za-z])", rf"\\operatorname{{{name}}}", source)
    source = re.sub(
        r"\\(mathbf|boldsymbol|mathcal|mathbb|mathrm)\s+(\\[A-Za-z]+|[A-Za-z])",
        r"\\\1{\2}",
        source,
    )
    # texmath 0.12 rejects nested explicit size modifiers such as
    # ``\Bigl(\bigl[...\bigr];...\Bigr)`` although the delimiters themselves
    # are ordinary and retain the same mathematical meaning without them.
    source = re.sub(r"\\(?:big|Big|bigg|Bigg)[lr]", "", source)
    # Scripts inside \mathrm are mathematical structure, not roman text.
    source = re.sub(r"\\mathrm\{([A-Za-z]+)\\,\\AA\^\{([^{}]+)\}\}", r"\\mathrm{\1}\,\\AA^{\2}", source)
    source = re.sub(
        r"\\(bar|hat|tilde|vec)\{\\mathbf\{([^{}]+)\}\}(_\{[^{}]+\}|_[A-Za-z0-9])",
        r"{\\\1{\\mathbf{\2}}}\3",
        source,
    )
    source = re.sub(r"\\AA(?![A-Za-z])", r"\\text{Å}", source)
    return source


def latex_to_omml(source: str) -> etree._Element:
    """Return one detached ``m:oMathPara`` compiled from display math.

    Conversion is deliberately strict: malformed input, Pandoc diagnostics,
    missing math output, or leaked LaTeX command names fail the build instead
    of producing a plausible-looking but incorrect equation.
    """
    source = source.strip()
    if not source:
        raise MathConversionError("display equation is empty")
    _pandoc_version()
    backend_source = _normalize_backend_source(source)
    markdown = f"$$\n{backend_source}\n$$\n"
    with tempfile.TemporaryDirectory(prefix="docforge-math-") as directory:
        root = Path(directory)
        input_path = root / "equation.md"
        output_path = root / "equation.docx"
        input_path.write_text(markdown, encoding="utf-8")
        command = [
            "pandoc",
            str(input_path),
            "--from=markdown",
            "--to=docx",
            "--output",
            str(output_path),
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
                timeout=30,
            )
        except subprocess.TimeoutExpired as exc:
            raise MathConversionError(
                f"pandoc math conversion timed out: {_source_excerpt(source)}"
            ) from exc
        diagnostic = "\n".join(
            part.strip() for part in (completed.stderr, completed.stdout) if part.strip()
        )
        if completed.returncode or diagnostic:
            detail = diagnostic or f"exit code {completed.returncode}"
            raise MathConversionError(
                f"pandoc could not convert display equation `{_source_excerpt(source)}`: {detail}"
            )
        try:
            with zipfile.ZipFile(output_path) as package:
                document_xml = package.read("word/document.xml")
        except (OSError, KeyError, zipfile.BadZipFile) as exc:
            raise MathConversionError("pandoc produced an invalid DOCX math result") from exc

    document = etree.fromstring(document_xml)
    math_paragraphs = document.xpath("//m:oMathPara", namespaces=NAMESPACES)
    if len(math_paragraphs) != 1:
        raise MathConversionError(
            f"expected one display equation from pandoc, found {len(math_paragraphs)}: "
            f"{_source_excerpt(source)}"
        )
    result = copy.deepcopy(math_paragraphs[0])
    if not result.xpath("./m:oMath", namespaces=NAMESPACES):
        raise MathConversionError("pandoc returned an empty OMML display equation")
    return result