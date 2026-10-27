"""Layout render gates.

A missing engine is ``skipped``. ``skipped`` is not ``passed``. Word visual
parity is ``passed`` only after Word itself writes a PDF. LibreOffice can
pass its own render gate without that claim.

Run the gate where the engine is installed::

    pytest tests/test_render_gates.py -q
    pytest tests/test_contract.py::test_contract_runs_redline_render_and_consistent_manifest -q
"""

from __future__ import annotations

from pathlib import Path
from typing import Mapping


def counts_as_pass(report: Mapping[str, object]) -> bool:
    """Return whether a layout report is acceptance evidence.

    A skip is not a pass. Callers must not treat a missing engine as success.
    """

    return report.get("status") == "passed"


def layout_gate(docx: Path, out_dir: Path, engine: str = "auto") -> dict[str, object]:
    """Render one DOCX and report the engine that actually ran.

    ``engine`` is ``auto``, ``word``, or ``libreoffice``. ``auto`` prefers
    Word, then LibreOffice. Neither substitution marks Word visual parity as
    passed.
    """

    from ..docxmd.verify import _soffice, _word_available, render_pdf

    if engine not in {"auto", "word", "libreoffice"}:
        raise ValueError("engine must be auto, word, or libreoffice")
    word = _word_available()
    soffice = _soffice() is not None
    if engine == "word":
        resolved = "word" if word else None
    elif engine == "libreoffice":
        resolved = "libreoffice" if soffice else None
    else:
        resolved = "word" if word else "libreoffice" if soffice else None
    report: dict[str, object] = {
        "requested_engine": engine,
        "resolved_engine": resolved,
        "status": "skipped",
        "visual_parity": "not-passed",
        "reason": _skip_reason(engine, word=word, soffice=soffice),
        "pdf_bytes": None,
        "page_count": None,
    }
    if resolved is None:
        return report
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        pdf = render_pdf(Path(docx), out_dir, resolved)
    except Exception as exc:
        report["status"] = "failed"
        report["reason"] = f"{resolved} render failed: {type(exc).__name__}"
        return report
    if not pdf.is_file() or pdf.stat().st_size <= 0:
        report["status"] = "failed"
        report["reason"] = f"{resolved} produced no PDF"
        return report
    report["status"] = "passed"
    report["pdf_bytes"] = pdf.stat().st_size
    report["page_count"] = _page_count(pdf)
    report["reason"] = f"{resolved} wrote a PDF"
    report["visual_parity"] = "passed" if resolved == "word" else "not-passed"
    return report


def _skip_reason(engine: str, *, word: bool, soffice: bool) -> str:
    if engine == "word" and not word:
        return "Word is unavailable; visual parity was not run. This skip is not a pass."
    if engine == "libreoffice" and not soffice:
        return "LibreOffice (soffice) is unavailable. This skip is not a pass."
    if engine == "auto" and not word and not soffice:
        return "Word and LibreOffice are unavailable. This skip is not a pass."
    return ""


def _page_count(pdf: Path) -> int | None:
    try:
        import fitz
    except ImportError:
        return None
    with fitz.open(pdf) as document:
        return len(document)


__all__ = ["counts_as_pass", "layout_gate"]
