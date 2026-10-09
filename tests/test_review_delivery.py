from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document

from docforge.cli import build_parser
from docforge.markdown import deliver


def test_redline_accept_baseline_named_arguments_parse() -> None:
    args = build_parser().parse_args(
        [
            "redline",
            "--baseline",
            "reviewed.docx",
            "--current",
            "current.docx",
            "--output",
            "tracked.docx",
            "--accept-baseline",
            "--revision-author",
            "Reviewer",
            "--report",
            "report.json",
        ]
    )
    assert args.baseline == Path("reviewed.docx")
    assert args.current_option == Path("current.docx")
    assert args.accept_baseline is True
    assert args.revision_author == "Reviewer"
    assert args.report_path == Path("report.json")


def test_deliver_accept_revisions_fails_closed_without_review_input(tmp_path: Path) -> None:
    profile = {
        "delivery_dir": "delivery",
        "article": {"inputs": ["body.md"], "template": "template.docx"},
    }
    (tmp_path / "body.md").write_text("Body.\n", encoding="utf-8")
    document = Document()
    document.add_paragraph("Template")
    document.save(tmp_path / "template.docx")
    path = tmp_path / "profile.json"
    path.write_text(json.dumps(profile), encoding="utf-8")
    with pytest.raises(ValueError, match="reviewed_docx"):
        deliver(path, accept_revisions=True)
