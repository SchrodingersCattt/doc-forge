"""The controlled bench harness records time, RSS, and package size."""

from __future__ import annotations

from pathlib import Path

from docforge.bench import median_rows, run_case


def test_small_markdown_case_records_time_rss_and_size(tmp_path: Path) -> None:
    row = run_case("md-docx", "small", tmp_path)
    assert row["op"] == "md-docx"
    assert row["kind"] == "small"
    assert float(row["seconds"]) > 0
    assert int(row["rss_kb"]) > 0
    assert int(row["bytes"]) > 1000
    assert (tmp_path / "small.docx").is_file()


def test_median_rows_uses_the_middle_run() -> None:
    rows = [
        {"op": "md-docx", "kind": "small", "seconds": 1.0, "rss_kb": 10, "bytes": 3},
        {"op": "md-docx", "kind": "small", "seconds": 3.0, "rss_kb": 30, "bytes": 5},
        {"op": "md-docx", "kind": "small", "seconds": 2.0, "rss_kb": 20, "bytes": 4},
    ]
    summary = median_rows(rows)
    assert summary == [{
        "op": "md-docx",
        "kind": "small",
        "seconds": 2.0,
        "rss_kb": 20,
        "bytes": 4,
        "runs": 3,
    }]
