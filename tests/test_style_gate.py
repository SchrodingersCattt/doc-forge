from pathlib import Path

from docforge.gates import baseline_payload, check_style, unresolved


def test_style_gate_fingerprint_can_be_baselined(tmp_path: Path) -> None:
    source = tmp_path / "paper.tex"
    source.write_text("A 1 - 2 range.\n", encoding="utf-8")
    findings = check_style([source])
    assert len(findings) == 1
    assert unresolved(findings, baseline_payload(findings)) == []
