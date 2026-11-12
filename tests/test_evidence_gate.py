import hashlib
import json
from pathlib import Path

from docforge.gates.evidence import read_sources, repair_unattested
from docforge.gates.figures import check_figure_freshness, file_sha256


def test_unattested_number_and_doi_are_replaced(tmp_path: Path) -> None:
    source = tmp_path / "note.md"
    source.write_text("密度为 0.25 g/cm3。\n", encoding="utf-8")
    draft = "密度为 0.25 g/cm3。虚构值为 99.9%。DOI 10.1000/synthetic.1。"
    repaired, notes = repair_unattested(draft, read_sources([source]))
    assert "0.25" in repaired
    assert "99.9" not in repaired
    assert "10.1000/synthetic.1" not in repaired
    assert repaired.count("[未提供]") == 2
    assert any("99.9" in note for note in notes)
    assert any("10.1000/synthetic.1" in note for note in notes)


def test_attested_doi_is_kept(tmp_path: Path) -> None:
    source = tmp_path / "refs.md"
    source.write_text("See 10.1000/synthetic.1 for the 0.25 g/cm3 density.\n", encoding="utf-8")
    draft = "The density is 0.25 g/cm3 (10.1000/synthetic.1)."
    repaired, notes = repair_unattested(draft, read_sources([source]))
    assert repaired == draft
    assert notes == []


def test_figure_manifest_detects_a_stale_file(tmp_path: Path) -> None:
    figure = tmp_path / "deliverable.svg"
    figure.write_text("<svg>0.25</svg>", encoding="utf-8")
    manifest = {"files": [{"path": "deliverable.svg", "sha256": file_sha256(figure)}]}
    assert check_figure_freshness(manifest, tmp_path) == []
    figure.write_text("<svg>changed</svg>", encoding="utf-8")
    assert check_figure_freshness(manifest, tmp_path) == ["stale deliverable.svg"]
    missing = {"files": [{"path": "absent.svg", "sha256": hashlib.sha256(b"").hexdigest()}]}
    assert check_figure_freshness(missing, tmp_path) == ["missing absent.svg"]
    (tmp_path / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
