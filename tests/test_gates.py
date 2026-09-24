import json
from pathlib import Path

from docforge.gates import audit_references, check_abbreviations, load_abbreviation_whitelist


def test_abbreviation_whitelist_supports_project_overrides(tmp_path: Path) -> None:
    config = tmp_path / "gates.json"
    config.write_text(json.dumps({"abbr-whitelist": ["MLIP"], "projects": {"demo": {"abbr-whitelist": ["OOD"]}}}), encoding="utf-8")
    assert {"MLIP", "OOD"}.issubset(load_abbreviation_whitelist(config, project="demo"))
    source = tmp_path / "paper.tex"
    source.write_text("A model (MLIP) and an OOD split use XYZ.", encoding="utf-8")
    findings = check_abbreviations([source], whitelist=load_abbreviation_whitelist(config, project="demo"))
    assert [item.acronym for item in findings] == ["XYZ"]


def test_reference_gate_is_project_neutral(tmp_path: Path) -> None:
    source = tmp_path / "paper.tex"
    source.write_text(r"\begin{document}\begin{figure}\label{fig:one}\end{figure} See Fig.~\ref{fig:one}.\end{document}", encoding="utf-8")
    issues = audit_references(tmp_path, {"documents": {"paper.tex": {}}, "main_document": "paper.tex"})
    assert issues == []
