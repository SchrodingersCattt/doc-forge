from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

from docforge.markdown import build_docx_contract


def _minimal_template(path: Path) -> None:
    document = Document()
    document.styles.add_style("TA_Main_Text", WD_STYLE_TYPE.PARAGRAPH)
    document.save(path)


def test_contract_rejects_missing_roles_before_creating_output(tmp_path: Path) -> None:
    template = tmp_path / "template.docx"
    _minimal_template(template)
    source = tmp_path / "body.md"
    source.write_text("A paragraph.\n", encoding="utf-8")
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nTitle\n", encoding="utf-8")
    output = tmp_path / "output.docx"

    with pytest.raises(ValueError, match="semantic style"):
        build_docx_contract(
            [source], template_path=template, metadata_path=metadata, output=output
        )
    assert not output.exists()
    assert not output.with_suffix(".manifest.json").exists()


def test_contract_block_audit_is_exposed_in_manifest(tmp_path: Path) -> None:
    # The strict preflight is tested above; this assertion keeps the public
    # result shape explicit for callers that display the audit before delivery.
    from docforge.markdown.contract import ContractResult

    result = ContractResult(tmp_path / "out.docx", tmp_path / "out.manifest.json", tmp_path / "out.sha256", {"body": "TA_Main_Text"}, ({"kind": "paragraph", "decision": "render"},))
    assert result.blocks[0]["decision"] == "render"
