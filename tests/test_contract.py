from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZipFile

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


def test_contract_accepts_dotx_and_publishes_final_output_descriptor(tmp_path: Path) -> None:
    document = Document()
    for name in (
        "BBAuthorName", "FACorrespondingAuthorFootnote", "BDAbstractTitle",
        "BDAbstract", "TA_Main_Text", "Heading 1", "Heading 2", "Heading 3",
        "Caption", "TFReferencesSection", "EndNote Bibliography",
    ):
        if name not in [style.name for style in document.styles]:
            document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    document.add_paragraph("Title", style="BBAuthorName")
    document.add_paragraph("Body", style="TA_Main_Text")
    docx = tmp_path / "template.docx"
    document.save(docx)
    dotx = tmp_path / "template.dotx"
    with ZipFile(docx) as source, ZipFile(dotx, "w") as target:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == "[Content_Types].xml":
                payload = payload.replace(
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                )
            target.writestr(info, payload)

    source = tmp_path / "body.md"
    source.write_text("Body paragraph.\n", encoding="utf-8")
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    result = build_docx_contract(
        [source], template_path=dotx, metadata_path=metadata, output=output, render=False,
    )
    assert result.output == output
    manifest = output.with_suffix(".manifest.json")
    assert manifest.exists()
    assert json.loads(manifest.read_text(encoding="utf-8"))["output"]["path"] == str(output.resolve())


def test_contract_rejects_template_output_collision(tmp_path: Path) -> None:
    document = Document()
    for name in ("BBAuthorName", "FACorrespondingAuthorFootnote", "BDAbstract", "TA_Main_Text", "EndNote Bibliography"):
        document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    document.add_paragraph("Title", style="BBAuthorName")
    document.add_paragraph("Body", style="TA_Main_Text")
    template = tmp_path / "template.docx"
    document.save(template)
    source = tmp_path / "body.md"
    source.write_text("Body paragraph.\n", encoding="utf-8")
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")

    with pytest.raises(ValueError, match="differ from the template"):
        build_docx_contract(
            [source], template_path=template, metadata_path=metadata, output=template, force=True
        )


def _contract_fixture(tmp_path: Path) -> tuple[Path, Path, Path, Path]:
    document = Document()
    for name in (
        "BBAuthorName", "FACorrespondingAuthorFootnote", "BDAbstractTitle",
        "BDAbstract", "TA_Main_Text", "Heading 1", "Heading 2", "Heading 3",
        "Caption", "TFReferencesSection", "EndNote Bibliography",
    ):
        if name not in [style.name for style in document.styles]:
            document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    document.add_paragraph("Title", style="BBAuthorName")
    document.add_paragraph("Body", style="TA_Main_Text")
    template = tmp_path / "template.docx"
    document.save(template)
    source = tmp_path / "body.md"
    source.write_text("Body paragraph with an edit.\n", encoding="utf-8")
    metadata = tmp_path / "metadata.md"
    metadata.write_text("# TITLE\n\nA title\n", encoding="utf-8")
    baseline = tmp_path / "baseline.docx"
    baseline_doc = Document()
    baseline_doc.add_paragraph("Body paragraph before the edit.")
    baseline_doc.save(baseline)
    return template, source, metadata, baseline


def test_contract_runs_redline_render_and_consistent_manifest(tmp_path: Path) -> None:
    from docforge.docxmd.verify import _soffice

    if _soffice() is None:
        pytest.skip("LibreOffice render gate is not executable in this environment; it stays open")
    template, source, metadata, baseline = _contract_fixture(tmp_path)
    output = tmp_path / "output.docx"
    result = build_docx_contract(
        [source],
        template_path=template,
        metadata_path=metadata,
        output=output,
        baseline_path=baseline,
        render_engine="libreoffice",
    )
    payload = json.loads(result.manifest.read_text(encoding="utf-8"))
    contract = payload["contract"]
    assert contract["accepted_view"]["equivalent"] is True
    assert contract["redline"]["enabled"] is True
    assert contract["redline"]["summary"]["changed"] == 1
    assert contract["render"]["enabled"] is True
    assert contract["render"]["documents"]
    assert payload["output"]["sha256"] == __import__("hashlib").sha256(output.read_bytes()).hexdigest()
    assert result.redline is not None and result.redline.exists()
    first_manifest = result.manifest.read_bytes()
    first_output = output.read_bytes()
    build_docx_contract(
        [source],
        template_path=template,
        metadata_path=metadata,
        output=output,
        baseline_path=baseline,
        render_engine="libreoffice",
        force=True,
    )
    assert output.read_bytes() == first_output
    assert result.manifest.read_bytes() == first_manifest


def test_contract_install_rolls_back_existing_bundle_on_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    template, source, metadata, _baseline = _contract_fixture(tmp_path)
    output = tmp_path / "output.docx"
    manifest = output.with_suffix(".manifest.json")
    checksum = output.with_suffix(".sha256")
    old = {path: f"old-{path.name}".encode() for path in (output, manifest, checksum)}
    for path, payload in old.items():
        path.write_bytes(payload)
    import docforge.markdown.contract as contract_module

    original_replace = contract_module.os.replace
    failures = {"remaining": 1}

    def fail_manifest(source_path, destination):
        if Path(destination) == manifest and failures["remaining"]:
            failures["remaining"] -= 1
            raise OSError("simulated install failure")
        return original_replace(source_path, destination)

    monkeypatch.setattr(contract_module.os, "replace", fail_manifest)
    with pytest.raises(RuntimeError, match="rollback restored"):
        build_docx_contract(
            [source],
            template_path=template,
            metadata_path=metadata,
            output=output,
            render=False,
            force=True,
        )
    assert {path: path.read_bytes() for path in old} == old
