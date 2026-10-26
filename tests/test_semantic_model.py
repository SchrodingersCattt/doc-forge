"""SemanticDocument can be built, checked, and hashed without Word or Pandoc."""

from __future__ import annotations

import subprocess
import sys

from docforge.semantic import BlockNode, DocumentMetadata, SemanticDocument, text_node


def test_round_trip_serialization_and_hash() -> None:
    document = SemanticDocument(
        metadata=DocumentMetadata(title="Sample"),
        blocks=(
            BlockNode("b1", "heading", inlines=(text_node("Sample"),), attrs=(("level", "1"),)),
            BlockNode("b2", "paragraph", inlines=(text_node("Alpha beta"),)),
        ),
    )
    restored = SemanticDocument.from_dict(document.to_dict())
    assert restored == document
    assert restored.semantic_hash() == document.semantic_hash()
    assert restored.validate() == ()
    assert "path" not in restored.projection()["blocks"][0]


def test_duplicate_ids_and_absolute_resource_paths_fail() -> None:
    document = SemanticDocument(blocks=(
        BlockNode("b1", "paragraph", inlines=(text_node("One"),)),
        BlockNode("b1", "paragraph", inlines=(text_node("Two"),)),
    ))
    assert any("duplicate" in item for item in document.validate())
    from docforge.semantic import ResourceDescriptor
    try:
        ResourceDescriptor("r", "image/png", path="/tmp/secret.png")
    except ValueError as exc:
        assert "relative" in str(exc)
    else:
        raise AssertionError("absolute resource path was accepted")


def test_model_import_does_not_load_word() -> None:
    script = "import docforge.semantic.model, sys; raise SystemExit('docx' in sys.modules)"
    completed = subprocess.run([sys.executable, "-c", script], check=False, capture_output=True, text=True)
    assert completed.returncode == 0, completed.stderr
