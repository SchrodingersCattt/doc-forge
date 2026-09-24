from pathlib import Path

import pytest

from docforge.bibliography import CitationResolver, format_entry, load_bib, load_json


def test_json_and_bib_loaders_preserve_entry_type_and_fields(tmp_path: Path) -> None:
    json_path = tmp_path / "refs.json"
    json_path.write_text(
        '{"paper": {"authors": ["A"], "year": 2024, "journal": "J", "doi": "10/x"}, "raw": "A raw record"}',
        encoding="utf-8",
    )
    entries = load_json(json_path)
    assert entries["paper"].get("year") == 2024
    assert entries["raw"].get("raw") == "A raw record"
    bib_path = tmp_path / "refs.bib"
    bib_path.write_text("@article{paper,\n author={A},\n year={2024}\n}\n", encoding="utf-8")
    assert load_bib(bib_path)["paper"].entry_type == "article"


def test_resolver_supports_inherited_and_source_order_numbering() -> None:
    entries = {key: {"title": key} for key in ("b", "a", "c")}
    resolver = CitationResolver(entries, strict=True)
    used, mapping = resolver.plan(["a", "b"], numbering="source-order")
    assert used == ("b", "a")
    assert mapping == {"b": 1, "a": 2}
    with pytest.raises(ValueError, match="Unknown citation"):
        resolver.resolve("missing")


def test_formatter_profiles_are_explicit() -> None:
    entry = load_json(Path(__file__).parent / "fixtures" / "bibliography.json")["paper"]
    assert "Journal" in format_entry(entry, profile="plain")
    assert "**2024**" in format_entry(entry, profile="markdown")

