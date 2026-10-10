from dataclasses import replace
from pathlib import Path
from types import MappingProxyType

import pytest

from docforge.bibliography import (
    BibliographyEntry,
    CitationResolver,
    format_entry,
    load_bib,
    load_json,
    parse_bib_text,
)


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
    bib_entry = load_bib(bib_path)["paper"]
    assert bib_entry.entry_type == "article"
    assert bib_entry.get("year") == 2024


def test_bibtex_year_is_accepted_by_markdown_formatter(tmp_path: Path) -> None:
    bib_path = tmp_path / "refs.bib"
    bib_path.write_text(
        "@article{paper, author={A}, title={A title}, journal={J}, year={2024}}\n",
        encoding="utf-8",
    )
    assert "**2024**" in format_entry(load_bib(bib_path)["paper"], profile="markdown")


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


@pytest.mark.parametrize("profile", ["plain", "markdown"])
def test_formatter_does_not_double_terminal_period(profile: str) -> None:
    entry = load_json(Path(__file__).parent / "fixtures" / "bibliography.json")["paper"]
    fields = {**entry.fields, "authors": ["Doe, Jane", "Roe, Alex B. C."]}
    entry = replace(entry, fields=MappingProxyType(fields))
    text = format_entry(entry, profile=profile)
    assert "Roe, Alex B. C. " in text
    assert ".." not in text



def test_bib_loader_balances_nested_braces_and_non_journal_venues(tmp_path: Path) -> None:
    path = tmp_path / "nested.bib"
    path.write_text(
        """@inproceedings{nested-key,
 author={Doe, Jane},
 title={A {Nested} title, with commas},
 booktitle={Proceedings of the {ACM} Symposium},
 year={2024}
}
""",
        encoding="utf-8",
    )
    entry = load_bib(path)["nested-key"]
    assert entry.get("title") == "A {Nested} title, with commas"
    assert entry.get("booktitle") == "Proceedings of the {ACM} Symposium"
    assert "Proceedings" in format_entry(entry, profile="plain")


def test_bib_loader_concatenates_hash_atoms_without_truncating_fields(tmp_path: Path) -> None:
    path = tmp_path / "concat.bib"
    path.write_text(
        """@article{concat,
 title = {Part A} # "Part B" # { C# },
 journal = {Journal},
 year = {2024}
}
""",
        encoding="utf-8",
    )
    entry = load_bib(path)["concat"]
    assert entry.get("title") == "Part APart B C#"
    assert entry.get("journal") == "Journal"
    assert entry.get("year") == 2024


def test_bib_loader_preserves_hash_inside_a_single_atom(tmp_path: Path) -> None:
    path = tmp_path / "literal-hash.bib"
    path.write_text("@misc{hash, title = {C# guide}}\n", encoding="utf-8")
    assert load_bib(path)["hash"].get("title") == "C# guide"


def test_bib_loader_resolves_string_macros_in_hash_concatenations() -> None:
    entries = parse_bib_text(
        '@string{journal = "Journal of"}\n'
        '@article{macro, title = "A" # " title", journal = journal # " Letters", year = {2024}}\n'
    )
    assert set(entries) == {"macro"}
    assert entries["macro"].get("title") == "A title"
    assert entries["macro"].get("journal") == "Journal of Letters"


def test_markdown_formatter_accepts_publisher_venue() -> None:
    entry = {"authors": ["Doe, Jane"], "title": "A book", "year": 2024, "publisher": "Press"}
    assert "*Press*" in format_entry(BibliographyEntry.from_mapping("book", entry), profile="markdown")
