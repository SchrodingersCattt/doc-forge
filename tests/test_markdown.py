"""Tests for the docforge markdown block parser and OMML math builders."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pytest

from docforge.markdown.blocks import Block
from docforge.markdown.launcher import normalize_typography, parse_markdown, render_blocks_to_doc


class BlockTests(unittest.TestCase):
    def test_block_as_dict(self) -> None:
        block = Block("table", rows=(("a", "b"), ("1", "2")))
        payload = block.as_dict()
        self.assertEqual(payload["kind"], "table")
        self.assertEqual(payload["rows"], [["a", "b"], ["1", "2"]])

    def test_block_defaults(self) -> None:
        block = Block("heading", "Title", level=1)
        self.assertEqual(block.rows, ())
        self.assertEqual(block.path, "")

    def test_bang_comments_are_ignored_and_image_options_are_parsed(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text(
                "! internal Chinese blueprint\n\nVisible paragraph.\n\n"
                "![Figure 1|columns=single](figure.png)\n\n"
                "Figure 1. Caption.\n",
                encoding="utf-8",
            )
            blocks = parse_markdown(path)
            self.assertEqual([block.kind for block in blocks], ["paragraph", "image", "paragraph"])
            self.assertEqual(blocks[1].options, (("columns", "single"),))
            kept = parse_markdown(path, strip_comments=False)
            self.assertIn("[TODO: internal Chinese blueprint]", [block.text for block in kept])


class OmmlTests(unittest.TestCase):
    def test_math_commands_table(self) -> None:
        from docforge.markdown.omml import MATH_COMMANDS

        self.assertIn(r"\alpha", MATH_COMMANDS)
        self.assertEqual(MATH_COMMANDS[r"\AA"], "Å")

    def test_parse_math_omml_runs(self) -> None:
        from docforge.markdown.omml import parse_math_omml

        elements = parse_math_omml(r"E = m c^2")
        # Whitespace, letters/digits/operators split into runs; the exponent
        # becomes an m:sSup element.
        names = [element.tag.rsplit("}", 1)[-1] for element in elements]
        self.assertTrue(all(name in ("r", "sSup", "sSub", "sSubSup") for name in names))
        self.assertIn("sSup", names)

        texts = "".join(
            node.text or ""
            for element in elements
            for node in element.iter()
            if node.tag.endswith("t") and node.text
        )
        self.assertEqual(texts.replace(" ", ""), "E=mc2")

    def test_parse_math_frac(self) -> None:
        from docforge.markdown.omml import parse_math_omml

        elements = parse_math_omml(r"\frac{1}{2}")
        self.assertEqual(elements[0].tag.rsplit("}", 1)[-1], "f")

    def test_inline_math_uses_ordinary_runs_and_true_scripts(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"The clock is $t_{\mathrm{chem}}$ and $ClO_4^-$. ")]
        )
        paragraph = document.paragraphs[0]
        self.assertNotIn("oMath", paragraph._p.xml)
        self.assertTrue(any(run.text == "t" and run.italic for run in paragraph.runs))
        self.assertTrue(any(run.text == "chem" and run.font.subscript for run in paragraph.runs))
        self.assertTrue(any(run.text == "4" and run.font.subscript for run in paragraph.runs))
        self.assertTrue(any(run.text == "–" and run.font.superscript for run in paragraph.runs))

    def test_display_equations_are_omml_and_numbered(self) -> None:
        document = render_blocks_to_doc(
            [Block("equation", r"E = mc^2"), Block("equation", r"a = b")]
        )
        xml = document._element.xml
        self.assertEqual(xml.count("<m:oMathPara"), 2)
        self.assertIn("(1)", xml)
        self.assertIn("(2)", xml)
        self.assertEqual(
            [run.text for paragraph in document.paragraphs for run in paragraph.runs if run.text],
            ["\t", "(1)", "\t", "(2)"],
        )


if __name__ == "__main__":
    unittest.main()

def test_chemical_formula_and_space_group_markup(tmp_path: Path) -> None:
    document = render_blocks_to_doc(
        [Block("paragraph", r"$\mathrm{H_2dabco^{2+}}$ adopts $Pa\bar{3}$.")]
    )
    paragraph = document.paragraphs[0]
    assert any(run.text == "2" and run.font.subscript for run in paragraph.runs)
    assert any(run.text == "2+" and run.font.superscript for run in paragraph.runs)
    assert any(run.text == "3̅" and not run.italic for run in paragraph.runs)

def test_default_table_uses_three_line_rules() -> None:
    document = render_blocks_to_doc(
        [Block("table", rows=(("A", "B"), ("1", "2")))]
    )
    table = document.tables[0]
    borders = table._tbl.tblPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders")
    assert borders is not None
    top = borders.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top")
    bottom = borders.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom")
    inside_h = borders.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}insideH")
    inside_v = borders.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}insideV")
    assert top.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz") == "10"
    assert bottom.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz") == "10"
    assert inside_h.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "nil"
    assert inside_v.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "nil"
    header_borders = table.rows[0].cells[0]._tc.tcPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders"
    )
    header_bottom = header_borders.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom"
    )
    assert header_bottom.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz") == "6"

def test_chemical_bond_hyphens_use_en_dash() -> None:
    value = normalize_typography("N-H, C-N, K-Cl6, A-X12, HClO4-forming, PAP-H2")
    assert value == "N–H, C–N, K–Cl6, A–X12, HClO4-forming, PAP-H2"


def test_scientific_units_and_r_squared_use_true_scripts() -> None:
    document = render_blocks_to_doc(
        [Block("paragraph", r"$R^2$; $\mathrm{cm^3\,mol^{-1}\,s^{-1}}$; $C_2(10\,\mathrm{ps})$.")]
    )
    runs = document.paragraphs[0].runs
    assert any(run.text == "R" and run.italic for run in runs)
    assert any(run.text == "2" and run.font.superscript for run in runs)
    assert any(run.text == "3" and run.font.superscript for run in runs)
    assert sum(run.font.superscript is True and run.text == "–1" for run in runs) == 2
    assert any(run.text == "C" and run.italic for run in runs)
    assert any(run.text == "2" and run.font.subscript for run in runs)


def test_arrhenius_operator_and_upright_subscript() -> None:
    document = render_blocks_to_doc(
        [Block("paragraph", r"$\mathrm{ln}(k)$, $\mathrm{ln}(A)$, and $E_{\mathrm{a}}$.")]
    )
    runs = document.paragraphs[0].runs
    assert any(run.text.startswith("ln") and run.italic is not True for run in runs)
    assert any(run.text == "a" and run.font.subscript and run.italic is not True for run in runs)
    assert any(run.text == "E" and run.italic for run in runs)


def test_explicit_table_caption_block_is_parsed() -> None:
    with TemporaryDirectory() as directory:
        path = Path(directory) / "source.md"
        path.write_text(
            "Table: Example values.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n",
            encoding="utf-8",
        )
        blocks = parse_markdown(path)
        assert [block.kind for block in blocks] == ["table_caption", "table"]
        assert blocks[0].text == "Example values."


def test_malformed_pipe_table_fails_with_source_location(tmp_path: Path) -> None:
    path = tmp_path / "malformed.md"
    path.write_text("| A | B |\n|-|-|\n| 1 | 2 |\n", encoding="utf-8")
    with pytest.raises(ValueError) as excinfo:
        parse_markdown(path)
    assert "Malformed Markdown table" in str(excinfo.value)
