"""Tests for the docforge markdown block parser and OMML math builders."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pytest
from lxml import etree

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
                "![Figure 1|span=page|orientation=portrait](figure.png)\n\n"
                "Figure 1. Caption.\n",
                encoding="utf-8",
            )
            blocks = parse_markdown(path)
            self.assertEqual([block.kind for block in blocks], ["paragraph", "image", "paragraph"])
            self.assertEqual(
                blocks[1].options,
                (("span", "page"), ("orientation", "portrait")),
            )
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

    def test_parse_display_pmatrix(self) -> None:
        from docforge.markdown.omml import parse_math_omml

        elements = parse_math_omml(
            r"M=\begin{pmatrix}3&0&0\\0&2&2\\0&-2&2\end{pmatrix}"
        )
        names = [node.tag.rsplit("}", 1)[-1] for element in elements for node in element.iter()]
        self.assertIn("d", names)
        self.assertIn("m", names)
        self.assertEqual(names.count("mr"), 3)
        self.assertEqual(names.count("e"), 10)
        self.assertNotIn("mPr", names)
        self.assertNotIn("count", names)

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

    def test_inline_math_preserves_reaction_arrows(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"$A \rightarrow B$ and $C \longrightarrow D$.")]
        )
        text = document.paragraphs[0].text
        self.assertEqual(text, "A → B and C ⟶ D.")

    def test_inline_math_spaces_binary_operators(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"$\tau=t-t_{\mathrm{chem}}$ and $E=-mR$.")]
        )
        text = document.paragraphs[0].text
        self.assertEqual(text, "τ = t – tchem and E = –mR.")

        spaced = render_blocks_to_doc(
            [Block("paragraph", r"$\tau = t - t_{\mathrm{chem}}$.")]
        )
        self.assertEqual(spaced.paragraphs[0].text, "τ = t – tchem.")

        uncertainty = render_blocks_to_doc(
            [Block("paragraph", r"$61.3 \pm 1.1\%$ and $t_{\mathrm{chem}} = 0$.")]
        )
        self.assertEqual(uncertainty.paragraphs[0].text, "61.3 ± 1.1% and tchem = 0.")
        self.assertNotIn("  ", uncertainty.paragraphs[0].text)

        split_math = render_blocks_to_doc(
            [Block("paragraph", r"61.3 $\pm$ 1.1% and mean $\pm$ S.E.")]
        )
        self.assertEqual(split_math.paragraphs[0].text, "61.3 ± 1.1% and mean ± S.E.")
        self.assertNotIn("  ", split_math.paragraphs[0].text)

    def test_standalone_font_override_preserves_inline_formatting(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", "C<sub>36</sub>Fe *d* **bold**")],
            font_family="Times New Roman",
            east_asia_font="Times New Roman",
        )
        runs = [run for paragraph in document.paragraphs for run in paragraph.runs]
        self.assertTrue(all(run.font.name == "Times New Roman" for run in runs))
        self.assertTrue(any(run.bold and run.text == "bold" for run in runs))
        self.assertTrue(any(run.italic and run.text == "d" for run in runs))
        self.assertTrue(any(run.font.subscript and run.text == "36" for run in runs))

    def test_display_equations_are_omml_and_numbered(self) -> None:
        document = render_blocks_to_doc(
            [Block("equation", r"E = mc^2"), Block("equation", r"a = b")]
        )
        xml = document._element.xml
        self.assertEqual(xml.count("<m:oMathPara>"), 2)
        self.assertEqual(len(document.tables), 2)
        self.assertEqual([table.cell(0, 2).text for table in document.tables], ["(1)", "(2)"])
        for table in document.tables:
            self.assertIn("oMathPara", table.cell(0, 1)._tc.xml)
            self.assertNotIn("w:tab", table._tbl.xml)

    def test_display_equation_source_preserves_authored_lines(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text(
                "$$\n\\begin{aligned}\na &= b,\\\\\nc &= d.\n\\end{aligned}\n$$\n",
                encoding="utf-8",
            )
            blocks = parse_markdown(path)
        self.assertEqual(
            blocks[0].text,
            "\\begin{aligned}\na &= b,\\\\\nc &= d.\n\\end{aligned}",
        )

    def test_unclosed_and_empty_display_equations_fail(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text("$$\na=b\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"Unclosed display equation.*:1"):
                parse_markdown(path)
            path.write_text("$$\n\n$$\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"Empty display equation.*:1"):
                parse_markdown(path)

    def test_pandoc_math_backend_emits_native_omml_structures(self) -> None:
        document = render_blocks_to_doc(
            [
                Block(
                    "equation",
                    r"""\begin{aligned}
M_{i,\ell} &= \sqrt{\frac14+\sum_{j\in\mathcal N_i}\chi_{ij,\ell}^2},\\
\mathbf h^{(\tau)} &= \mathbf v^{(\tau)}\odot\mathrm{SiLU}\bigl(\mathbf g^{(\tau)}\bigr),\\
\Pi_{i,\eta\kappa} &= \left\lVert\mathbf Q_\eta\mathbf v_\kappa\right\rVert^2.
\end{aligned}""",
                )
            ]
        )
        xml = document._element.xml
        self.assertEqual(xml.count("<m:oMathPara>"), 1)
        # Pandoc 2.9 emits ``m:m`` while 3.9 emits an equivalent ``m:eqArr``
        # wrapper for aligned rows.
        self.assertTrue("<m:m>" in xml or "<m:eqArr>" in xml)
        self.assertIn("<m:rad>", xml)
        self.assertIn("<m:f>", xml)
        self.assertIn("<m:nary>", xml)
        self.assertIn("<m:d>", xml)
        self.assertTrue(
            '<m:sty m:val="b"' in xml or '<m:sty m:val="bi"' in xml
        )
        root = etree.fromstring(document._element.xml.encode("utf-8"))
        rendered_math = "".join(
            root.xpath(
                "//m:oMathPara//m:t/text()",
                namespaces={"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"},
            )
        )
        for leaked in ("mathcal", "mathbf", "left", "right", "lVert", "rVert"):
            self.assertNotIn(leaked, rendered_math)
        self.assertIn("(1)", xml)

    def test_math_text_may_contain_command_like_words(self) -> None:
        document = render_blocks_to_doc(
            [Block("equation", r"x_{\mathrm{left}}=\operatorname{sqrt}(y)")]
        )
        xml = document._element.xml
        self.assertIn("left", xml)
        self.assertIn("sqrt", xml)

    def test_standalone_heading_text_and_style_are_black(self) -> None:
        document = render_blocks_to_doc(
            [Block("heading", "Methods", level=1), Block("heading", "Details", level=2)]
        )
        for paragraph in document.paragraphs:
            if not paragraph.style.name.startswith("Heading"):
                continue
            assert paragraph.style.font.color.rgb is not None
            assert str(paragraph.style.font.color.rgb) == "000000"
            assert paragraph.runs
            assert all(run.font.color.rgb is not None for run in paragraph.runs)
            assert all(str(run.font.color.rgb) == "000000" for run in paragraph.runs)


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
    value = normalize_typography("N-H, C-N, K-Cl6, A-X12, HClO4-forming, ABX4")
    assert value == "N–H, C–N, K–Cl6, A–X12, HClO4-forming, ABX4"


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
