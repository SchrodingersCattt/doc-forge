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

    def test_table_row_width_mismatch_fails_with_source_location(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text("| A | B |\n|---|---|\n| one |\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"Malformed Markdown table.*:3; row has 1 cells, expected 2"):
                parse_markdown(path)

    def test_table_rows_keep_pipes_inside_math_code_and_escapes(self) -> None:
        from docforge.markdown.launcher import split_table_row

        row = r"| label | `$x|y$` | $x|y$ | escaped \| pipe |"
        assert split_table_row(row) == (
            "label",
            "`$x|y$`",
            "$x|y$",
            "escaped | pipe",
        )

        # Backtick spans may use more than one delimiter character.
        assert split_table_row("| one | ``a|b`` | three |") == (
            "one", "``a|b``", "three"
        )

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


def test_cli_standalone_manuscript_preset(tmp_path: Path) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docforge.cli import main

    source = tmp_path / "source.md"
    source.write_text("# Heading\n\nExample paragraph.\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    assert main(["md2docx", str(source), "-o", str(output), "--standalone-style", "manuscript"]) == 0
    heading, body = Document(output).paragraphs
    assert heading.style.name == "Heading 1"
    assert body.alignment == WD_ALIGN_PARAGRAPH.LEFT
    assert body.runs[0].font.size.pt == 11


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

    def test_inline_chemical_delimiters_use_formatted_runs(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"The unit is $[\mathrm{K}(\mathrm{ClO}_4)_3]$.")]
        )
        paragraph = document.paragraphs[0]
        self.assertNotIn("oMath", paragraph._p.xml)
        self.assertEqual(paragraph.text, "The unit is [K(ClO4)3].")
        placed = [
            (char, "subscript" if run.font.subscript else "superscript" if run.font.superscript else "baseline")
            for run in paragraph.runs
            for char in run.text
        ]
        self.assertIn(("4", "subscript"), placed)
        self.assertIn(("3", "subscript"), placed)
        for char in "()[]":
            self.assertIn((char, "baseline"), placed)
            self.assertNotIn((char, "subscript"), placed)

    def test_inline_chemical_delimiter_keeps_following_space(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"loss of $[\mathrm{K}(\mathrm{ClO}_4)_6]$ coordination")]
        )
        xml = document.paragraphs[0]._p.xml
        self.assertNotIn("oMath", xml)
        self.assertIn('xml:space="preserve"> coordination', xml)

    def test_inline_math_preserves_reaction_arrows(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"$A \rightarrow B$ and $C \longrightarrow D$.")]
        )
        text = document.paragraphs[0].text
        self.assertEqual(text, "A → B and C ⟶ D.")

    def test_inline_math_preserves_triangle_markers(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"Series A ($\triangle$) and B ($\triangledown$).")]
        )
        self.assertEqual(document.paragraphs[0].text, "Series A (△) and B (▽).")

    def test_inline_math_greek_symbols_and_unbraced_macro_subscript(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"$\mathbf v_\kappa$, $\Pi_{i,\eta\kappa}$, "
                r"$\boldsymbol\Xi$, $\boldsymbol\Theta_\tau$, and $\otimes$.")]
        )
        paragraph = document.paragraphs[0]
        self.assertEqual(paragraph.text, "vκ, Πi,ηκ, Ξ, Θτ, and ⊗.")
        self.assertTrue(any(run.text == "κ" and run.font.subscript for run in paragraph.runs))
        self.assertTrue(any(run.text == "Ξ" and run.bold for run in paragraph.runs))
        self.assertTrue(any(run.text == "τ" and run.font.subscript for run in paragraph.runs))

    def test_inline_bold_vector_bar_and_max_operator(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", r"mean $\bar{\mathbf f}_i$ and $\max_i\mathrm{dev}_i$.")]
        )
        paragraph = document.paragraphs[0]
        self.assertEqual(paragraph.text, "mean f\u0305i and maxidevi.")
        self.assertTrue(any("f\u0305" in run.text and run.bold for run in paragraph.runs))
        self.assertTrue(any(run.text == "i" and run.font.subscript for run in paragraph.runs))

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

    def test_doubled_dollar_inside_prose_stays_literal(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", "Keep $$broken x + y$$ spaces after.")]
        )
        self.assertEqual(document.paragraphs[0].text, "Keep $$broken x + y$$ spaces after.")
        self.assertEqual(len(document.paragraphs[0].runs), 1)

    def test_doubled_dollar_inside_source_paragraph_reports_location(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text("Before $$broken display$$ after.\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"Display equation delimiter.*source\.md:1"):
                parse_markdown(path)

    def test_doubled_dollar_in_code_span_or_escape_stays_literal(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text("Use `price$$value` or `\\$\\$` or \\$\\$ as literal text.\n", encoding="utf-8")
            blocks = parse_markdown(path)
        document = render_blocks_to_doc(blocks)
        self.assertEqual(document.paragraphs[0].text, "Use price$$value or \\$\\$ or $$ as literal text.")

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

    def test_standalone_defaults_use_times_and_justified_body(self) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn

        document = render_blocks_to_doc([
            Block("heading", "A heading", level=2),
            Block("paragraph", "Ordinary *emphasis* and H<sub>2</sub>O."),
        ])
        body = next(paragraph for paragraph in document.paragraphs if paragraph.text.startswith("Ordinary"))
        self.assertEqual(body.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
        self.assertEqual(document.styles["Normal"].paragraph_format.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
        self.assertEqual(document.styles["Normal"].font.name, "Times New Roman")
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                if run.text:
                    self.assertEqual(run.font.name, "Times New Roman")
                    self.assertEqual(run._element.rPr.rFonts.get(qn("w:eastAsia")), "Times New Roman")
        self.assertTrue(any(run.italic and run.text == "emphasis" for run in body.runs))
        self.assertTrue(any(run.font.subscript and run.text == "2" for run in body.runs))

    def test_standalone_manuscript_preset_has_honest_heading_levels(self) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn

        document = render_blocks_to_doc(
            [Block("heading", "Methods", level=1), Block("heading", "Details", level=2),
             Block("paragraph", "Plain *emphasis* and H<sub>2</sub>O.")],
            title="An example", standalone_style="manuscript",
        )
        title, first, second, body = document.paragraphs
        self.assertEqual(title.style.name, "Title")
        self.assertIsNone(title.style._element.get_or_add_pPr().find(qn("w:pBdr")))
        self.assertEqual(str(title.style.font.color.rgb), "000000")
        self.assertEqual(first.style.name, "Heading 1")
        self.assertEqual(second.style.name, "Heading 2")
        self.assertEqual(first._p.pPr.find(qn("w:outlineLvl")).get(qn("w:val")), "0")
        self.assertEqual(second._p.pPr.find(qn("w:outlineLvl")).get(qn("w:val")), "1")
        self.assertEqual(body.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertEqual(document.styles["Normal"].font.size.pt, 11)
        self.assertTrue(all(run.font.size.pt == 11 for run in body.runs if run.text))
        self.assertTrue(any(run.italic and run.text == "emphasis" for run in body.runs))
        self.assertTrue(any(run.font.subscript and run.text == "2" for run in body.runs))

    def test_inline_code_keeps_consolas_in_rendered_blocks(self) -> None:
        document = render_blocks_to_doc(
            [
                Block("paragraph", "Use `loss.type` in prose."),
                Block("bullet", "Select `ener` here."),
                Block("heading", "Setting `loss.type`", level=2),
                Block("table", rows=(("Key", "Value"), ("`loss.type`", "`ener`"))),
            ],
            font_family="Arial",
            east_asia_font="SimHei",
        )
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                if run.text in {"loss.type", "ener"}:
                    self.assertEqual(run.font.name, "Consolas")
        for cell in document.tables[0].rows[1].cells:
            self.assertEqual(next(run for run in cell.paragraphs[0].runs if run.text).font.name, "Consolas")
        self.assertEqual(document.paragraphs[0].runs[0].font.name, "Arial")

    def test_escaped_table_pipe_is_plain_cell_text_in_docx(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text("| Header | Value |\n| --- | --- |\n| left \\| right | ok |\n", encoding="utf-8")
            blocks = parse_markdown(path)
        document = render_blocks_to_doc(blocks)
        assert document.tables[0].cell(1, 0).text == "left | right"

    def test_nested_scripts_inside_bold_or_italic_markup_are_preserved(self) -> None:
        document = render_blocks_to_doc(
            [Block("paragraph", "**(C<sub>6</sub>H<sub>8</sub>N)SnCl<sub>3</sub>**")]
        )
        runs = document.paragraphs[0].runs
        assert any(run.text == "6" and run.font.subscript and run.bold for run in runs)
        assert any(run.text == "3" and run.font.subscript and run.bold for run in runs)

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

    def test_math_fence_is_native_display_equation(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.md"
            path.write_text("``` math\nD_{zz}(r)=\\frac{n(+E_z)}{E_z^2}\n```\n", encoding="utf-8")
            blocks = parse_markdown(path)
        self.assertEqual([block.kind for block in blocks], ["equation"])
        document = render_blocks_to_doc(blocks)
        self.assertIn("oMathPara", document._element.xml)


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


def test_adjacent_parentheticals_merge_into_the_preceding_group() -> None:
    value = normalize_typography(r"($\eta^2 = 0.12$) (Figure 2b). (a) text. (b) next.")
    assert value == r"($\eta^2 = 0.12$, Figure 2b). (a) text. (b) next."
    assert normalize_typography("(H$_2$dabco)(NH$_4$)(ClO$_4$)$_3$") == "(H$_2$dabco)(NH$_4$)(ClO$_4$)$_3$"
    assert normalize_typography("(foo) (bar)") == "(foo) (bar)"
    source = "(Figures S1–S29 and Tables S1–S17) (PDF)"
    assert normalize_typography(source) == source
    assert normalize_typography("(0.12) (Table S2)") == "(0.12, Table S2)"


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


def test_omml_paragraphs_keep_independent_lines_and_math():
    import shutil
    from docforge.markdown import render_omml_paragraphs

    if shutil.which("pandoc") is None:
        pytest.skip("Pandoc is required for native Word math")
    paragraphs = render_omml_paragraphs([
        r"**Input:** $x_1$",
        r"**if** $\lvert A \rvert \geq 2$ **then** $B \setminus C$",
    ])
    assert len(paragraphs) == 2
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
    assert "Input:" in "".join(paragraphs[0].xpath(".//w:t/text()", namespaces=ns))
    assert paragraphs[0].xpath(".//m:oMath", namespaces=ns)
    assert len(paragraphs[1].xpath(".//m:oMath", namespaces=ns)) == 2


def test_standalone_cli_defaults_preserve_scripts_and_formatting(tmp_path: Path) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docforge.cli import main

    source = tmp_path / "source.md"
    source.write_text("Dear Reader,\n\nA *journal* and H<sub>2</sub>O are examples.\n", encoding="utf-8")
    output = tmp_path / "output.docx"
    assert main(["md2docx", str(source), "-o", str(output)]) == 0
    document = Document(output)
    for paragraph in document.paragraphs:
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
        assert not paragraph.paragraph_format.first_line_indent
        assert paragraph.paragraph_format.space_after.pt == 6
        assert all(run.font.name == "Times New Roman" for run in paragraph.runs if run.text)
    assert any(run.font.subscript and run.text == "2" for run in document.paragraphs[1].runs)
    assert any(run.italic and run.text == "journal" for run in document.paragraphs[1].runs)


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


def test_numbered_table_caption_is_scoped_to_the_following_table(tmp_path: Path) -> None:
    path = tmp_path / "source.md"
    path.write_text(
        "Table 2. An ordinary paragraph.\n\n"
        "Table 7.Caption without a separating space.\n\n"
        "| A | B | C | D |\n| :--- | :---: | ---: | --- |\n| a | b | c | d |\n",
        encoding="utf-8",
    )
    blocks = parse_markdown(path)
    assert [block.kind for block in blocks] == ["paragraph", "table_caption", "table"]
    assert dict(blocks[1].options) == {
        "caption_label": "Table 7", "caption_punctuation": ".", "caption_gap": "",
    }
    assert blocks[2].alignments == ("left", "center", "right", "")
    document = render_blocks_to_doc(blocks)
    assert document.paragraphs[1].text == "Table 7.Caption without a separating space."
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    for row in document.tables[0].rows:
        assert [cell.paragraphs[0].alignment for cell in row.cells] == [
            WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT,
        ]


def test_authored_and_generated_table_numbers_cannot_collide(tmp_path: Path) -> None:
    path = tmp_path / "source.md"
    path.write_text(
        "Table: First generated.\n\n| A |\n|---|\n| one |\n\n"
        "Table 1. Authored.\n\n| A |\n|---|\n| two |\n\n"
        "Table: Next generated.\n\n| A |\n|---|\n| three |\n",
        encoding="utf-8",
    )
    with pytest.warns(UserWarning, match="table.number_collision"):
        document = render_blocks_to_doc(parse_markdown(path))
    assert [paragraph.text for paragraph in document.paragraphs if paragraph.text] == [
        "Table 2. First generated.", "Table 1. Authored.", "Table 3. Next generated.",
    ]


def test_table_and_note_keep_an_explicit_superscript(tmp_path: Path) -> None:
    path = tmp_path / "source.md"
    path.write_text(
        "| Column<sup>a</sup> |\n|---|\n| item |\n\n"
        "Table note: <sup>a</sup>Note text.\n",
        encoding="utf-8",
    )
    document = render_blocks_to_doc(parse_markdown(path))
    table = document.tables[0]
    note = next(paragraph for paragraph in document.paragraphs if paragraph.text)
    assert note.text == "aNote text."
    assert note.runs[0].text == "a" and note.runs[0].font.superscript is True
    header = table.cell(0, 0).paragraphs[0]
    assert header.text == "Columna"
    assert header.runs[-1].text == "a" and header.runs[-1].font.superscript is True


def test_duplicate_authored_table_labels_are_reported(tmp_path: Path) -> None:
    path = tmp_path / "source.md"
    path.write_text(
        "Table 4. One.\n\n| A |\n|---|\n| one |\n\n"
        "table 4. Two.\n\n| A |\n|---|\n| two |\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="Duplicate authored table caption label"):
        render_blocks_to_doc(parse_markdown(path))


def test_semantic_adapter_preserves_numbered_caption_and_column_alignments(tmp_path: Path) -> None:
    from docforge.adapters.markdown import MarkdownReader, render_latex, render_markdown
    from docforge.adapters.protocol import ReadOptions

    path = tmp_path / "source.md"
    path.write_text(
        "Table 8. Measurements.\n\n"
        "| First | Second | Third |\n| :--- | :---: | ---: |\n| a | b | c |\n",
        encoding="utf-8",
    )
    document = MarkdownReader().read(path, options=ReadOptions()).document
    markdown = render_markdown(document)
    assert "Table 8. Measurements." in markdown
    assert "| :--- | :---: | ---: |" in markdown
    assert r"\begin{tabular}{lcr}" in render_latex(document)


def test_malformed_pipe_table_fails_with_source_location(tmp_path: Path) -> None:
    path = tmp_path / "malformed.md"
    path.write_text("| A | B |\n|-|-|\n| 1 | 2 |\n", encoding="utf-8")
    with pytest.raises(ValueError) as excinfo:
        parse_markdown(path)
    assert "Malformed Markdown table" in str(excinfo.value)

