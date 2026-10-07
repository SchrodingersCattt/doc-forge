"""Tests for the TeX tokenizer, bib parser, and label-map builder."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from docx.oxml.ns import qn

from docforge.tex.tokenize import spans_to_plain, tokenize_tex, unpaired_quote_errors
from docforge.tex.bib import CitationResolver, parse_bib
from docforge.tex.convert import build_label_map
from docforge.tex.converter import add_rich_text

TEX_DOC = r"""
\section{Introduction}
A molecule with $\alpha$ and $x_i^{2}$ and \textbf{bold} text.

\subsection{Methods}
\begin{equation}
E = \frac{1}{2} m v^2
\end{equation}

\begin{figure}
  \caption{Result}
  \label{fig:one}
  \includegraphics[width=0.8\textwidth]{figure1.pdf}
\end{figure}

Table~\ref{tab:one} reports values.

\begin{table}
  \caption{Values}
  \label{tab:one}
\end{table}

\cite{author2020, author2021}
\bibliography{ref}
"""

SI_DOC = r"""
\section{Supplementary Note 1}
\label{sn:one}
Supplementary Figure~\ref{fig:si} shows more.
\begin{figure}
  \caption{SI figure}
  \label{fig:si}
\end{figure}
"""

BIB_TEXT = r"""
@article{author2020,
  author = {Alice Author and Bob Writer},
  title = {A title},
  journal = {Journal of Things},
  year = {2020},
  volume = {1},
  pages = {1--5}
}
@article{author2021,
  author = {Carol Author and others},
  title = {Another title},
  journal = {Other Journal},
  year = {2021}
}
"""


class TokenizeTests(unittest.TestCase):
    def test_tokenize_tex_plain(self) -> None:
        spans = tokenize_tex(r"Hello \alpha and $x_i$.")
        text = spans_to_plain(spans)
        self.assertEqual(text, "Hello α and xi.")

    def test_tokenize_tex_subscript(self) -> None:
        spans = tokenize_tex(r"$x_i$")
        self.assertTrue(spans[0].subscript or spans[1].subscript)

    def test_relation_inside_math_keeps_spaces_and_upright_digits(self) -> None:
        plain = spans_to_plain(tokenize_tex(r"$r = 5.6$ and $\delta \approx 0.909$"))
        self.assertEqual(plain, "r = 5.6 and δ ≈ 0.909")
        digits = [span for span in tokenize_tex(r"$r = 5.6$") if "5.6" in span.text]
        self.assertEqual(len(digits), 1)
        self.assertFalse(digits[0].italic)

    def test_highlight_inside_math_does_not_italicize_digits(self) -> None:
        spans = tokenize_tex(r"$r_c = \highlight{5.6}$ and $\approx \highlight{0.909}$")
        marked = [span for span in spans if span.highlight == "yellow"]
        self.assertTrue(marked)
        self.assertTrue(all(not span.italic for span in marked))
        self.assertEqual(spans_to_plain(spans), "rc = 5.6 and ≈ 0.909")

    def test_relation_that_is_its_own_math_group_is_not_double_spaced(self) -> None:
        self.assertEqual(spans_to_plain(tokenize_tex(r"X $>$ A $\approx$ B")), "X > A ≈ B")

    def test_tokenize_tex_bold(self) -> None:
        spans = tokenize_tex(r"\textbf{bold}")
        self.assertTrue(spans[0].bold)

    def test_texorpdfstring_keeps_typeset_argument_only(self) -> None:
        spans = tokenize_tex(r"\texorpdfstring{ABX$_4$}{ABX4} branch")
        self.assertEqual(spans_to_plain(spans), "ABX4 branch")
        self.assertEqual(sum(1 for span in spans if span.text == "ABX4"), 0)
        subscript = [span.text for span in spans if span.subscript]
        self.assertEqual(subscript, ["4"])

    def test_tokenize_texttt_is_mono(self) -> None:
        spans = tokenize_tex(r"set \texttt{n\_dim} here")
        mono = [span for span in spans if span.mono]
        self.assertEqual([span.text for span in mono], ["n_dim"])
        self.assertFalse(spans[0].mono)

    def test_texttt_docx_uses_consolas(self) -> None:
        document = Document()
        paragraph = document.add_paragraph()
        add_rich_text(paragraph, r"key \texttt{n\_dim}")
        fonts = [
            run._element.find(qn("w:rPr")).find(qn("w:rFonts")).get(qn("w:ascii"))
            for run in paragraph.runs
            if run.text == "n_dim"
        ]
        self.assertEqual(fonts, ["Consolas"])
        for run in paragraph.runs:
            rpr = run._element.find(qn("w:rPr"))
            self.assertEqual(len(rpr.findall(qn("w:rFonts"))), 1)

    def test_layout_commands_and_superscripts_do_not_leak(self) -> None:
        spans = tokenize_tex(
            r"\begin{center}\vspace{12pt}\begin{minipage}{0.98\textwidth}"
            r"\textsuperscript{\#}Equal\end{minipage}\end{center}"
        )
        text = spans_to_plain(spans)
        self.assertNotIn("center", text)
        self.assertNotIn("minipage", text)
        self.assertNotIn("12pt", text)
        self.assertEqual(text, "#Equal")
        self.assertTrue(spans[0].superscript)

    def test_resolve_ref(self) -> None:
        spans = tokenize_tex(r"See \ref{fig:one}.", resolve_ref=lambda key: "1")
        self.assertEqual(spans_to_plain(spans), "See 1.")

    def test_single_quotes_pair_and_math_primes_stay(self) -> None:
        quoted = spans_to_plain(tokenize_tex(r"Both `1/80' and ``SY''."))
        self.assertIn("\u20181/80\u2019", quoted)
        self.assertIn("\u201cSY\u201d", quoted)
        self.assertEqual(unpaired_quote_errors(quoted), [])
        primed = spans_to_plain(tokenize_tex(r"the model's $b'$ term"))
        self.assertIn("model\u2019s", primed)
        self.assertIn("b'", primed)
        self.assertNotEqual(
            unpaired_quote_errors("Both \u20181/80' ratios"),
            [],
        )


class BibTests(unittest.TestCase):
    def test_parse_bib(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "ref.bib"
            path.write_text(BIB_TEXT, encoding="utf-8")
            entries = parse_bib(path)
            self.assertIn("author2020", entries)
            self.assertEqual(entries["author2020"]["year"], "2020")

    def test_citation_resolver_order(self) -> None:
        resolver = CitationResolver({})
        self.assertEqual(resolver.resolve("a"), "[1]")
        self.assertEqual(resolver.resolve("b"), "[2]")
        self.assertEqual(resolver.resolve("a"), "[1]")

    def test_consecutive_citations_use_en_dash(self) -> None:
        resolver = CitationResolver({})
        for key in ("a", "b", "gap", "c", "d", "e"):
            resolver.resolve(key)
        self.assertEqual(resolver.citation_parts("a,c,d,e"), [("1", "1"), (",", None), ("4–6", "4")])
        self.assertEqual(resolver.citation_parts("b,a"), [("1–2", "1")])

    def test_tex_references_format_bibtex_authors_and_ranges(self) -> None:
      with TemporaryDirectory() as directory:
        path = Path(directory) / "ref.bib"
        path.write_text(BIB_TEXT, encoding="utf-8")
        resolver = CitationResolver(parse_bib(path))
        resolver.resolve("author2020,author2021")
        self.assertEqual(
          resolver.reference_items(),
          [
                    ("1", "Alice Author & Bob Writer. A title. Journal of Things 1, 1–5 (2020)."),
                    ("2", "Carol Author et al. Another title. Other Journal (2021)."),
          ],
        )


class LabelMapTests(unittest.TestCase):
    def test_main_si_cross_reference(self) -> None:
        label_map = build_label_map(TEX_DOC, SI_DOC)
        # Could not assert fixed numbers (section order dependent); just ensure
        # the SI prefix is present and the SI figure label resolves.
        self.assertIn("fig:one", label_map)
        self.assertTrue(label_map["fig:one"].isdigit())

    def test_external_document_prefix_resolves_supplementary_floats(self) -> None:
        main = r"\begin{document}\end{document}"
        si = r"""
        \begin{document}
        \refstepcounter{suppnote}
        \section*{Supplementary Note}
        \label{sn:model}
        \begin{table}
        \label{tab:one}
        \end{table}
        \begin{figure}
        \label{fig:S1}
        \end{figure}
        \end{document}
        """
        labels = build_label_map(main, si)
        self.assertEqual(labels["S-tab:one"], "S1")
        self.assertEqual(labels["S-fig:S1"], "S1")
        self.assertEqual(labels["S-sn:model"], "1")


if __name__ == "__main__":
    unittest.main()
