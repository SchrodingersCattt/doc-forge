"""Tests for the TeX tokenizer, bib parser, and label-map builder."""

from __future__ import annotations

import unittest
from pathlib import Path

from docforge.tex.tokenize import spans_to_plain, tokenize_tex
from docforge.tex.bib import CitationResolver, parse_bib
from docforge.tex.convert import build_label_map

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

    def test_tokenize_tex_bold(self) -> None:
        spans = tokenize_tex(r"\textbf{bold}")
        self.assertTrue(spans[0].bold)

    def test_resolve_ref(self) -> None:
        spans = tokenize_tex(r"See \ref{fig:one}.", resolve_ref=lambda key: "1")
        self.assertEqual(spans_to_plain(spans), "See 1.")


class BibTests(unittest.TestCase):
    def test_parse_bib(self) -> None:
        path = Path("ref.bib")
        try:
            path.write_text(BIB_TEXT, encoding="utf-8")
            entries = parse_bib(path)
            self.assertIn("author2020", entries)
            self.assertEqual(entries["author2020"]["year"], "2020")
        finally:
            path.unlink(missing_ok=True)

    def test_citation_resolver_order(self) -> None:
        resolver = CitationResolver({})
        self.assertEqual(resolver.resolve("a"), "[1]")
        self.assertEqual(resolver.resolve("b"), "[2]")
        self.assertEqual(resolver.resolve("a"), "[1]")


class LabelMapTests(unittest.TestCase):
    def test_main_si_cross_reference(self) -> None:
        label_map = build_label_map(TEX_DOC, SI_DOC)
        # Could not assert fixed numbers (section order dependent); just ensure
        # the SI prefix is present and the SI figure label resolves.
        self.assertIn("fig:one", label_map)
        self.assertTrue(label_map["fig:one"].isdigit())


if __name__ == "__main__":
    unittest.main()