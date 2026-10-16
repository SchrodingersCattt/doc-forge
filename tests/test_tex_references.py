"""Regression tests for the reference list written from TeX sources."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

from docforge.tex.converter import latex_to_docx
from docforge.tex.tokenize import plain_tex

BIB = {
    "Early2020": {"author": "Early, A.", "title": "Cited in the main text",
                  "journal": "J. Test", "volume": "1", "pages": "1--2", "year": "2020"},
    "Late2021": {"author": "Late, B.", "title": "Cited only in Methods",
                 "journal": "Ironmaking \\& Steelmaking", "volume": "2", "pages": "3--4",
                 "year": "2021"},
    "Later2022": {"author": "Later, C.", "title": "Also cited only in Methods",
                  "journal": "J. Test", "volume": "3", "pages": "5--6", "year": "2022"},
}

SOURCE = r"""
\begin{document}
\section*{Introduction}
Main text cites one work\cite{Early2020}.
\bibliographystyle{unsrtnat}
\bibliography{ref}
\section*{Methods}
Methods cite two more works\cite{Late2021,Later2022}.
% A commented citation\cite{Early2020,Commented2023} must not be numbered.
Repeated citation\cite{Early2020}.
\end{document}
"""


def _reference_paragraphs(path: Path) -> list[str]:
    texts = [p.text for p in Document(str(path)).paragraphs]
    start = texts.index("References") + 1
    return [t for t in texts[start:] if t.startswith("[")]


class TrailingCitationTests(unittest.TestCase):
    def test_methods_citations_after_bibliography_are_listed(self) -> None:
        with TemporaryDirectory() as tmp:
            output = Path(tmp) / "paper.docx"
            latex_to_docx(SOURCE, bib=BIB, output=output)
            refs = _reference_paragraphs(output)
        self.assertEqual(len(refs), 3)
        self.assertTrue(refs[0].startswith("[1] Early"))
        self.assertTrue(refs[1].startswith("[2] Late"))
        self.assertTrue(refs[2].startswith("[3] Later"))
        self.assertIn("Ironmaking & Steelmaking", refs[1])
        self.assertNotIn("\\&", refs[1])


class EscapedAmpersandTests(unittest.TestCase):
    def test_escaped_ampersand_renders_plain(self) -> None:
        self.assertEqual(plain_tex(r"Materials \& Design"), "Materials & Design")


if __name__ == "__main__":
    unittest.main()
