"""Tests for the docforge markdown block parser and OMML math builders."""

from __future__ import annotations

import unittest
from pathlib import Path

from docforge.markdown.blocks import Block


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


if __name__ == "__main__":
    unittest.main()