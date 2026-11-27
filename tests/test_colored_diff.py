"""Color-only DOCX comparison preserves text and contains no Word revisions."""

import json
import tempfile
import unittest
from pathlib import Path

from docx import Document

from docforge.docxdiff.colored import create_colored_docx
from docforge.markdown.launcher import verify_clean_review_state


class ColoredDiffTests(unittest.TestCase):
    def test_colors_only_changed_text_and_defaults_ambiguous_attribution(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            base, current, result = (root / name for name in ("base.docx", "current.docx", "marked.docx"))
            old = Document()
            new = Document()
            for a, b in (("Header", "Header"), ("First old text", "First new text"),
                         ("Second old text", "Second new text"), ("Footer", "Footer")):
                old.add_paragraph(a)
                new.add_paragraph(b)
            old.save(base)
            new.save(current)
            rules = root / "rules.json"
            rules.write_text(json.dumps([
                {"pattern": "^First", "color": "0070C0"},
                {"pattern": "^Second", "color": "C0504D"},
                {"pattern": "^Second new", "color": "548235"},
            ]), encoding="utf-8")
            summary = create_colored_docx(base, current, result, color_map_path=rules)
            marked = Document(result)
            self.assertEqual([p.text for p in marked.paragraphs], [p.text for p in new.paragraphs])
            self.assertTrue(any(r.text.strip() and str(r.font.color.rgb) == "0070C0"
                                for r in marked.paragraphs[1].runs))
            self.assertTrue(any(r.text.strip() and str(r.font.color.rgb) == "FF0000"
                                for r in marked.paragraphs[2].runs))
            self.assertGreater(summary["colored_runs"], 0)
            verify_clean_review_state(result)

    def test_uses_table_caption_for_added_cell_text(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            base, current, result = (root / name for name in ("base.docx", "current.docx", "marked.docx"))
            old = Document()
            new = Document()
            for doc in (old, new):
                doc.add_paragraph("Introduction")
                doc.add_paragraph("Table S13. Energetic properties")
                table = doc.add_table(rows=1, cols=1)
                table.cell(0, 0).text = "Material"
                doc.add_paragraph("End")
            new.tables[0].add_row().cells[0].text = "Added material"
            old.save(base)
            new.save(current)
            rules = root / "rules.json"
            rules.write_text('[{"pattern":"^Table S13\\\\.","color":"548235"}]', encoding="utf-8")
            create_colored_docx(base, current, result, color_map_path=rules)
            marked = Document(result)
            cell = marked.tables[0].cell(1, 0)
            self.assertEqual(cell.text, "Added material")
            self.assertTrue(any(str(run.font.color.rgb) == "548235" for p in cell.paragraphs for run in p.runs))
            verify_clean_review_state(result)


if __name__ == "__main__":
    unittest.main()
