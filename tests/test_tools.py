"""Tests for the source-pack and aigc helpers (no network)."""

from __future__ import annotations

import unittest
import zipfile
from pathlib import Path

from docforge.sourcepack import package_files
from docforge.aigc import load_prompt, discover_prompts, select_prompts


class SourcePackTests(unittest.TestCase):
    def test_package_files_manifest(self) -> None:
        root = Path("__tmp_pack_root")
        (root / "docs").mkdir(parents=True, exist_ok=True)
        (root / "main.tex").write_text(r"\documentclass{article}", encoding="utf-8")
        (root / "docs" / "note.md").write_text("note", encoding="utf-8")
        try:
            output = package_files(
                root,
                allowlist=["main.tex"],
                glob_dirs=[root / "docs"],
                output=root / "dist" / "pkg.zip",
            )
            self.assertTrue(output.exists())
            with zipfile.ZipFile(output) as archive:
                names = set(archive.namelist())
                self.assertIn("main.tex", names)
                self.assertIn("docs/note.md", names)
                self.assertIn("SOURCE_PACKAGE_MANIFEST.csv", names)
        finally:
            import shutil

            shutil.rmtree(root, ignore_errors=True)


class AigcTests(unittest.TestCase):
    def test_load_prompt_front_matter(self) -> None:
        path = Path("__tmp_prompt.md")
        path.write_text(
            "---\noutput_name: fig1\naspect_ratio: 4:3\nstatus: draft\n---\nDraw a molecule.\n",
            encoding="utf-8",
        )
        try:
            spec = load_prompt(path, root=Path("."))
            self.assertEqual(spec.output_name, "fig1")
            self.assertEqual(spec.aspect_ratio, "4:3")
            self.assertEqual(spec.prompt, "Draw a molecule.")
        finally:
            path.unlink(missing_ok=True)

    def test_discover_and_select(self) -> None:
        prompt_dir = Path("__tmp_prompts")
        prompt_dir.mkdir(exist_ok=True)
        (prompt_dir / "a_prompt.md").write_text("# A\n", encoding="utf-8")
        (prompt_dir / "b_prompt.md").write_text("# B\n", encoding="utf-8")
        try:
            paths = discover_prompts(prompt_dir)
            self.assertEqual(len(paths), 2)
            selected = select_prompts(paths, ["a"], all_prompts=False)
            self.assertEqual(len(selected), 1)
            self.assertEqual(selected[0].stem, "a_prompt")
        finally:
            import shutil

            shutil.rmtree(prompt_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()