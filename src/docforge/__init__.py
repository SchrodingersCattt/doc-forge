"""docforge: assemble Markdown, LaTeX, and AI-generated art into DOCX.

docforge is a project-agnostic Python toolkit for turning plain-text sources
into Word documents. It was distilled from document tooling that previously
lived inside individual research repositories, so the implementation keeps
three properties that those scripts had:

- a thin CLI that can run standalone from a checkout without installation;
- explicit input/output paths instead of hard-coded repository layout;
- auditability: line-level revisions, sha256 sidecars, and manifest CSV
  accompany every generated artifact.

Core capabilities
-----------------
- ``docforge.markdown`` — parse a conservative Markdown subset (headings,
  paragraphs, bullet/ordered lists, tables, code blocks, fenced math,
  quotes, images, separators, reference lines, ``<u>**..**</u>``
  underline-bold tokens) into blocks, render them into a template DOCX
  with python-docx, retain official template sections, and verify the
  generated package (headings, image relationships, typography, scripts).
- ``docforge.docxdiff`` — build a second DOCX that shows every paragraph,
  token, and inline-format change as native Word tracked revisions
  (w:ins/w:del/w:moveFrom/w:moveTo/w:rPrChange), carrying review comments
  and their anchors forward and validating the revised package.
- ``docforge.tex`` — convert a LaTeX manuscript (main + SI, cross
  references, BibTeX entries, ``figure``/``table``/``longtable``/
  ``algorithm`` environments, math with fractions) into a styled DOCX.
- ``docforge.aigc`` — generate proposal/marketing figures from Markdown
  prompt files through an OpenAI-compatible LiteLLM image endpoint or the
  Google GenAI image API, recording the prompt sha256 and provider
  metadata in a JSON sidecar next to every image.
- ``docforge.sourcepack`` — build a lightweight, auditable TeX source ZIP
  from an explicit allowlist, with an embedded SHA-256 manifest CSV and
  README.
- ``docforge.bibliography`` — normalized JSON/BibTeX entries, shared citation
  numbering, inherited maps, and formatter profiles.
- ``docforge.gates`` — project-neutral abbreviation and TeX reference gates.
- ``docforge.plotting`` — shared Matplotlib visual system (restrained
  palette, mm-based figure sizes, PDF/SVG/600-dpi PNG output helpers)
  used by the figure-drawing entry points.
"""

from __future__ import annotations

__version__ = "0.1.0"
