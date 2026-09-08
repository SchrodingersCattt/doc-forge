# docforge

Assemble Markdown, LaTeX, and AI-generated art into auditable Word documents.

`docforge` is a project-agnostic Python toolkit distilled from the NSFC
proposal and manuscript tooling that previously lived inside individual
research repositories. It keeps the three properties those scripts had:

- a thin CLI that runs standalone from a checkout without installation;
- explicit input/output paths instead of a hard-coded repository layout;
- auditability: line-level revisions, sha256 sidecars, and manifest CSVs
  accompany every generated artifact.

## Why "docforge"

The package covers more than papers: grant proposals, reports, and
manuscripts all follow the same pattern — plain-text sources in, a Word
document out, with provenance that survives review. *forge* (手工锻造)
captures the deterministic, artifact-producing nature of the toolchain
better than paper-specific names such as "tex2docx" or "papertools".

## Install

```bash
pip install -e .            # core (python-docx, lxml, Pillow)
pip install -e ".[plot]"    # + matplotlib figure style helpers
pip install -e ".[aigc]"    # + OpenAI-compatible / Google image backends
pip install -e ".[tex]"     # + numpy for TeX rendering
```

Python 3.10+.

## CLI

```bash
# Markdown -> DOCX (merged blocks)
docforge md2docx section_01.md section_02.md -o proposal.docx --title "..."

# Word tracked-changes redline against a reviewed DOCX
docforge redline reviewed.docx fresh.docx -o fresh_tracked.docx

# LaTeX manuscript -> DOCX (main + optional SI + bib)
docforge tex2docx main.tex --si si.tex --bib ref.bib -o main.docx

# Generate figures from front-matter Markdown prompts (auditable sidecars)
docforge aigc figures/_prompts --backend litellm --all
docforge aigc figures/_prompts --dry-run

# Auditable TeX source ZIP from an explicit allowlist
docforge sourcepack . --allow main.tex --allow si.tex --glob-dir figures -o dist/tex_source.zip

# Validate a generated DOCX
docforge check output.docx --verify-clean
```

All paths are explicit; nothing is inferred from the current directory.

## Library layout

| Module | Purpose |
| ------ | ------- |
| `docforge.markdown` | Conservative Markdown block parser + python-docx renderer (headings, lists, tables, code, math → native OMML equations, images, `<u>**..**</u>` bold-underline tokens), template-section preservation, unicode-script conversion, navigation/typography verification |
| `docforge.docxdiff` | Build a DOCX whose final view is the current document with all edits as native Word revisions (`w:ins`/`w:del`/`w:move*`/`w:rPrChange`), carrying review comments, comment parts, and package validation |
| `docforge.tex` | LaTeX subset → styled DOCX: sections, floats, tables, algorithms, math, BibTeX parsing, cross-reference label maps (`main.tex` ↔ `si.tex` with `externaldocument` prefixes) |
| `docforge.aigc` | Front-matter Markdown prompt files → image generation (OpenAI-compatible LiteLLM or Google GenAI) with a JSON sidecar recording prompt sha256, provider metadata, and artifact hash |
| `docforge.sourcepack` | Whitelisted, SHA-256-manifested TeX source ZIP (`SOURCE_PACKAGE_MANIFEST.csv` + README inside the archive) |
| `docforge.plotting` | Shared Matplotlib visual system: restrained palette, mm-based figure sizes, panel labels, schematic boxes/arrows/status chips, PDF/SVG/600-dpi PNG output |
| `docforge.cli` | Thin `docforge` entry point; feature modules are imported lazily per subcommand |

## Auditability conventions

- Every generated image carries a `.json` sidecar with `prompt.sha256`,
  provider/endpoint metadata, and the artifact's own sha256.
- Every source ZIP embeds `SOURCE_PACKAGE_MANIFEST.csv` (`Path, Size_bytes, SHA256`).
- Every tracked-diff DOCX is validated so that accepting all revisions
  reproduces the fresh document exactly (text and non-text layout controls),
  and relationship ids cannot leak across packages.
- Clean DOCX outputs reject residual review metadata and em dashes; numeric
  ranges must use en dashes.

## Development

```bash
pip install -e ".[tests]"
pytest
```

## License

MIT — see [LICENSE](LICENSE).