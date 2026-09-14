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

The `docx2md` and `docx2tex` commands require the external `pandoc` command
to be installed and available on `PATH`. docforge does not install, discover,
or replace Pandoc.

## CLI

```bash
# Markdown -> DOCX (merged blocks)
docforge md2docx section_01.md section_02.md -o proposal.docx --title "..."

# Markdown -> DOCX using an existing Word template and ordered references
docforge md2docx 01.abstract.md 02.main.md -o main.docx \
  --template MolCrysKit_JCIM-0118.docx --metadata 00.metadata.md \
  --bibliography 07.references.json --style template \
  --font "Times New Roman" --east-asia-font "Times New Roman" \
  --line-numbers on --columns two --force

# Supporting information can request a readable one-column body
docforge md2docx 06.supporting-information.md -o si.docx \
  --template MolCrysKit_JCIM-0118.docx --metadata 00.metadata.md \
  --bibliography 07.references.json --citation-base main-paph2.manifest.json \
  --style template --font "Times New Roman" --east-asia-font "Times New Roman" \
  --line-numbers on --columns one --numbering-prefix S --force

# Word tracked-changes redline against a reviewed DOCX
docforge redline reviewed.docx fresh.docx -o fresh_tracked.docx

# LaTeX manuscript -> DOCX (main + optional SI + bib)
docforge tex2docx main.tex --si si.tex --bib ref.bib -o main.docx

# DOCX -> Markdown / LaTeX (Pandoc must be installed)
docforge docx2md reviewed.docx -o reviewed.md --track-changes accept
docforge docx2tex reviewed.docx -o reviewed.tex --track-changes accept

# DOCX -> mapped section Markdown with extracted media
docforge docx2md reviewed.docx --split-dir sections \
  --section-map section-map.json --force

# Generate figures from front-matter Markdown prompts (auditable sidecars)
docforge aigc figures/_prompts --backend litellm --all
docforge aigc figures/_prompts --dry-run

# Auditable TeX source ZIP from an explicit allowlist
docforge sourcepack . --allow main.tex --allow si.tex --glob-dir figures -o dist/tex_source.zip

# Validate a generated DOCX
docforge check output.docx --verify-clean
```

All paths are explicit; nothing is inferred from the current directory.

Template assembly replaces sample body content while retaining the supplied
template's styles, section geometry, headers, footers, numbering, and embedded assets.

Inline figures use ordinary Markdown plus a following caption paragraph, for example:
`![Figure 1](figures/structure.png)` followed by `Figure 1. Caption text.`.
When the template contains figure sections, docforge reuses their inline drawing
slots and the surrounding one-column/two-column section properties; the image
and caption remain separate Word paragraphs. Figure slots are detected from
each image paragraph and its following Figure/Scheme/Chart caption, even when
an SI template keeps sample prose and multiple figure pairs in one section.

Inline TeX formulae use ordinary Word runs rather than OMML, so variables are
italic, ``\mathrm{...}`` text and numerals are upright, and ``^``/``_`` become
true Word superscript/subscript run properties. For example,
``$t_{\mathrm{chem}}$`` renders as an italic *t* with an upright subscript.
A standalone display block delimited by ``$$`` is rendered as native OMML and
receives an automatic sequential number, e.g. ``$$`` / ``k = Ae^{-E_a/RT}`` /
``$$`` becomes equation (1). The tabulated number is retargeted to the active
body-column width when a template is used. Unicode chemical scripts in ordinary
text are converted to the same Word run properties during the final audit.

Lines beginning with `! ` or `！ ` are treated as source comments and omitted by
default; `--keep-comments` preserves them as TODO notes. An image alt-text option
such as `![Figure 1|columns=double](figure.png)` requests a two-column image
section. Without an image option, template mode uses the template's figure-slot
section (MolCrysKit is single-column for figures and two-column for body text),
while `--columns one|two` changes body sections. Continuous `<w:sectPr>`
properties are copied section by section, including their headers/footers and
column definitions.

For SI output, pass --numbering-prefix S to apply the S prefix to figure captions, explicit Table: captions, and display-equation numbers.

For a Supporting Information build, pass the main manifest with
`--citation-base`: citations present in the main document reuse their numeric
labels; keys that occur only in SI receive `S1`, `S2`, and so on. The main and SI
bibliographies therefore share a single ordered JSON source while keeping SI-only
entries visibly separate.
`--metadata` reads simple level-one Markdown fields (`TITLE`, `AUTHOR`,
`AFFILIATION`, `Acknowledgement`, and `Code Availability`). Citation tokens
(`\\cite{key}` and `\\citep{key}`) are converted to numbered references from
the insertion order of the JSON bibliography. Unknown keys fail the build.
Each template output receives adjacent `.manifest.json` and `.sha256` audit
files.

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

Reverse conversion is exposed as `docforge.markdown.docx_to_markdown` and
`docforge.tex.docx_to_tex`; both invoke the installed `pandoc` command. The
Markdown exporter can additionally split the converted document using an
ordered JSON section map and writes a `manifest.json` beside the section files.

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



## Template assembly arguments

Metadata keeps authors and contact lines separate. Use an EMAILS heading, or put
email lines in the AUTHOR section, with one marker-preserving address per line.
At most two contact lines are accepted, and docforge does not infer a
corresponding-author mapping.

The template assembly CLI accepts --font and --east-asia-font for explicit
generated-run font overrides. --style template preserves the styles discovered
from the supplied DOCX; alternatively --style PATH.json supplies a semantic
role-to-existing-style mapping such as {"body": "TA_Main_Text1"}. The
--line-numbers option is template, on, or off and applies to every retained
front, body, figure, and terminal section. --heading-before sets explicit
heading spacing before in points. --no-title removes level-one Markdown
headings, including the generated References/Acknowledgment/Code Availability
headings, while retaining the metadata title block. These rendering arguments
are written to the assembly manifest.

`--native-toc` inserts a native Word TOC field for heading levels 1–3 after the
template front matter. `--restart-heading-numbering` adds decimal H2 labels
that restart at 1 after every H1; H3 remains unnumbered. Generated headings are flush left; template
body-paragraph indentation remains unchanged unless `--body-first-line-chars`
sets an explicit first-line indent in character units.

The generic rule is level-based rather than filename-based: any input Markdown
file may contain level-one headings, and --no-title removes them uniformly.
References are generated from the ordered bibliography JSON; when
--citation-base points to the main manifest, citations shared with the main
document retain their numeric labels and SI-only entries receive S1, S2, and
so on.
