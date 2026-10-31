# docforge

Build a Word document from Markdown or LaTeX. A build can write a sha256 sidecar beside the file, and `redline` can record the edits as Word tracked changes.

## Install

```bash
pip install -e .            # core (python-docx, lxml, Pillow)
pip install -e ".[plot]"    # matplotlib
pip install -e ".[aigc]"    # OpenAI-compatible and Google image clients
pip install -e ".[tex]"     # numpy, used when rendering TeX
```

Python 3.10+.

`md2docx` and `tex2docx` call Pandoc to turn display mathematics into Word OMML. Supported Pandoc versions for that step are 2.9.2.1, 3.1.11.1, and 3.9.0.1. `docx2md` and `docx2tex` also call `pandoc`. Install Pandoc and leave it on `PATH`.

## Commands

Details are in `docforge <command> --help`.

- `md2docx` — Markdown to DOCX
- `tex2docx` — LaTeX to DOCX, with an optional second TeX file and a BibTeX file
- `docx2md` — DOCX to Markdown via Pandoc
- `docx2tex` — DOCX to LaTeX via Pandoc
- `redline` — a DOCX of Word tracked changes against a reviewed file
- `md-export`, `md-build`, `md-roundtrip` — a reversible Markdown bundle; see [docs/docxmd-syntax.md](docs/docxmd-syntax.md)
- `roundtrip` — copy an unchanged DOCX, or patch its package; see [docs/roundtrip-contract.md](docs/roundtrip-contract.md)
- `apply-delta` — apply matched paragraph edits onto a source DOCX
- `export-block-map` — write paragraph indices and opaque block ids
- `contract` — build one template DOCX and its audit files
- `deliver` — build the DOCX files listed in a `docforge.delivery.v1` profile; see `docforge deliver --help`
- `aigc` — images from Markdown prompt files
- `sourcepack` — a TeX source ZIP from an allowlist
- `check` — check a DOCX
- `gate abbr`, `gate references`, `gate style` — source checks from a JSON config
- `pipeline` — run a `docforge.pipeline.v1` spec; see [docs/pipeline.md](docs/pipeline.md)
- `convert`, `restyle`, `audit` — semantic conversion, DOCX restyle, and package audit; see [docs/migration.md](docs/migration.md)

`md2docx` and `tex2docx` stay the renderers for those pairs. `convert` calls them for Markdown or LaTeX to DOCX, and `restyle` edits an existing DOCX package from a `docforge.style.v1` profile.

```bash
docforge md2docx main.md -o main.docx --template template.docx \
  --metadata metadata.md --bibliography references.json

docforge tex2docx main.tex --si si.tex --bib references.bib -o main.docx

docforge docx2md reviewed.docx -o reviewed.md --track-changes accept

docforge redline reviewed.docx current.docx -o current_tracked.docx

docforge md-export reviewed.docx md --id main --split split.json --bibliography references.json
docforge md-build md --id main -o revised.docx

docforge aigc figures/_prompts --backend litellm --all
docforge aigc figures/_prompts --dry-run

docforge sourcepack . --allow main.tex --allow si.tex --glob-dir figures -o dist/tex_source.zip

docforge check output.docx --verify-clean

docforge contract main.md -o main.docx --template template.docx \
  --metadata metadata.md --bibliography references.json --force

docforge gate abbr main.tex si.tex --config gates.json --project paper
docforge gate references --root . --config reference-gate.json

docforge pipeline pipeline.json --workspace . --output-dir dist
```

## Templates, math, figures, and citations

Template assembly replaces sample content while retaining styles, section geometry, headers, footers, numbering, and reusable figure layouts. Generated text inherits template fonts, sizes, and colors. The explicit `--body-font-size`, `--abstract-font-size`, `--caption-font-size`, `--reference-font-size`, `--font-family`, and `--east-asia-font` options override those values. Inline code retains its monospace font.

`--style profile.json` accepts either the legacy semantic-role-to-style map or a `docforge.style.v1` profile. See [template roles and layout](docs/template-layout.md) for roles, profiles, table notes, algorithms, and end placement.

A line that begins with `! ` or `！ ` is omitted. `--keep-comments` keeps that line as a `[TODO: ...]` note.

Inline math such as `$t_{\mathrm{chem}}$` becomes Word runs: variables are italic, `\mathrm{...}` text and numerals are upright, and `^` / `_` are superscript or subscript. With `--template`, add `--inline-math omml` to compile inline formulae to native Office Math, preserving nested scripts, limits, and math alphabets. This requires a supported Pandoc version. The default `--inline-math runs` preserves the historical text representation. A display block between `$$` lines is OMML and receives the next equation number, such as `(1)`. `--numbering-prefix S` applies that prefix to figure captions, `Table:` captions, and those equation numbers.

A figure is an image paragraph plus the caption paragraph that follows:

```markdown
![Figure 1](figures/structure.png)

Figure 1. Caption text.
```

`--figure-span column|page` sets the default width. Figures also fit within the active section text height, reserving estimated caption space and preserving aspect ratio. `span=` on the image overrides it, as in `![Figure 1|span=page](figure.png)`. `span=column` stays in the current column. In a two-column body, `span=page` puts the figure and caption in a temporary one-column section. `orientation=portrait|landscape` is separate, and landscape requires `span=page`.

`--citation-base` reads numbers from an existing manifest. A key already in that manifest keeps its number; a key that appears only here is numbered `S1`, `S2`, and so on. The reference list then contains the new entries. `--bibliography-scope all` repeats the shared entries as well. `\cite{key}` and `\citep{key}` use the JSON bibliography. An unknown key fails the build. The DOCX is written with `.manifest.json` and `.sha256` beside it.

`--metadata` reads these level-one fields: `TITLE`, `AUTHOR`, `AFFILIATION`, `Acknowledgement`, and `Code Availability`.

## Bibliography and gates

`docforge.bibliography` loads JSON or BibTeX into `BibliographyEntry` records, resolves citations, and formats them with the `markdown` or `plain` profile. Numbering is `first-citation` or `source-order`. A later build can inherit the number map from a manifest. The Markdown and TeX adapters use that resolver.

Abbreviation lists stay in configuration:

```json
{
  "abbr-whitelist": ["AI", "MLIP"],
  "projects": {
    "paper": {"abbr-whitelist": ["OOD", "SI"]}
  }
}
```

`docforge gate abbr` reports an abbreviation that is undefined or used once. `docforge gate style` checks the rules in its config and can compare them with a baseline. `docforge gate references` checks TeX figure and table labels from its config.

## Library layout

| Module | Purpose |
| --- | --- |
| `docforge.markdown` | Markdown blocks to DOCX: headings, lists, tables, code, images, inline math, and OMML display equations. Can fill a template. |
| `docforge.docxdiff` | Tracked-changes DOCX (`w:ins`, `w:del`, `w:move*`, `w:rPrChange`), including existing review comments. |
| `docforge.tex` | A LaTeX subset to DOCX: sections, floats, tables, algorithms, math, and BibTeX. A second TeX file can supply cross-references. |
| `docforge.aigc` | Prompt files to images (OpenAI-compatible or Google), with a JSON sidecar. |
| `docforge.sourcepack` | An allowlisted TeX ZIP that includes `SOURCE_PACKAGE_MANIFEST.csv`. |
| `docforge.plotting` | Matplotlib helpers for figure size, colors, and PDF, SVG, or PNG export. |
| `docforge.cli` | The `docforge` command. Each subcommand imports its module when invoked. |

`docx2md` and `docx2tex` call the installed `pandoc`. With `--split-dir`, `docx2md` can also write one Markdown file per section and a `manifest.json`.

## Audit files

- `aigc` writes a `.json` sidecar next to each image. The sidecar stores the prompt file's sha256, provider metadata, and the image sha256.
- `sourcepack` stores `SOURCE_PACKAGE_MANIFEST.csv` in the ZIP, with columns `Path`, `Size_bytes`, and `SHA256`.
- `redline` checks the tracked package: relationship ids stay in that package, and accepting the revisions is compared with the current document.
- `check --verify-clean` rejects leftover comments and unaccepted revisions. Template assembly runs the same check. Markdown text converts an em dash to an en dash; `gate style` can flag an em dash or a hyphen used as a numeric range.

## Development

```bash
pip install -e ".[tests]"
pytest
```

Public API, CLI, contract, migration, and ADR changes update the stable docs in the same change. The review bar is in [AGENTS.md](AGENTS.md). The pull request checklist is [`.github/pull_request_template.md`](.github/pull_request_template.md).

## License

MIT — see [LICENSE](LICENSE).
