# Migration notes for the semantic hub

Issue 50 is the acceptance ledger. This slice adds the hub and keeps the historical Markdown and TeX DOCX renderers.

## API

| Call | Behavior |
| --- | --- |
| `docforge.md2docx` | Unchanged command. Implementation lives in `docforge.adapters.legacy`. |
| `docforge.tex2docx` | Unchanged command. Same legacy builder. |
| `convert(md, docx)` / `convert(tex, docx)` | Calls those builders. Optional `style=` restyles the result. |
| `convert(docx, docx)` | `EXACT_REUSE` when no style is requested. With a profile, `SOURCE_PATCH`. |
| `convert(md, tex)` and `convert(tex, md)` | Supported subset through `SemanticDocument`. |
| `convert(docx, md)` and `convert(docx, tex)` | Existing Pandoc exporters. `fidelity=preserve` refuses the pair. |
| `restyle` | Edits styles in the DOCX package and checks that paragraph text is unchanged. `StyleProfile.template`, when set, is the style source for the named styles. |
| `write(..., fidelity="preserve")` on an edited DOCX | `SOURCE_PATCH` for a supported body paragraph or heading replace, insert, or delete. Any other edit is refused and the destination is left untouched. |
| `audit` | ZIP integrity and relationship targets, including header and footer references. |
| `export` / `build` | `preserve` wraps `md-export` / `md-build`. `semantic` writes or reads one Markdown file. |

`read()` always returns a `SemanticDocument`. Absolute paths are not part of `to_dict()` or the semantic hash.

`loss_policy=error` rejects a diagnostic when its severity is `error` or its category is `loss`, and it does so before an output file is replaced. Informational warnings are not losses. `fidelity=preserve` escalates loss warnings to errors. A preserve-mode write of an unchanged DOCX still checks that the file on disk matches `SourceContext.sha256`. A preserve-mode write of an edited DOCX patches the package or refuses the edit. It does not rebuild.

`restyle` audits the staged package and replaces the destination only when that audit passes. `StyleProfile.template` is a DOCX whose named styles, including their `basedOn` chain, are copied in before role font and paragraph overrides. A missing, non-DOCX, or styles-less template raises `STYLE_TEMPLATE` and does not replace the destination. A relative template path is resolved next to the profile file. A role with `select` assigns its Word style to paragraphs that already use the named style. A role that matches no paragraph is returned as unused.

Supported `SOURCE_PATCH` content edits are confined to simple body paragraphs and headings: replace their text (plain, emphasis, or strong), insert a plain paragraph, or delete one. The zip delta is `word/document.xml` only. Tables, figures, fields, math, revisions, drawings, bookmarks, and story parts are refused with `UNSUPPORTED_PATCH`. Anchors are `word/document.xml:child:N` on body blocks. Keeping that anchor replaces the paragraph in place. Dropping it is an insert plus a deletion of the old block.

## Style authority

`md2docx` and `tex2docx` without a StyleProfile keep their historical direct font, size, and color passes, including the template assembly stamps of 11 pt body, 10 pt captions, 18 pt unsized titles, and black runs.

`assemble_markdown_template(..., style_authority=profile)` skips those implicit stamps and then restyles the package with the profile. `convert(..., style=profile)` and `write(..., style=profile)` drop the same direct defaults from the freshly rebuilt DOCX before restyle. An existing document passed only to `restyle` keeps its direct formatting except where the profile policy clears it.

## Intentionally unchanged

Legacy `md2docx` and `tex2docx` stay, including when no StyleProfile is passed.

`replace_report_and_resume` keeps its previous default markers. Callers can pass `start_marker` and `end_markers`.

Comment ids in a reversible bundle still keep `cN` as Word id `N`. A non-numeric id now skips every `cN` already present in the Markdown, including when that note appears first. The previous scanner looked for backspace characters, so it never saw those numbers.

## Checks

`pytest` is the executable check. It does not render pages.

`tests/test_contract.py::test_contract_accepts_dotx_and_publishes_final_output_descriptor` checks the package and manifest with rendering turned off, so it runs without Word or LibreOffice.

`tests/test_contract.py::test_contract_runs_redline_render_and_consistent_manifest` is the LibreOffice render gate inside contract publication. When `soffice` is absent the test is skipped. A skip is not a pass.

`tests/test_render_gates.py` calls `docforge.gates.render.layout_gate`. `status=skipped` does not count as a pass (`counts_as_pass` is false). Word visual parity is `passed` only when the resolved engine is Word and Word wrote a PDF. LibreOffice can pass its own render gate while `visual_parity` stays `not-passed`. On this runner Word and `soffice` were both absent, so both gates were skipped and Word visual parity is not passed.

Run them where the engines exist:

```bash
pytest tests/test_render_gates.py -q
pytest tests/test_contract.py::test_contract_runs_redline_render_and_consistent_manifest -q
```

`python -m docforge.bench --iterations 3` is the controlled bench. Numbers from this runner are in `docs/bench.md`. No measured operation rose by 25% in wall time or peak RSS. Package sizes were unchanged.

Pandoc is required for display-math OMML and for DOCX export tests that invoke it. PyMuPDF is required only for the PDF-figure regression; without it that test is skipped and a PDF figure raises instead of becoming a missing-figure placeholder.

## Pair corpus

Slice toward epic #50. This note does not close the epic.

Synthetic fixtures live in `tests/fixtures/pairs/`. `pytest tests/test_pair_corpus.py` runs the six-pair matrix. `pytest` runs that file with the rest of the suite. Constructing and validating a `SemanticDocument` still does not import Pandoc, Word, or LibreOffice.

| Pair | What this slice locks |
| --- | --- |
| MD → DOCX | API and CLI on `layout.md`. Heading text, emphasis, subscript, inline math text, table grid, figure SHA-256, and list item text. `body-style.yaml` keeps paragraph text and sets the body size. |
| DOCX → MD | Semantic read of `native.docx`, plus `export --fidelity preserve` for the reversible bundle. Readable Pandoc Markdown when Pandoc is installed. |
| TeX → DOCX | API and CLI on `layout.tex` with `refs.bib`. Section, emphasis, table, figure bytes, and legacy citation and cross-reference numbers. |
| DOCX → TeX | Pandoc LaTeX for `native.docx` when Pandoc is installed. Unsupported text boxes stay on the loss report. |
| MD → TeX | `semantic.md` matches checked-in `semantic.tex` and the shared semantic view, including math, citations, and cross-references. |
| TeX → MD | That same view, then back to TeX. |

`loss_policy=error` refuses a missing image (`resource.missing`), an unknown TeX command or environment (`latex.unsupported_command`, `latex.unsupported_environment`), and a DOCX text box (`docx.textboxes`). The destination is left untouched and the CLI exits 2. Sources are not modified.

A TeX `tabular` column specification is not a cell. Display `align` is written as `align` with `amsmath`, not nested inside `equation`. `pdflatex` or `latexmk` compiles that fixture when one is installed; a missing engine skips the compile check, and that skip is not a pass. `\label` and `\ref` round-trip in the semantic inline codec. A `\label` inside `figure` or `table` is reported as `latex.float_label` and is not a semantic node. The legacy TeX renderer still resolves those labels when the loss policy is `warn`.

Legacy Markdown `#` is Word Heading 2. A TeX `\section` is Heading 1. Legacy lists are paragraphs. The DOCX reader does not declare list or citation nodes. OMML is `docx.omml`, not TeX.

Pandoc goldens are `tests/fixtures/pairs/goldens/`. Exact normalized text is compared when the installed Pandoc matches `pandoc-version.txt` (3.1.11.1). Another installed version skips that byte comparison after the structural checks. No Pandoc skips the golden tests. Display-math DOCX cases skip unless Pandoc is 2.9.2.1, 3.1.11.1, or 3.9.0.1. A skip is not a pass.

Markdown ↔ TeX round-trips compare the semantic projection. Round-trips through DOCX compare heading text, emphasis, and table cells. They do not require DOCX byte identity.

## This slice and what remains

Slice toward epic #50. This note is not the acceptance gate for that epic.

Matrix boxes exercised here:

| Box | Evidence |
| --- | --- |
| B. SOURCE_PATCH content edits | Paragraph and heading replace, insert, and delete. `word/document.xml` is the only changed part. Media bytes stay equal. |
| B. No silent reconstruction | `fidelity=preserve` on an unsupported edit raises `UNSUPPORTED_PATCH` and leaves the destination untouched. Semantic mode can still rebuild. |
| StyleEngine template source | `StyleProfile.template` copies named styles. An invalid template fails closed. Role font overrides still win. |
| StyleEngine authority | With a StyleProfile, historical 11 pt / 10 pt / 18 pt / black stamps do not win. Without one, `md2docx` keeps them. |
| D. Render evidence | LibreOffice and Word gates report `skipped` or `passed`. Word visual parity is not passed on this runner. |
| D. Bench | `docs/bench.md`. No measured case rose by 25%. |
| RT-CONTENT | Patched paragraph text changes; sibling paragraph properties stay. |
| RT-STYLE | Template definitions and role overlays; profile authority over historical stamps. |
| RT-REL / RT-MEDIA | Patch leaves relationships and media parts byte-identical. |
| RT-UNSUPPORTED | Tables, figures, opaque blocks, and other unsupported edits are refused. |
| RT-EXACT | Unchanged DOCX copy is still byte-identical. |

Still open for the epic ledger:

- Word visual parity. Word was unavailable here, so it is not passed.
- A LibreOffice page render on a machine with `soffice`. The skip on this runner is not a pass.
- `SOURCE_PATCH` for tables, figures, fields, math, revisions, comments, headers, and footers.
- Legacy `md2docx` and `tex2docx` remain the DOCX builders until a deliberate deprecation.
- A `\label` inside `figure` or `table` is still the diagnostic `latex.float_label`. It is not a semantic node.
- Pixel-identical pagination across Word, LibreOffice, and LaTeX.
