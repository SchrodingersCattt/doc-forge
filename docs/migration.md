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
| `restyle` | Edits styles in the DOCX package and checks that paragraph text is unchanged. |
| `audit` | ZIP integrity and relationship targets, including header and footer references. |
| `export` / `build` | `preserve` wraps `md-export` / `md-build`. `semantic` writes or reads one Markdown file. |

`read()` always returns a `SemanticDocument`. Absolute paths are not part of `to_dict()` or the semantic hash.

`loss_policy=error` rejects a diagnostic when its severity is `error` or its category is `loss`, and it does so before an output file is replaced. Informational warnings are not losses. `fidelity=preserve` escalates loss warnings to errors. A preserve-mode write of a DOCX refuses an edited semantic tree, and it refuses an unchanged tree when the file on disk no longer matches `SourceContext.sha256`.

`restyle` audits the staged package and replaces the destination only when that audit passes. `StyleProfile.template` is rejected there until a template is actually applied. A role with `select` assigns its Word style to paragraphs that already use the named style. A role that matches no paragraph is returned as unused.

## Intentionally unchanged

Template assembly still applies its historical font, size, and color passes. Pass an explicit StyleProfile to `restyle` or `convert --style` when those passes should not be the authority.

`replace_report_and_resume` keeps its previous default markers. Callers can pass `start_marker` and `end_markers`.

Comment ids in a reversible bundle still keep `cN` as Word id `N`. A non-numeric id now skips every `cN` already present in the Markdown, including when that note appears first. The previous scanner looked for backspace characters, so it never saw those numbers.

## Checks

`pytest` is the executable check. It does not render pages.

`tests/test_contract.py::test_contract_accepts_dotx_and_publishes_final_output_descriptor` checks the package and manifest with rendering turned off, so it runs without Word or LibreOffice.

`tests/test_contract.py::test_contract_runs_redline_render_and_consistent_manifest` is the LibreOffice render gate. When `soffice` is absent the test is skipped. A skip is not a pass; that gate stays open until the test runs where LibreOffice is installed. There is no Word layout gate in this repository.

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

## Still open

Deferred after the pair-corpus slice:

- Word and LibreOffice page renders. The LibreOffice gate above still skips when `soffice` is absent, and that skip is not a pass. There is still no Word layout gate.
- `StyleProfile.template` is still rejected.
- `SOURCE_PATCH` content edits beyond restyle and exact reuse.
- Performance numbers from a controlled long-running bench.
- Template assembly still applies its historical font, size, and color passes.
- Legacy `md2docx` and `tex2docx` remain the DOCX builders.
- Float labels are diagnostics, not semantic nodes.
- The acceptance ledger for the epic stays open.
