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

Pandoc is required for display-math and DOCX export tests that invoke it. PyMuPDF is required only for the PDF-figure regression; without it that test is skipped and a PDF figure raises instead of becoming a missing-figure placeholder.

## Still open on issue 50

Word and LibreOffice render evidence, the full six-pair fixture matrix with Pandoc goldens, source-patch edits beyond restyle and exact reuse, and performance numbers from a controlled long-running bench. Core model tests do not need Pandoc or Word. The LibreOffice render gate above is part of that open set.
