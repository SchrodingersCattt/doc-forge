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

`md2docx --template` inherits fonts, sizes, and colors from named template styles. Explicit typography flags override them. Standalone Markdown and TeX rendering retain their defaults. See [template roles and layout](template-layout.md).

`assemble_markdown_template(..., style_authority=profile)` applies the profile after assembly; the CLI accepts the same v1 schema through `--style`. `convert(..., style=profile)` and `write(..., style=profile)` drop the same direct defaults from the freshly rebuilt DOCX before restyle. For existing documents, `preserve` retains direct formatting. `clear_all` removes run presentation and paragraph spacing, indents, tabs, and alignment while retaining bold, italic, scripts, character styles, and structural properties. `clear_conflicting` clears properties supplied by the target style, its base styles, document defaults, or explicit role overrides; emphasis remains unless the role explicitly overrides it. Nested hyperlink runs are included.

## Cover and caption prototypes (#65, #69)

Named styles stay authoritative unless the caller opts in. `--preserve-metadata-prototypes` keeps detected cover samples, including Normal paragraphs with direct formatting and paragraphs that already use a named role style. `--preserve-caption-prototypes` keeps discovered figure and table caption samples independently, including paragraph formatting and the longest neutral run. The two flags do not change each other. `--standalone-style manuscript` is the neutral preset for a conversion without a template. `legacy` remains the default, so a Markdown level-one heading still becomes Heading 2 unless that preset is selected.

When a detected cover sample uses a different paragraph style, and the selected style is unused or the sample carries direct alignment, size, emphasis, or color, the conflict is recorded in `rendering.metadata_style_conflicts` (and copied under `verification`) and emitted as a `UserWarning`. The DOCX still uses the named style, so existing documents stay stable. Clear the conflict with `--preserve-metadata-prototypes` or by assigning that role a paragraph style in `--style`. `rendering.caption_prototype_roles` lists the caption roles that adopted a sample. See [template roles and layout](template-layout.md).

## Table numbers (#66)

`Table: Measurements.` is still numbered automatically. To keep an authored number, put `Table 5. Measurements.` on the line immediately before the table. That caption keeps its number and the `table_caption` role in the DOCX and in the assembly manifest (`tables[].numbering` is `authored` or `generated`). `--numbering-prefix` is concatenated with the generated number exactly as supplied.

`Table: Table 5. Measurements.` no longer renders as `Table 1. Table 5. Measurements.`. It fails with `table.caption_ambiguous` before an output file is published. Use the bare numbered line, or remove the authored number. If an automatic number would reuse an authored label, the generator skips that label and warns `table.number_collision`. Duplicate authored labels still fail.

The Markdown semantic adapter writes an automatic caption back as `Table: ...` and an authored caption back as `Table 5. ...`, so reading that Markdown again restores the caption role.

## Image alt text (#68)

Standalone `md2docx` used to print image alt text as a caption paragraph and also print a following caption. The following caption is now the only visible caption, as in template assembly. The alt text is stored on the drawing description (`wp:docPr/@descr` and `pic:cNvPr/@descr`).

An image with no following caption still uses a descriptive alt as its visible caption. Standalone conversion prints that alt text. Template assembly uses it as the caption body and applies the next figure number, which is unchanged. An empty alt does not create a caption. An alt that is only a filename is not printed; add an explicit caption paragraph when that text should be visible. `add_image(doc, path, caption)` still prints `caption`. Pass `alt=` when the accessibility text differs from the visible caption. `figures[].alt` in the assembly manifest is the authored alt text.

## Table layout (#67, #70, #71)

Pipe-table alignment markers (`:---`, `:---:`, `---:`) are written on header and body cells. An unmarked separator does not replace a template style's alignment. Standalone documents still default an unmarked column to left. The Markdown semantic adapter and a DOCX read both keep explicit left, center, and right alignment; a cell without direct `w:jc` stays unmarked.

`--table-span` defaults to `column`. Tables keep fitting the active column, however many columns they contain. `page`, or a per-table `Table span: page` / `|span=page` marker, uses the printable page width. In a multi-column section the caption, table, and its associated notes share one continuous full-width section and the previous column count is restored. A one-column section does not gain extra breaks. `--table-span` requires `--template`. `widths=fixed` (default) and `widths=proportional` both fill that target width.

`--table-grouping` defaults to `auto`. Repeated labels under the documented grouping headers still merge, and repeated values in other columns do not. `off`, or a per-table `grouping=off` marker, leaves every Markdown cell as its own DOCX cell. Existing callers that omit the option keep automatic grouping. TeX conversion stays on that automatic policy.

## Table notes (#63)

`--table-notes` defaults to `legacy` so existing manuscripts keep historical recognition: `Table note:` / `Table foot:` lines and post-table `[a]` heuristics. Opt into `--table-notes explicit` for declaration-only notes that stay scoped to the preceding table. Escape a literal directive with a leading backslash. Table notes remain body paragraphs; they are not Word page footnotes. See [template roles and layout](template-layout.md).

## Assembly sidecars (#72)

`md2docx --template` still writes `<stem>.manifest.json` and `<stem>.sha256` beside the DOCX when no destination flag is set. `--manifest-dir DIR` still moves only the manifest; the checksum stays beside the DOCX. `--sidecar-dir DIR` opts in to storing both sidecars in `DIR`. `--checksum-path FILE` opts in to a separate checksum file. The same three arguments are `manifest_dir`, `sidecar_dir`, and `checksum_path` on `write_assembly_sidecars`.

The checksum is the SHA-256 of the finished DOCX bytes, not of an intermediate package. After any external program rewrites that DOCX, call `refresh_assembly_sidecars` so the manifest digest and the checksum are replaced together. `docforge contract` still publishes both sidecars beside its DOCX. See [template roles and layout](template-layout.md).

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
| DOCX → TeX | Pandoc LaTeX for `native.docx` when Pandoc is installed. Text-box geometry stays on the loss report; Pandoc may project its text into the document flow. |
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
| StyleEngine authority | Template-backed Markdown inherits typography; explicit flags and profile overlays can override it. |
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
