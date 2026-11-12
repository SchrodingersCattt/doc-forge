# Template roles and layout

`docforge md2docx source.md -o output.docx --template template.docx --metadata metadata.md` discovers semantic paragraph styles from the supplied template. Font family, size, and color inherit from those styles; reference paragraphs also inherit hanging indents and spacing. Sample run typography does not become direct formatting on generated text. Explicit typography flags remain available. A cover sample that would be replaced by a different unused style, or that carries direct formatting the selected style would drop, is reported in the assembly manifest. It is not copied unless prototype preservation is requested.

Some templates format their cover samples directly instead of defining named
styles. Pass `--preserve-metadata-prototypes` to reuse the detected cover
samples' paragraph style, alignment, font size, color, and emphasis for the
banner, title, authors, affiliation, and contact lines. Each generated line
uses the presentation of the longest text run in its sample paragraph. Sample text itself is
never copied; review markup and highlighting are discarded. The option affects
cover metadata only. Body and heading roles still use the named styles. It
intentionally takes precedence over the semantic style chosen for those
metadata paragraphs; an explicit `--font` or `--east-asia-font` overrides the
sample fonts, and a subsequent restyle profile can apply its own policy.

The same choice is available to Python callers as
`assemble_markdown_template(..., preserve_metadata_prototypes=True)` and is
recorded in the assembly manifest. Omit it when the template's named role
styles, rather than its sample paragraphs, should remain authoritative.

Named styles stay authoritative by default, including an unused built-in
`Title` style. When a detected cover sample uses a different paragraph style,
and the selected style is unused or the sample carries direct alignment, size,
emphasis, or color, assembly keeps the named style and records one
`rendering.metadata_style_conflicts` entry per affected role
(`supporting_information` for the banner, plus `title`, `authors`,
`affiliations`, and `contacts`). The same list is copied to `verification`.
The command also emits a `UserWarning`. Two documented overrides clear that
conflict: `--preserve-metadata-prototypes`, or an explicit `--style` mapping
that assigns the role a paragraph style (a legacy role entry or a v1
`word_style`). A sample that already uses the selected
style is not a conflict; direct formatting on that paragraph stays subordinate
to the style definition until prototype preservation is requested.

The option covers both shapes. A Normal paragraph with direct formatting is
reused only for its own cover line. A paragraph that already uses an explicit
role style such as `Title1`, `Authors`, or `Address` keeps that style, and
prototype preservation adds the sample's direct alignment and run formatting.
Unrelated body samples are not imported onto generated prose.

Some templates provide figure and table caption examples using direct
formatting, even though an unused built-in `Caption` style also exists. Pass
`--preserve-caption-prototypes` to use the paragraph style, alignment, spacing,
and neutral text-run formatting of the first numbered figure and table caption
examples independently. Sample text is never copied. Missing examples leave
the corresponding discovered role unchanged. An explicit semantic style
profile remains authoritative and an explicit `--caption-font-size`, `--font`,
or `--east-asia-font` overrides the sampled metrics. The Python equivalent is
`assemble_markdown_template(..., preserve_caption_prototypes=True)`; the choice
is recorded in the manifest as `rendering.preserve_caption_prototypes` and
`rendering.caption_prototype_roles`. The role list is independent for
`caption` and `table_caption`: a missing sample omits that role and leaves
its discovered style unchanged. Paragraph alignment, spacing, and indentation
come from the sample paragraph. Size, color, and font come from its longest
neutral text run. Authored bold, italic, subscript, and superscript in the
caption text are kept. An explicit role mapping, `--caption-font-size`,
`--font`, or `--east-asia-font` overrides the sample. Omitting the option
preserves the existing style discovery behavior. Cover prototype preservation
is a separate flag and does not change caption roles.

For a document without a template, `--standalone-style manuscript` selects a
neutral preset: black text, a left-aligned 11 pt body with 1.2 line spacing,
0 pt before and 6 pt after, and Markdown heading level N mapped to Word
Heading N. A level-one heading is Heading 1 at 14 pt, not Heading 2. A
supplied standalone title uses the Word Title style at 16 pt black, without
the built-in Title bottom border. The pre-existing standalone defaults remain
available as `--standalone-style legacy`, which is the default for backward
compatibility and still maps a Markdown level-one heading to Heading 2.
Template conversions should use the template's own styles instead.

## Semantic roles

Existing roles include `title`, `authors`, `affiliations`, `abstract`, `heading_1` through `heading_3`, `body`, `caption`, and `reference`. Additional roles are optional and fall back as shown:

| Role | Example style names | Fallback role |
| --- | --- | --- |
| `table_caption` | TableCaption, Table Caption | caption |
| `table_head` | TableHead, Table Header | body |
| `table_body` | TableBody, Table Body | body |
| `table_foot` | TableFoot, Table Note | body |
| `keywords` | Keywords | body |
| `back_matter_heading` | BackMatterHeading, Back Matter Heading | heading_1 |
| `back_matter_body` | BackMatter, Back Matter | body |
| `toc_text` | TableOfContentText, TOC Text, Graphical Abstract | body |
| `algorithm` | Algorithm, Pseudocode | body |

Style IDs and names are matched without case or punctuation differences. A legacy JSON map such as `{"table_caption": "My Table Caption"}` overrides discovery. A v1 profile supports style definitions and policies:

```json
{
  "schema": "docforge.style.v1",
  "template": "template.docx",
  "roles": {
    "body": {"word_style": "P1"},
    "table_caption": {"word_style": "TableCaption"}
  },
  "policy": {"direct_formatting": "clear_conflicting"}
}
```

Pass this file with `--style profile.json`. Relative profile-template paths resolve beside the profile. Semantic role names choose styles during generation. For `restyle`, `select` names an existing paragraph style to remap; otherwise paragraphs already using `word_style` are matched. New generation uses styles present in the supplied assembly template.

`clear_all` removes presentation formatting while preserving authored bold, italic, superscript/subscript, underline, strikeout, and character styles. It removes paragraph spacing, indentation, alignment, tabs, shading, and borders while retaining lists, sections, outline levels, and pagination flags. `clear_conflicting` clears only properties supplied by the target style chain, defaults, or role overrides, including partial font and paragraph properties. Explicit profile bold/italic overrides take precedence under `clear_conflicting`. `preserve` leaves direct formatting intact.

## Tables, keywords, and back matter

```markdown
# Abstract
An abstract paragraph.
Keywords: first · second

# Results
Table: Measurements.

| Item | Mean<sup>[a]</sup> | Error<sup>[b]</sup> |
| --- | --- | --- |
| Sample 1 | 2.4 | 0.3 |

Table note: [a] Mean value.
Table note: [b] Estimated uncertainty.

This is an ordinary paragraph with [a] and [b].
```

Keywords become a separate front-matter paragraph, including when no blank line precedes them. A `Table:` caption uses `table_caption`; the first row uses `table_head`, and other rows use `table_body`. `Table 5. Caption text` immediately before a table is the explicit authored caption: the number and caption role are kept, including in the assembly manifest (`tables[].number`, `tables[].label`, `tables[].caption`, `tables[].role` = `table_caption`, `tables[].numbering` = `authored`). `Table: Caption text` stays automatic (`numbering` = `generated`). The optional numbering prefix is the string supplied by the caller, such as `--numbering-prefix P`; it is not tied to a particular document type. Automatic numbering skips a label that an authored caption already uses and reports `table.number_collision`. Repeated authored labels are rejected. A numbered table mention that is not followed by a table remains ordinary prose.

`Table: Table 5. Caption text` is ambiguous: `Table:` would assign a new number in front of an authored number. Parsing fails with `table.caption_ambiguous` and does not write a caption that contains both numbers. Keep the number by writing `Table 5. Caption text` on the line immediately before the table, or drop the authored number and write `Table: Caption text`.

The same manifest records each figure's authored alt text in `figures[].alt` and the visible caption in `figures[].caption`. A paragraph that is a figure caption immediately after an image is the only visible caption. The image alt text is stored on the drawing (`wp:docPr/@descr` and, when present, `pic:cNvPr/@descr`) and is not printed again. With no following caption, a descriptive alt is the caption body; template assembly still applies the next figure number, and standalone conversion prints the alt text itself. Template assembly numbers Figure, Scheme, Chart, and Supplementary Fig. captions in independent sequences; `figures[].number` and `figures[].label` record the number and label from the appropriate sequence. The caller's numbering prefix applies to each sequence. An image without a recognized caption kind belongs to the Figure sequence, including when it has no visible caption. An empty alt or an alt that is only a media filename does not become a visible caption. An empty alt falls back to the visible caption for the drawing description when a caption exists, and stays empty when it does not.

Pipe-table separator cells specify alignment: `:---` is left, `:---:` is centered, and `---:` is right. An unmarked `---` does not write a direct alignment, so a template paragraph style keeps its own alignment. In a fresh standalone document an unmarked column defaults to left. Markers apply to every header and body cell in that column, including cells that contain emphasis, superscripts, or text of unequal length. The separator may use different dash runs (`:---`, `:-----`, `---:`); alignment follows the colons, not the dash count. A row whose cell count does not match the header is still rejected. Configure the `TableCaption` and `TableFoot` styles in the template to control first-line indentation and paragraph spacing.

The semantic Markdown adapter stores an explicit alignment on each cell. Rendering that document back to Markdown writes the same `:---`, `:---:`, and `---:` markers. Reading a DOCX keeps a direct `w:jc` of left, center, or right and leaves a cell with no direct alignment unmarked, so a template's inherited alignment is not turned into a pipe marker.

### Table notes

Table notes are body paragraphs associated with the immediately preceding Markdown table. They are **not** Word page footnotes and are never converted into `w:footnote` references.

Use one declaration per note, immediately after the table or after another note for that table. Blank lines may separate declarations. An unrelated paragraph, heading, or other block ends the association window.

```markdown
Table note: [a] Mean value.
Table foot: [b] Estimated uncertainty.
```

`--table-notes explicit` (Python `table_notes="explicit"`) recognizes only these declarations. An ordinary unmarked paragraph that contains `[a]` or `[b]`, including one immediately after a table, stays ordinary prose. Prefix a directive with a backslash to keep it literal: `\Table note: not a note`.

`--table-notes legacy` (default) preserves historical recognition: `Table note:` / `Table foot:` always become notes, and a line that begins with `[a]` immediately after a table or another note is also treated as a note. Existing consumers keep that default so behavior does not change silently. Prefer explicit mode for new manuscripts; migrate by rewriting legacy `[a]` lines as `Table note: [a] ...` and passing `--table-notes explicit`.

Within one table scope, matching `<sup>[a]</sup>` cell markers to `Table note: [a] ...` declarations is diagnostic only. A superscript without a matching declaration remains authored superscript text. Labels may repeat across separate tables. `--table-note-diagnostics {warn,error,ignore}` controls orphan declarations, missing cell labels, and duplicate labels in the same table (`warn` default).

Standalone output uses a built-in `TableFoot` style (smaller left-aligned text). Template assembly maps the `table_foot` role to `TableFoot` / `Table Foot` / `Table Note`, falling back to body when absent. Notes keep with their table through Word `keepNext` chaining; long tables or notes may still paginate without an impossible keep-together constraint. Authored bold, italic, subscript, superscript, mathematics, and links are preserved.

Only repeated labels under grouping headers (`level`, `group`, `category`, `class`, `type`, `family`, `site`, `site level`, `factor`, `block`, or a header ending in `level`) can merge vertically. Repeated measurements in other columns stay visible. Blank cells break a run, so two equal labels with an empty cell between them stay separate. Markdown and TeX share this column policy. The default mode is `auto`, which preserves that behavior for existing callers.

`--table-grouping off` (Python `table_grouping="off"`) keeps every Markdown cell as its own DOCX cell: no inferred `w:vMerge`, and each row keeps the text the author wrote. `auto` remains available as an explicit mode. A caption marker or a directive overrides the default for one table:

```markdown
Table grouping: off
Table: Measurements.|grouping=off

| Group | Type | Value |
| --- | --- | --- |
| A | Alpha | 1 |
| A | Alpha | 1 |
```

`grouping=auto` on one table opts back in when the default is `off`. TeX conversion stays in automatic grouping.

### Page span

Generated tables are fitted to the active text column. That remains the default, including for a table with many columns. `--table-span page` (Python `table_span="page"`, or the same key in a delivery profile) asks for the printable page width instead. A per-table marker overrides the default and is never inferred from the column count:

```markdown
Table span: page
Table: Measurements.|widths=proportional

| A | B | C | D |
| --- | --- | --- | --- |
| 1 | 2 | 3 | 4 |

Table note: Units are arbitrary.
```

`span=column` keeps the current column when the default is `page`. `widths=fixed` (the default) scales the generated grid's existing proportions into a fixed-width table of the target width. `widths=proportional` weights columns by the longest rendered cell, then writes those shares as fixed twips of the same target width. Both fill either the column or the printable page.

In a multi-column section, `span=page` places the caption, the table, and the notes explicitly associated with that table in one continuous full-width section, then restores the previous column count. A one-column section is already full width, so a page span does not add section breaks. The first row remains a repeating header. Tables taller than one page flow across pages. Headers, footers, line numbering, and the surrounding text stay with the section properties they had. `--table-span` requires template assembly, as `--figure-span` does. Standalone documents are already a single column; per-table span markers are accepted there and do not insert breaks.

Acknowledgments/acknowledgements, conflict of interest, author contributions, data availability, code availability, data and software availability, and supporting information headings select back-matter roles. Metadata-generated acknowledgments and availability statements use these roles too.

## Math and algorithms

`--inline-math omml` compiles inline `$...$` to native `m:oMath` using the same strict Pandoc compiler as display equations. Nested scripts, operator limits, and math alphabets stay structured. The compatible default `--inline-math runs` retains the historical text representation. The option currently requires template assembly. Supported Pandoc versions are listed in [migration notes](migration.md). Conversion diagnostics fail the build rather than silently flattening math.

Use an explicit fenced block for pseudocode:

````markdown
```algorithm
**Require:** $n > 0$
    while $i < n$
        return $i$
```
````

`pseudocode` is an alias. Each source line becomes an `algorithm` paragraph. Spaces, tabs expanded to four spaces, and nonbreaking spaces determine indentation relative to the style's left indent and font size. Blank lines remain paragraphs. A named template style controls typography; the body-style fallback uses monospace text. Add `--inline-math omml` for native equations within algorithm lines.

Paired bare absolute-value bars in math containers become delimiter objects for Word/LibreOffice interoperability. A lone conditional or set-builder bar remains an operator. For ambiguous multi-bar expressions, write explicit `\left|...\right|` delimiters.

## Pagination and contents

Headings and table captions keep with the following paragraph; image paragraphs keep with their captions. Images fit both the active section width and text height, with caption space estimated from text length and font size. Aspect ratio is preserved, including template figure slots, column figures, and page-spanning figures. Caption estimates are conservative layout guidance; inspect unusually long captions in the target renderer.

`--heading-before 6` sets 6 pt before every generated heading level. To set a value for one level only, leave this option unset and configure the corresponding heading style in the supplied template.

`--end-matter-placement end` places level-one `Table of Contents`, `TOC`, and `Graphical Abstract` source sections after references, starting on a new page. The default `source` preserves source position. The option also places a requested native TOC at the end.

`--native-toc` inserts a native Word TOC field. Its cached result contains the emitted heading text immediately, without invented page numbers; `w:updateFields` requests a refresh on open. Word computes page numbers when it updates the field. Automatic update behavior in other viewers varies. Without a native TOC request, the generator disables automatic field updates as before.

The Python options are `inline_math="omml"`, `end_matter_placement="end"`, `table_notes="explicit"`, `table_note_diagnostics="error"`, `table_span="page"`, `table_grouping="off"`, and `native_toc=True` on `assemble_markdown_template`. `render_blocks_to_doc(..., table_grouping="off")` is the standalone equivalent for grouping. Delivery profiles accept `table_span` and `table_grouping` with the other rendering options. Output manifests record the selected settings, and each captioned table record includes its resolved `span` together with `role` and `numbering`. The renderer uses UTF-8 text I/O and closes temporary file/package handles before replacement for Windows compatibility.

## Assembly sidecars

`md2docx --template` writes two audit files for the finished DOCX: a JSON manifest named `<stem>.manifest.json` and a SHA-256 sidecar named `<stem>.sha256`. By default both sit beside the DOCX. The manifest `output.path` is the absolute path of that DOCX, and `output.sha256` is the digest of the bytes on disk at publication time. The sidecar is one line, `<digest><two spaces><name>`. When the sidecar is beside the DOCX, `<name>` is the filename. Otherwise `<name>` is a `/`-separated path relative to the sidecar, or a `/`-separated absolute path when the files are on different Windows drives or UNC shares. `verify_assembly_checksum` resolves that name and rejects a digest that does not match the file.

`--manifest-dir DIR` stores only the manifest in `DIR`. The checksum stays beside the DOCX. This flag does not collect every audit file.

`--sidecar-dir DIR` stores both the manifest and the checksum in `DIR`, including when `DIR` is nested and does not yet exist. That is the switch for keeping the DOCX separate from every assembly audit file. `--manifest-dir` still overrides only the manifest location, and `--checksum-path FILE` still overrides only the checksum file, so the two files can be sent to different directories. All three flags require `--template`. Existing sidecar files are replaced as one pair. If publication fails, the previous pair is left in place: a new manifest is not written beside an old checksum.

The Python entry point is `write_assembly_sidecars(..., manifest_dir=..., sidecar_dir=..., checksum_path=...)`. `docforge contract` still publishes its manifest and checksum beside the DOCX; it does not take these destination flags.

Any tool that rewrites the DOCX after the build, including a word processor refreshing fields, makes both the manifest digest and the checksum stale. Rehash with `refresh_assembly_sidecars(output, manifest_path=..., checksum_path=...)`. It reads the finished file once, updates `output.path`, `output.sha256`, and `output.size_bytes`, and rewrites the checksum from that same digest. `generated_at_utc` stays at the assembly time. Do not edit one sidecar by hand. A failed refresh leaves the previous pair unchanged.
