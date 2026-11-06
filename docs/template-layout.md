# Template roles and layout

`docforge md2docx source.md -o output.docx --template template.docx --metadata metadata.md` discovers semantic paragraph styles from the supplied template. Font family, size, and color inherit from those styles; reference paragraphs also inherit hanging indents and spacing. Sample run typography does not become direct formatting on generated text. Explicit typography flags remain available.

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

Some templates provide figure and table caption examples using direct
formatting, even though an unused built-in `Caption` style also exists. Pass
`--preserve-caption-prototypes` to use the paragraph style, alignment, spacing,
and neutral text-run formatting of the first numbered figure and table caption
examples independently. Sample text is never copied. Missing examples leave
the corresponding discovered role unchanged. An explicit semantic style
profile remains authoritative and an explicit `--caption-font-size`, `--font`,
or `--east-asia-font` overrides the sampled metrics. The Python equivalent is
`assemble_markdown_template(..., preserve_caption_prototypes=True)`; the choice
is recorded in the manifest. Omitting the option preserves the existing style
discovery behavior.

For a document without a template, `--standalone-style manuscript` selects a
neutral left-aligned 11 pt body and maps Markdown heading levels to their
corresponding Word Heading levels. A supplied standalone title uses the Word
Title style. The pre-existing standalone defaults remain available as
`--standalone-style legacy`, which is the default for backward compatibility.
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

Keywords become a separate front-matter paragraph, including when no blank line precedes them. A `Table:` caption uses `table_caption`; the first row uses `table_head`, and other rows use `table_body`. `Table 5. Caption text` immediately followed by a table is also a caption, with its authored number and spacing preserved. Automatic `Table:` numbering avoids labels already supplied anywhere in the document; repeated authored labels are rejected. A numbered table mention that is not followed by a table remains ordinary prose.

Pipe-table separator cells specify alignment: `:---` is left, `:---:` is centered, and `---:` is right. An unmarked `---` inherits the applicable template paragraph style. In a fresh document it defaults to left. Configure the `TableCaption` and `TableFoot` styles in the template to control first-line indentation and paragraph spacing.

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

Only repeated labels under grouping headers (`level`, `group`, `category`, `class`, `type`, `family`, `site`, `site level`, `factor`, `block`, or a header ending in `level`) can merge vertically. Repeated measurements in other columns stay visible. Markdown and TeX share this policy.

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

The Python options are `inline_math="omml"`, `end_matter_placement="end"`, `table_notes="explicit"`, `table_note_diagnostics="error"`, and `native_toc=True` on `assemble_markdown_template`. Delivery profiles accept the same options. Output manifests record the selected settings. The renderer uses UTF-8 text I/O and closes temporary file/package handles before replacement for Windows compatibility.
`md2docx --template` writes the JSON assembly manifest beside the DOCX by default. Pass `--manifest-dir DIR` to store the manifest in a separate directory. The checksum file stays beside the DOCX.
