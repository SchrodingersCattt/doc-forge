# Template roles and layout

`docforge md2docx source.md -o output.docx --template template.docx --metadata metadata.md` discovers semantic paragraph styles from the supplied template. Font family, size, and color inherit from those styles; reference paragraphs also inherit hanging indents and spacing. Sample run typography does not become direct formatting on generated text. Explicit typography flags remain available.

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

| Group | Sample | Value |
| --- | --- | --- |
| A | one | 2.0 |
| A | two | 2.0 |

[a] A note directly after the table.

Table note: Another explicit table note.
```

Keywords become a separate front-matter paragraph, including when no blank line precedes them. A `Table:` caption uses `table_caption`; the first row uses `table_head`, and other rows use `table_body`. Use `Table note:` or `Table foot:` for explicit notes, or `[a]`-style lines immediately following a table or another table note.

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

`--end-matter-placement end` places level-one `Table of Contents`, `TOC`, and `Graphical Abstract` source sections after references, starting on a new page. The default `source` preserves source position. The option also places a requested native TOC at the end.

`--native-toc` inserts a native Word TOC field. Its cached result contains the emitted heading text immediately, without invented page numbers; `w:updateFields` requests a refresh on open. Word computes page numbers when it updates the field. Automatic update behavior in other viewers varies. Without a native TOC request, the generator disables automatic field updates as before.

The Python options are `inline_math="omml"`, `end_matter_placement="end"`, and `native_toc=True` on `assemble_markdown_template`. Delivery profiles accept the same options. Output manifests record the selected settings. The renderer uses UTF-8 text I/O and closes temporary file/package handles before replacement for Windows compatibility.
