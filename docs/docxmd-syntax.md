# Reversible DOCX ↔ Markdown bundles (`docforge.docxmd`)

`docforge.docxmd` turns a DOCX into section Markdown files plus a references
JSON, and turns them back into a DOCX that Word lays out identically. The
Markdown is the only editable source; a tracked-change DOCX is then made by
redlining the rebuilt file against the original.

```text
docforge md-export  paper.docx md/ --id main --split split.json --bibliography 07.references.json
docforge md-build   md/ --id main -o paper.clean.docx
docforge redline    --baseline paper.docx --current paper.clean.docx -o paper.tracked.docx --revision-author NAME
docforge md-roundtrip paper.docx md/ --id main --split split.json --render word
```

`md-roundtrip` exports, rebuilds and compares. It passes only if:

- the document models are equal (block order, paragraph properties, every
  character with its resolved visible formatting, revisions, comments,
  pictures, tables, page setup and shared parts);
- a redline of the source against the rebuilt file reports zero changed,
  inserted or deleted blocks;
- with `--render`, every page of both PDFs has the same glyphs (character,
  font, size, colour) within 0.5 pt and the same number of images.

Word is the reference renderer (`--render word`, Windows with pywin32).
LibreOffice shapes text per run, so merging two runs with identical formatting
can move its line breaks even though Word lays the text out the same way.

## Atomic one-pass contract

For template-driven Markdown assembly, `docforge contract` is the project-neutral
one-pass entry point:

```text
docforge contract 01.main.md -o main.docx --template template.docx \
  --metadata 00.metadata.md --bibliography references.json --force
```

The command discovers every semantic paragraph role before creating output,
builds and verifies in a sibling temporary directory, then publishes the DOCX,
`.manifest.json`, and `.sha256` sidecars together. A missing role, malformed
input, or failed verification leaves an existing output untouched and does not
publish a misleading partial file. The manifest carries
`contract.schema=docforge.docx-contract.v1`, the resolved role map, and one
decision record for each parsed Markdown block (source, kind, role, and text
hash).

## Bundle layout

```text
md/
  bundle.json              documents, their parts, templates and metadata
  00.metadata.md …         Markdown parts in document order (one file per part)
  07.references.json       shared bibliography (docforge.bibliography.v2)
  media/<doc-id>/…         pictures referenced from the Markdown
  _docx/<doc-id>.docx      the source package (styles, numbering, headers, theme)
```

Several documents (for example `main` and `si`) can share one bundle and one
references JSON. `_docx/<id>.docx` supplies everything that is not body text:
styles, numbering, headers and footers, footnote and endnote bodies, settings
and fonts. Body content comes only from the Markdown.

### Split configuration

`--split` takes a JSON list. A part starts at the first block whose text
matches `start` (a regular expression); the first entry has `start: null`.

```json
[
  {"file": "00.metadata.md", "start": null},
  {"file": "01.abstract.md", "start": "^ABSTRACT"},
  {"file": "02.introduction.md", "start": "^Molecular substitution"}
]
```

## Document header

The first part starts with a comment that describes the page setup and the
paragraph classes:

```text
<!-- docforge:document id=main
section: pgSz@w=11906 pgSz@h=16838 pgMar@top=1440 pgMar@left=1800 …
heading1: Heading1
class TAMainText | p: pStyle=TAMainText spacing@after=240 ind@firstLine=480 | r: rFonts@cs=Times
class VAFigureCaption | p: pStyle=VAFigureCaption | r:
-->
```

A class is a named paragraph format: paragraph properties (`p:`) and the base
run properties of its text (`r:`). Every part then starts with
`<!-- docforge:part default=CLASS -->`; paragraphs without an explicit class
use that default.

## Property tokens

Properties are written as flat tokens taken from the WordprocessingML element
tree:

| Token | Meaning |
| --- | --- |
| `jc=center` | `<w:jc w:val="center"/>` |
| `spacing@after=240` | attribute `w:after` of `<w:spacing>` |
| `keepNext` | presence of an empty element |
| `rPr.lang@eastAsia=zh-CN` | nested element (`pPr/rPr/lang`) |
| `tab#2@pos=720` | second repeated sibling |
| `-color` | remove an inherited element (removals apply first) |

Values containing spaces or `"{}|` are double-quoted. Children are rebuilt in
schema order, so the token order does not matter.

## Paragraphs and headings

A paragraph is one line of inline Markdown, optionally followed by an
attribute line:

```text
Single-crystal X-ray diffraction confirms that …
{: .TAMainText ind@firstLine=0 r:color=000000}
```

- `.Class` selects the class; other tokens override its paragraph properties.
- `r:` tokens change the base run properties of the whole paragraph.
- `tc:` and `tr:` tokens (first paragraph of a table cell) set cell and row
  properties.
- `role=title|authors|affiliation|keywords|correspondence` marks front-matter
  paragraphs and fills `bundle.json` metadata.
- An attribute line alone is an empty paragraph; `{:}` uses the default class.

Headings use `#` × (outline level + 1), for example `## Results`. The heading
class comes from `headingN:` in the header unless the attribute line names
one.

## Inline syntax

| Markdown | DOCX |
| --- | --- |
| `**bold**`, `*italic*` | `w:b`, `w:i` |
| `<u>…</u>`, `~~…~~`, `==…==` | underline, strike, yellow highlight |
| `<sup>…</sup>`, `<sub>…</sub>` | `vertAlign` |
| `[text]{color=C00000 sz=20}` | any other run properties as tokens |
| `[text](https://…)` | hyperlink |
| `{++text++}{author="A" date=…}` | tracked insertion |
| `{--text--}{author="A" date=…}` | tracked deletion |
| `\citep{key1,key2}` | citation rendered as the reference numbers |
| `![alt](media/main/image1.png){cx=… cy=… name=… crop=l,t,r,b}` | inline picture |
| `<tab/>`, `<br/>`, `<cr/>`, `<nbhyphen/>`, `<shy/>` | tab, line break, carriage return, non-breaking hyphen, soft hyphen |
| `<sym font="Symbol" char="F062"/>` | symbol character |
| `<fld-begin/>`, `<instr code="PAGE"/>`, `<fld-sep/>`, `<fld-end/>` | complex field |
| `<comment-start id=c1/>`, `<comment-end id=c1/>` | comment range (the end implies the reference mark) |
| `<bookmark-start name="_Ref1"/>`, `<bookmark-end name="_Ref1"/>` | bookmark (Word's automatic `OLE_LINK`/`_Hlk`/`_GoBack` bookmarks are dropped unless a field or link refers to them) |
| `<ooxml level="run">…</ooxml>` | raw OOXML for anything without a Markdown form |

Spans only carry properties that change what is visible. For example, an
`eastAsia` font hint on ASCII text is dropped because Word draws ASCII with
the `ascii` font anyway. Spans nest outermost first:
revision, link, span, highlight, underline, strike, bold, italic,
superscript/subscript.

The characters `\ * < { } |` are always escaped with a backslash, and doubled
`==` and `~~` are escaped too. Leading or trailing spaces of a paragraph are
written as `&#32;`. A paragraph that would start like a heading or an
attribute line is escaped (`\#`, `\:`).

## Blocks

```text
::: table {: tblStyle=TableGrid grid=4620,4620 default=Normal}
| Property {: tc:tcW@w=4620} | Value |
|---|---|
| Density | 1.9 <p/> second paragraph in the cell |
:::

::: comment {: id=c1 author="A. B." initials=AB date=2026-01-01T00:00:00Z done=0}
Comment text
{: pStyle=CommentText}
:::

::: references {: .TFReferencesSection style=main label="\\[{n}\\]<tab/>"}
:::

::: ooxml
<w:p>…</w:p>
:::
```

- Table rows use GitHub table syntax; ` <p/> ` separates paragraphs in a cell.
- Comment bodies are written at the end of the part that anchors them.
- `::: references` renders the bibliography of `style` in JSON order. `label`
  is the prefix of each entry (`{n}` is the number); without a label the
  class numbering supplies it.

## References JSON

```json
{
  "_schema": "docforge.bibliography.v2",
  "smith2020": {
    "rendered": {"main": "Smith, J. A study. *J. Chem.* **2020**, 1, 1."},
    "layout": {"main": "{: ind@firstLine=0}"}
  }
}
```

Keys are generated from the first author's surname and the year. `rendered`
is inline Markdown per style; `layout` optionally overrides one entry's
paragraph format. Within a style, JSON order is the reference order, and
`\citep` numbers are positions in that order. A run of three or more
consecutive numbers is written as a range (`3–5`).

## Editing

- Edit text in place; keep the attribute line under its paragraph.
- New paragraphs need no attribute line when the part default fits.
- Add a reference by inserting an entry in JSON order and citing its key.
- Run `md-build`, then `redline` against the original DOCX. Because the
  unedited bundle rebuilds to a file Word lays out the same way as the source,
  the redline contains only the edits.
