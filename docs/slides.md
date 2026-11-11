# Slides

`docforge slides` writes an editable PowerPoint file. The input is a JSON
list. Each object has a `title` string and a `bullets` list of strings.
The title is required and must not be empty. Characters that XML 1.0
cannot store are refused before a file is opened.

An existing file is left unchanged unless `--force` is set. A directory
at the output path is left unchanged. The parent directory must already
exist; it is not created. When the output cannot be written, the error
names that requested path and the command exits 2.

```bash
docforge slides slides.json -o deck.pptx
docforge slides slides.json -o deck.pptx --force
```

```json
[
  {"title": "Overview", "bullets": ["First point", "Second point"]},
  {"title": "Next", "bullets": []}
]
```

The same records can be written from Python with
`docforge.slides.write_pptx(path, slides)`.

The file is an Office Open XML package. Each slide has one title text
box. Bullets are paragraphs in a second text box. A slide with no bullets
has only the title box. Both boxes have an explicit position, size, and
black text, so a reader can show that text without a theme. This command
does not style a consumer template, and it does not restyle an existing
deck. Characters that are special in XML are escaped. Slide text is not
used as a path inside the package.
