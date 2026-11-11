# Slides

`docforge slides` writes an editable PowerPoint file. The input is a JSON
list. Each object has a `title` string and a `bullets` list of strings.
The title is required and must not be empty. An existing output is left
unchanged unless `--force` is set.

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

The file is an Office Open XML package. Each title and each bullet is
stored in a shape with an explicit position, size, and black text. A
reader can show that text without a theme. This command does not style a
consumer template, and it does not restyle an existing deck. Characters
that are special in XML are escaped. Slide text is not used as a path
inside the package.
