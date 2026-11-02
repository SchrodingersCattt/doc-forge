# Conservative DOCX and TeX review bundles

`tex-review-export` compares the accepted visible text of a reviewed DOCX with
the DOCX on which it was based. It applies supported review edits to the
supplied TeX source and writes a private, hash-checked bundle:

```shell
docforge tex-review-export baseline.docx reviewed.docx source.tex --workdir review-bundle
# Edit review-bundle/document.tex.
docforge tex-review-build review-bundle -o edited.docx
```

`document.tex` is the editable source. `original.tex`, `baseline.tex`, the two
DOCX snapshots, and `manifest.json` record the inputs, source offsets, and
SHA-256 values; keep the bundle private because it contains the full document.
The command does not modify the original inputs. Multiple changed paragraphs
require stable matching Word paragraph IDs. Comments are retained in the
reviewed DOCX snapshot rather than converted into TeX.

If `document.tex` is unchanged, rebuilding copies the reviewed DOCX exactly
and verifies its SHA-256. For a changed file, the edit must fall within one
uniquely mapped plain paragraph and one existing Word text run. Only that
run's text is changed. Other package parts, fields, images, bookmarks, and
comment anchors remain in the source package. Snapshots and mapped text are
verified before publication.

The supported subset is intentionally narrow: a body paragraph must match a
unique, complete, standalone TeX prose line. It cannot contain TeX commands,
special characters, inline math, or other markup; code and other environments
are not source anchors. Inserted, deleted, reordered, or mixed-format
paragraphs, edits to fields, drawings, tables, footnotes, and bibliography
structure, and edits crossing text-run or anchor boundaries are refused.
Multiple edits can succeed only when their combined changed range stays within
one Word text run. Comment bodies are retained as review data, not interpreted
or applied to TeX. This workflow does not expand macros or guarantee that an
arbitrary manuscript can be mapped; a refusal requires a separate, carefully
reviewed edit. It is not a replacement for full TeX compilation and scientific
review.
