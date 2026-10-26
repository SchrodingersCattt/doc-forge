# ADR 0001: SemanticDocument hub

## Status

Accepted for the first implementation slice of issue 50. Later slices extend adapters and fidelity coverage; they do not replace this model.

## Decision

`SemanticDocument` (`docforge.semantic.v1`) is a concrete, frozen tree. Format readers and writers are adapters. A DOCX package is not stored inside the tree.

`SourceContext` holds the source SHA-256, file name, opaque part names, and retained native features. Exact byte identity applies only to an unchanged DOCX copy (`EXACT_REUSE`). Supported package edits, including restyle, use `SOURCE_PATCH`. Markdown and LaTeX conversions rebuild.

`StyleProfile` (`docforge.style.v1`) is independent of the readers. A Word template remains one way to name styles. Restyle updates style definitions and, under `clear_conflicting` or `clear_all`, direct run properties. It does not change `w:t` text.

## Node taxonomy

Blocks: paragraph, heading, list, list item, table, table row, table cell, figure, caption, math, code, quote, thematic break, bibliography, reference entry, opaque.

Inlines: text, emph, strong, code, link, math, break, subscript, superscript, underline, strike, span, citation, opaque.

Unsupported source constructs become diagnostics. A diagnostic is a loss when its severity is `error` or its category is `loss`. `loss_policy=error` and `fidelity=preserve` refuse to publish a loss before replacing an output. Informational warnings do not block publication. Semantic conversion defaults to `warn`.

DOCX projection walks `w:body` in order, including nested tables and drawings. `restyle` publishes only after the staged package passes `audit`. `StyleProfile.template` is not applied. A new style is attached only through `RoleSpec.select`; otherwise the role is unused.

## Non-goals

Full TeX macro expansion, pixel-identical pagination, treating plain Markdown as a lossless DOCX, and moving consumer templates or manuscript text into this package.

## Compatibility

`md2docx`, `tex2docx`, and the other existing commands stay. `convert` to DOCX calls those builders. `set_sizes` still changes the active TeX point size, but it no longer rewrites function default tuples. Output names may contain `final`; a delivery profile opts into `forbid_name_tokens` when it wants that ban.
