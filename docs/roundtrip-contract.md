# Cross-format round-trip contract

Adapters share one portable contract for reversible round trips. A fixture is
created under an isolated temporary directory (`tmp_path` in tests); committed
fixtures use opaque IDs and synthetic content. Real local decks and workbooks
are opt-in through an environment-controlled fixture directory and never enter
reports or committed tests.

The runner has three modes:

- `EXACT_REUSE` copies an unchanged source to a distinct destination. It first
  verifies the opaque source snapshot SHA-256, rejects resolved path collisions,
  leaves the source untouched, and checks byte-for-byte SHA equality after the
  copy.
- `SOURCE_PATCH` applies a source-package edit. It compares semantic content
  and structure and does not require package bytes to match.
- `REBUILD` exports and rebuilds a changed model. It compares only the adapter's
  semantic projection.

Every adapter exposes `probe`, `export`, `build`, `project`, and
`unsupported_features`. The shared projection uses these invariant codes:

`RT-OPEN`, `RT-EXACT`, `RT-CONTENT`, `RT-STRUCTURE`, `RT-STYLE`,
`RT-GEOMETRY`, `RT-REL`, `RT-MEDIA`, `RT-IDEMP`, and `RT-UNSUPPORTED`.

A public receipt contains only a synthetic artifact ID, format, mode, invariant
statuses, aggregate counts, stable diagnostic categories, and block ordinals.
It does not contain source text, XML, user filenames, absolute paths, project
identifiers, or raw hashes. Volatile package metadata is excluded from semantic
projections. Missing or corrupted snapshots fail with
`SNAPSHOT_SHA_MISMATCH`; source/output/report collisions fail with
`PATH_COLLISION` before publication.

For OOXML sources, format dispatch checks both extension and package magic
before staging or invoking Pandoc. Word-processing OOXML is supported;
presentations, spreadsheets, OLE2 workbooks, and extension/content mismatches
return typed `UNSUPPORTED_FORMAT` or `INVALID_FORMAT` results without creating
work directories or outputs.
