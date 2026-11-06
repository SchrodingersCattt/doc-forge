## Summary

<!-- What changed. -->

## How to verify

<!-- Commands, tests, or checks a reviewer can run. -->

## Documentation

In-scope changes update stable docs in this pull request. Stable docs are `docs/pipeline.md`, `docs/roundtrip-contract.md`, `docs/docxmd-syntax.md`, `docs/migration.md`, and ADRs under `docs/adr/`. Update `AGENTS.md` when the change adds a standing package-boundary rule, and update `docs/migration.md` or an ADR under `docs/adr/` when that migration story or design changes. The review bar is in `AGENTS.md`. Prefer fixtures, goldens, and contract tests as the behavioral truth.

- Docs paths touched, or `Docs: N/A` and a one-line reason:
- Public API, CLI, Markdown/TeX/DOCX contracts, migration, or ADR-level design changed: yes / no
- `AGENTS.md`, `docs/migration.md`, and/or an ADR under `docs/adr/` updated where relevant: yes / no / not relevant
- [ ] Project-agnostic: no consumer project names, journal-template identifiers, sample material names, or consumer repository paths
