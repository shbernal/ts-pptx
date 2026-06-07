# Documentation

This directory contains the maintained project documentation for this PptxGenJS
fork. Prefer docs here over legacy upstream-era notes in demos or generated
artifacts.

## Start Here

- [Project target](project-target.md): what this project is optimized for.
- [Runtime and package support](runtime-and-package-support.md): supported
  imports, dropped upstream support, and shipped artifacts.
- [Development guide](development.md): setup, source layout, generated outputs,
  and contribution rules.
- [Testing guide](testing.md): regression, schema, package, demo, and manual
  verification.
- [Release workflow](RELEASING.md): scoped-package release preparation,
  automated npm publishing, and package-surface checks.
- [Agent development guide](agent-development.md): expectations for Codex and
  other agent-assisted changes.
- [OOXML agent context](ooxml-agent-context.md): project-specific OOXML
  reference and validation workflow.
- [Upstream signal workflow](upstream-signal-workflow.md): how to classify
  upstream issues and PRs without reintroducing dropped package targets.
- [Legacy autoloop workflow](legacy-autoloop.md): retained legacy automation
  notes.

## Documentation Rules

- Keep docs aligned with the current package target.
- Do not document CJS or IIFE as supported workflows.
- Keep release runtime and declaration artifacts under `dist/` treated as
  generated outputs unless a task explicitly asks to refresh them.
- For OOXML behavior, prefer small repo-specific notes with section references
  over copied standards text.
