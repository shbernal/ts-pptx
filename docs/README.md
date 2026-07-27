---
doc-schema-version: 1
title: "Documentation"
summary: "Index of maintained ts-pptx documentation and documentation rules."
read_when:
  - Looking for the maintained docs surface
  - Updating documentation structure or rules
doc_type: "overview"
---

# Documentation

This directory contains the maintained project documentation for ts-pptx.
Prefer docs here over legacy upstream-era notes in demos or generated
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
- [Backlog workflow](backlog-workflow.md): how to classify
  upstream issues and PRs without reintroducing dropped package targets.

## Feature Guides

- [Grouping objects](groups.md): `addGroup()` / `groupObjects()`, the identity
  child space, framing, nesting, and cross-references into a group.
- [Connectors](connectors.md): `addConnector()` straight/elbow/curved lines,
  bend control, shape binding, and the `addShape()` vs. `addConnector()` split.
- [Image embedded in a shape](image-in-shape.md): clip a picture to a preset or
  freeform shape and crop it to fill the box.
- [OLE embedded objects](ole-objects.md): `addOleObject()` embeds a workbook,
  document, or any payload so it opens in place on double-click.
- [Animations and transitions](animations-and-transitions.md): slide/shape
  animation and transition emit.
- [Native backgrounds and gradients](native-backgrounds-and-gradients.md):
  native PPTX gradient fills and required review gates.
- [PPTX to script](reference/pptx-to-script.md): turn an existing deck into
  runnable TypeScript, the two output tiers, and the fidelity notes that state
  what a conversion drops.
- [Math and LaTeX](math-latex.md): OMML math emit.
- [Embedded fonts](embedded-fonts.md): font embedding, merge, and fixtures.
- [Measured text fit](measured-text-fit.md): the export-time shrink/resize pass.

## Documentation Rules

- Keep docs aligned with the current package target.
- Do not document CJS or IIFE as supported workflows.
- Keep release runtime and declaration artifacts under `dist/` treated as
  generated outputs unless a task explicitly asks to refresh them.
- For OOXML behavior, prefer small repo-specific notes with section references
  over copied standards text.
