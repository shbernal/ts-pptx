---
doc-schema-version: 1
title: "PptxGenJS"
summary: "Start here for the purpose, setup, and main workflows in PptxGenJS."
read_when:
  - Getting oriented in this project
  - Updating the main project overview
doc_type: "overview"
---

# PptxGenJS

PptxGenJS generates PowerPoint `.pptx` packages from TypeScript and modern
JavaScript. This fork targets modern ESM applications and keeps the public
package boundary explicit.

## Start Here

- [Project target](project-target.md): what this fork is optimized for.
- [Runtime and package support](runtime-and-package-support.md): supported
  package imports and dropped upstream surfaces.
- [Development guide](development.md): setup, source layout, and generated
  output rules.
- [Testing guide](testing.md): regression, schema, package, demo, and manual
  verification.
- [Reference](reference/index.md): public API reference and stable command
  surfaces.

## Maintenance Focus

- Keep PptxGenJS focused on reusable PPTX generation behavior.
- Treat `package.json` exports and generated declarations as the package API
  boundary.
- Keep OOXML behavior grounded in fixtures, schema validation, and small local
  notes.
- Keep downstream-specific deck production behavior in `downstream`, not this
  package.

## Standard Verification

```bash
pnpm run build
pnpm run typecheck
pnpm run test:package
```
