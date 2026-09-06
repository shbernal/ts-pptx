---
layout: home
doc-schema-version: 1
title: "ts-pptx"
summary: "Start here for the purpose, setup, and main workflows in ts-pptx."
read_when:
  - Getting oriented in this project
  - Updating the main project overview
doc_type: "overview"
hero:
  name: "ts-pptx"
  text: "PowerPoint decks from TypeScript"
  tagline: "Write .pptx packages that open cleanly in PowerPoint, read them back, and turn one into the script that would rebuild it."
  actions:
    - theme: brand
      text: See a deck built in your browser
      link: /demos
    - theme: alt
      text: Start here
      link: /project-target
    - theme: alt
      text: API reference
      link: /reference/api/
features:
  - title: Write
    details: "Slides, masters and themes, charts with real embedded workbooks, tables, groups, connectors, gradients, images in shapes, OLE objects, 3D models and LaTeX maths. Emitted as OOXML, validated against the ECMA-376 schemas."
    link: /reference/api/
    linkText: API reference
  - title: Read
    details: "Open an existing package and inspect it through typed accessors. What the reader cannot yet see is measured and published rather than guessed at."
    link: /reference/pptx-read
    linkText: Reading a deck
  - title: Convert to a script
    details: "Turn a .pptx into the TypeScript that would rebuild it, with a fidelity note for everything the conversion could not carry."
    link: /reference/pptx-to-script
    linkText: pptx → script
  - title: Runs where you do
    details: "One ESM build. Node 24+, where require() reaches it through Node's own interop; any bundler; a browser, straight from an ESM CDN if you have no build step. No office application in the loop anywhere."
    link: /runtime-and-package-support
    linkText: Runtime support
---

## Install

```bash
pnpm add pptx-ts
```

`@shbernal/ts-pptx` is an alias for the same package, published from the same commit
at the same version. It is the name this project published under first, and installs
that already use it keep resolving. Install one or the other rather than both. These
docs use `pptx-ts` throughout.

## Start here

- [Project target](project-target.md): what this project is optimized for.
- [Runtime and package support](runtime-and-package-support.md): supported package
  imports, and how each runtime loads the one build.
- [Development guide](development.md): setup, source layout, and generated
  output rules.
- [Testing guide](testing.md): regression, schema, package, browser, and manual
  verification.
- [Reference](reference/index.md): public API reference and stable command
  surfaces.

## Maintenance focus

- Keep ts-pptx focused on reusable PPTX generation behavior.
- Treat `package.json` exports and generated declarations as the package API
  boundary.
- Keep OOXML behavior grounded in fixtures, schema validation, and small local
  notes.
- Keep consumer-specific deck production behavior in the downstream consumer, not
  this package.

## Standard verification

```bash
pnpm run verify
```
