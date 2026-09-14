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
      text: Get started
      link: /getting-started/introduction
    - theme: alt
      text: See a deck built in your browser
      link: /demos
    - theme: alt
      text: API reference
      link: /reference/api/
features:
  - title: Write
    details: "Slides, masters and themes, charts with real embedded workbooks, tables, groups, connectors, gradients, images in shapes, OLE objects, 3D models and LaTeX maths. Emitted as OOXML, validated against the ECMA-376 schemas."
    link: /getting-started/first-deck
    linkText: Your first deck
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

[Installation](getting-started/installation.md) covers npm, CommonJS, the optional math
dependencies and the package's other name.

## Where to go next

- [Introduction](getting-started/introduction.md): what the library does, what you can build with
  it, and its limits.
- [Your first deck](getting-started/first-deck.md): a small deck built from data, start to
  finish.
- [Demos](demos.md): a deck built in your browser and previewed in the page.
- [Comparison with PptxGenJS](comparison.md): what each library emits, measured by building the
  same decks with both.
