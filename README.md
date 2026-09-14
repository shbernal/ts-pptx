# ts-pptx

[![npm](https://img.shields.io/npm/v/pptx-ts)](https://www.npmjs.com/package/pptx-ts)
[![weekly downloads](https://img.shields.io/npm/dw/pptx-ts.svg?label=npm%20downloads&logo=npm)](https://www.npmjs.com/package/pptx-ts)
[![total downloads](https://img.shields.io/npm/dt/pptx-ts.svg?label=npm%20total%20downloads&logo=npm)](https://www.npmjs.com/package/pptx-ts)
[![CI](https://github.com/shbernal/ts-pptx/actions/workflows/ci.yml/badge.svg)](https://github.com/shbernal/ts-pptx/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**Write a program, get a PowerPoint file.**

A `.pptx` is a zip full of XML. ts-pptx writes that zip for you, so you describe
slides in TypeScript and a `.pptx` comes out the other end. PowerPoint never runs, no
Office licence is involved, and nothing has to be installed on the machine doing the
writing. The file it produces is the real format: it opens in PowerPoint on Windows
and on a Mac, and imports into Google Slides, Keynote and LibreOffice Impress.

Reach for it when a deck has to be built from data that changes: a monthly report,
one deck per customer, a hundred decks per night, or a download button on a page that
hands the user a deck built from what they are looking at.

## Install

```bash
pnpm add pptx-ts
```

The scoped name [`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx)
is the same package, same version, published from the same commit. It is the name this
project shipped under first, so installs that already use it keep working. Pick one of
the two names and stay on it: two copies of the library in one dependency tree are two
separate libraries as far as your program is concerned. Everything here uses `pptx-ts`.

## Quick start

```ts
import TsPptx from "pptx-ts"

const pptx = new TsPptx()
const slide = pptx.addSlide()

slide.addText("Hello from ts-pptx", {
  x: 1,
  y: 1,
  w: 8,
  h: 1,
  fontSize: 24,
  color: "363636",
})

await pptx.writeFile({ fileName: "example.pptx" })
```

That is the whole shape of it. Make a presentation, add a slide, put things on the
slide, write the file. Positions are in inches by default, so `x: 1, y: 1` is an inch
in from the top-left corner of a 10 by 5.625 inch slide.

## What you can put on a slide

Text and rich paragraphs. Tables, including ones that spill onto as many slides as
they need. Shapes and connectors between them. Pictures, SVGs, video and audio.
Charts, with a real embedded workbook behind them, so double-clicking a chart in
PowerPoint opens its data the way it does for a chart a human made. Speaker notes,
sections, slide masters and layouts, gradients, an image clipped to a shape, a
spreadsheet embedded as an object, a 3D model, and LaTeX maths.

Two features worth knowing about by name:

- **[An HTML table becomes slides](docs/html-tables.md).** Point `tableToSlides` at a
  `<table>` you already have and it comes out as a PowerPoint table, paged across
  slides. Works in a browser and under Node.
- **[Text that has to fit](docs/measured-text-fit.md).** ts-pptx can measure the text
  against the real font and shrink or grow the box before it writes the file, instead
  of leaving you to guess at font sizes.

## Reading decks, not only writing them

Writing is half of it. ts-pptx also opens a `.pptx` you already have, which is
unusual: the library it descends from generates decks and does not read them.

- **[Look inside one](docs/reference/pptx-inspection.md)** and get slide count, size,
  parts, media and fonts, without loading the whole thing into a model.
- **[Edit one](docs/reference/pptx-read.md)**. Open a deck, change the text on slide
  four, save it back. Parts you did not touch come out byte for byte as they went in.
- **[Turn one into code](docs/reference/pptx-to-script.md)**. Point it at a deck and
  get TypeScript that rebuilds it. Anything it could not express is reported to you
  rather than dropped in silence. It is the fastest way to learn the API: build a
  slide by hand in PowerPoint, then read the script for it.

## Where it runs

- **Node 24 and up.** `import` it, or `require()` it. Node loads ES modules through
  `require()` since 22.12, so `const { default: TsPptx } = require("pptx-ts")` works
  on every version of Node this package supports.
- **Browsers.** Import it in any app built with Vite, Webpack, Rollup, or any bundler
  at all. With no build step, an ESM CDN serves it straight to a module script:

  ```html
  <script type="module">
    import TsPptx from "https://esm.sh/pptx-ts/browser"
    const pptx = new TsPptx()
    pptx.addSlide().addText("Built in your browser", { x: 1, y: 1, w: 8, h: 1 })
    await pptx.writeFile({ fileName: "example.pptx" }) // downloads the file
  </script>
  ```

  The browser build is checked in CI against a real Chromium, and the deck a browser
  assembles is compared part for part against the one Node builds. They are identical.
- **Deno, Bun, edge workers.** They author and hand back bytes like anywhere else.
  `writeFile()` is the one thing they cannot do, because there is no disk to write to
  and no page to download onto.

Full detail is in [runtime and package support](docs/getting-started/runtime.md).

<!-- comparison:start -->
<!-- GENERATED REGION. Do not edit by hand.
     Regenerate with `pnpm run comparison:render`.
     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->

## How this compares with PptxGenJS

ts-pptx is an independent derivative of
[PptxGenJS](https://github.com/gitbrent/PptxGenJS), detached at its v4.0.1. Both were
measured on 2026-09-06 by building the same 22 deck intents with each library and reading
the bytes that came out.

- **Construct coverage:** ts-pptx emitted 21 of 22, pptxgenjs 10 of 22. Nothing in the
  corpus is emitted by pptxgenjs and not by ts-pptx.
- **Schema validity:** of the decks each library built, 21 of 21 ts-pptx decks and 0 of 10
  pptxgenjs decks validate with no error against the Open XML SDK.
- **Adoption:** pptxgenjs is downloaded 11,116,327 times a month, against 2,019 for
  ts-pptx. If a large installed base matters to you more than the differences above, use
  pptxgenjs.
- **Activity:** last commit on the default branch, 2026-09-05 for ts-pptx and 2025-06-26
  for pptxgenjs. Last npm publish, 2026-08-29 and 2025-06-26.

The full tables, the method behind them, and where the two libraries part company are on
the [comparison page](docs/comparison.md). Every intent as each library expresses it,
including the calls that differ, is on [side-by-side syntax](docs/comparison-syntax.md).

<!-- comparison:end -->

## Documentation

The full documentation site, including the generated API reference, is at
**<https://shbernal.github.io/ts-pptx/>**.

The [demos page](https://shbernal.github.io/ts-pptx/demos) builds a quarterly review
deck in your browser and previews the slides. Nothing to clone, nothing to install.

- [Tables](docs/tables.md), [groups](docs/groups.md),
  [connectors](docs/connectors.md), [HTML tables to slides](docs/html-tables.md)
- [Errors](docs/errors.md) and [diagnostics](docs/diagnostics.md): what the library
  throws, what it warns about, and how to route or silence the warnings
- [Troubleshooting](docs/troubleshooting.md)

## Something wrong, or missing?

Open an issue: <https://github.com/shbernal/ts-pptx/issues>. Errors the library knows
are its own fault print that link themselves.

If an agent writes most of your code, install the `ts-pptx-upstream` skill that ships
inside the package. An agent that hits a library defect usually routes around it in
silence, and nobody ever hears about it. The skill turns that moment into a filed issue
with a reproduction small enough to become a regression test here:

```bash
npx skills add ./node_modules/pptx-ts -s '*' -a claude-code -a codex -a universal -y
```

Name the runtimes you actually use, as above. [CONTRIBUTING.md](CONTRIBUTING.md) covers
what the skill does with the report, keeping your decks off a public tracker, and
refreshing the installed copy after a version bump.

## License

Copyright (c) 2015-2022 Brent Ely.
Modifications copyright (c) 2026 shbernal.

[MIT](LICENSE)
