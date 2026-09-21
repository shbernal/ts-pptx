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
writing. The file opens cleanly in desktop PowerPoint, and Keynote, LibreOffice Impress
and Google Slides import it on a best-effort basis.

Reach for it when a deck has to be built from data that changes: a monthly report,
one deck per customer, a hundred decks per night, or a download button on a page that
hands the user a deck built from what they are looking at.

## Install

```bash
pnpm add pptx-ts
```

`@shbernal/ts-pptx` is the same package under its first name, so install one or the other.
[Installation](docs/getting-started/installation.md) covers npm, CommonJS and the optional
math dependencies.

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
[Your first deck](docs/getting-started/first-deck.md) builds a bigger one from data, with a
table, a chart and speaker notes.

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
- **[Text that has to fit](docs/text-fit.md).** ts-pptx can measure the text
  against the real font and shrink or grow the box before it writes the file, instead
  of leaving you to guess at font sizes.

## Reading decks, not only writing them

Writing is half of it. ts-pptx also opens a `.pptx` you already have, which is
unusual: the library it descends from generates decks and does not read them.

- **[Look inside one](docs/reference/pptx-inspection.md)** and get slide count, size,
  parts, media and fonts, without loading the whole thing into a model.
- **[Edit one](docs/reading/read-and-edit.md)**. Open a deck, change the text on slide
  four, save it back. Parts you did not touch come out byte for byte as they went in.
- **[Turn one into code](docs/reference/pptx-to-script.md)**. Point it at a deck and
  get TypeScript that rebuilds it. Anything it could not express is reported to you
  rather than dropped in silence. It is the fastest way to learn the API: build a
  slide by hand in PowerPoint, then read the script for it.

## Where it runs

| Runtime | Load it with | `writeFile` |
| --- | --- | --- |
| Node.js 24 or later | `import`, or `require()` with the class on `.default` | writes to disk |
| A browser app built with a bundler | `import` | downloads the file |
| A browser page with no build step | `import TsPptx from "https://esm.sh/pptx-ts/browser"` in a module script | downloads the file |
| Deno, Bun, edge workers | `import` | throws; use `toBytes()` |

[Where it runs](docs/getting-started/runtime.md) lists every entry point and how each runtime
loads the one ESM build.

<!-- comparison:start -->
<!-- GENERATED REGION. Do not edit by hand.
     Regenerate with `pnpm run comparison:render`.
     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->

## How this compares with PptxGenJS

ts-pptx is an independent derivative of
[PptxGenJS](https://github.com/gitbrent/PptxGenJS), detached at its v4.0.1. Both were
measured on 2026-09-21 by building the same 22 deck intents with each library and reading
the bytes that came out.

- **Construct coverage:** ts-pptx emitted 21 of 22, pptxgenjs 10 of 22. Nothing in the
  corpus is emitted by pptxgenjs and not by ts-pptx.
- **Schema validity:** of the decks each library built, 21 of 21 ts-pptx decks and 0 of 10
  pptxgenjs decks validate with no error against the Open XML SDK.
- **Adoption:** pptxgenjs is downloaded 10,539,687 times a month, against 1,780 for
  ts-pptx. If a large installed base matters to you more than the differences above, use
  pptxgenjs.
- **Activity:** last commit on the default branch, 2026-09-15 for ts-pptx and 2025-06-26
  for pptxgenjs. Last npm publish, 2026-08-29 and 2025-06-26.

Where the two libraries part company is on the [comparison page](docs/comparison.md), and
[how it was measured](docs/comparison-method.md) has every full table. Every intent as
each library expresses it, including the calls that differ, is on [porting from
PptxGenJS](docs/comparison-syntax.md).

<!-- comparison:end -->

## Documentation

The documentation site, with the generated API reference, is at
**<https://shbernal.github.io/ts-pptx/>**. The [demos page](https://shbernal.github.io/ts-pptx/demos)
builds a quarterly review deck in your browser and previews the slides.

- Start with the [Introduction](docs/getting-started/introduction.md),
  [Your first deck](docs/getting-started/first-deck.md) and
  [Core concepts](docs/getting-started/concepts.md)
- [Tables](docs/tables.md), [connectors](docs/connectors.md), [groups](docs/groups.md) and
  [HTML tables to slides](docs/html-tables.md)
- [Smaller bundles](docs/bundle-size.md) for a browser program that composes only what it uses
- [Errors and warnings](docs/errors-and-warnings.md) and
  [troubleshooting](docs/troubleshooting.md)

## Something wrong, or missing?

Open an issue: <https://github.com/shbernal/ts-pptx/issues>. Errors the library knows
are its own fault print that link themselves.

If an agent writes most of your code, install the `ts-pptx-upstream` skill that ships inside
the package. It turns a library defect the agent hits into a filed issue with a small
reproduction, instead of a silent workaround:

```bash
npx skills add ./node_modules/pptx-ts -s '*' -a claude-code -a codex -a universal -y
```

Name the runtimes you use. [CONTRIBUTING.md](CONTRIBUTING.md) covers what the skill does
with a report and how to refresh it after a version bump.

## License

Copyright (c) 2015-2022 Brent Ely.
Modifications copyright (c) 2026 shbernal.

[MIT](LICENSE)
