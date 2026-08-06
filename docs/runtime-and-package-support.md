---
doc-schema-version: 1
title: "Runtime And Package Support"
summary: "Supported imports, dropped upstream support, and shipped artifacts."
read_when:
  - Changing package exports or runtime support
  - Explaining ESM-only package behavior
  - Updating shipped artifact policy
doc_type: "reference"
---

# Runtime And Package Support

This project ships an ESM-only package for TypeScript and modern JavaScript
applications.

## Supported Surface

Use the package export:

```ts
import TsPptx from "@shbernal/ts-pptx"
```

The package publishes:

- `dist/index.js` and `dist/index.d.ts` as the default ESM package entry — it
  also exports the public enums, shared types, layout constants, and unit helpers.
  See [Which Build The Bare Import Gives You](#which-build-the-bare-import-gives-you)
  — under Node and in a browser bundle the same import gives you `dist/node.js` /
  `dist/browser.js` instead.
- `dist/inspect.js` and `dist/inspect.d.ts` for low-level PPTX package
  inspection, slide/object extraction, and geometry helpers.
- `dist/measure.js` and `dist/measure.d.ts` for headless text-measurement and
  autofit helpers (see [Measured Text Fit](measured-text-fit.md)).
- `dist/read.js` and `dist/read.d.ts` for opening, editing, and round-tripping
  an existing `.pptx` (see [PPTX Read / Round-Trip](reference/pptx-read.md)).
- `dist/script.js` and `dist/script.d.ts` for turning an existing `.pptx` into
  TypeScript source that rebuilds it through the write API (see
  [PPTX To Script](reference/pptx-to-script.md)).
- `dist/math.js` and `dist/math.d.ts` for LaTeX/MathML → OMML conversion (see
  [Math Equations](math-latex.md)).
- `dist/zip.js` and `dist/zip.d.ts` for the shared OPC/zip package plumbing
  used by `read` and `inspect`.
- `dist/html.js` and `dist/html.d.ts` for converting an existing HTML `<table>`
  into slides. One artifact serves both runtimes — there is deliberately no
  `browser`/`node` condition split, because the entry works with whatever DOM
  the caller has (see
  [HTML tables → slides](https://github.com/shbernal/ts-pptx/blob/master/README.md#html-tables--slides)).
- `dist/node.js` and `dist/node.d.ts` for explicit Node.js consumers.
- `dist/browser.js` and `dist/browser.d.ts` for explicit browser consumers.
- package `exports` entries for `.`, `./inspect`, `./measure`, `./read`,
  `./script`, `./math`, `./zip`, `./html`, `./node`, and `./browser`.

Supported environments are modern module-aware environments:

- Node.js `>=24`.
- Vite, Rolldown, Rollup, Webpack, and similar modern bundlers.
- React, Angular, Electron, and other app frameworks that consume ESM packages.
- Browser applications when the app is built around ESM or a bundler.

Supported package imports:

```ts
import TsPptx, { ShapeType } from "@shbernal/ts-pptx"
import { inspectPptx } from "@shbernal/ts-pptx/inspect"
import { measureText } from "@shbernal/ts-pptx/measure"
import { Presentation } from "@shbernal/ts-pptx/read"
import { readModelToIr, printScript } from "@shbernal/ts-pptx/script"
import { latexToOmml } from "@shbernal/ts-pptx/math"
import { tableToSlides } from "@shbernal/ts-pptx/html"
import pptxgenNode from "@shbernal/ts-pptx/node"
import pptxgenBrowser from "@shbernal/ts-pptx/browser"
```

## Which Build The Bare Import Gives You

`import TsPptx from "@shbernal/ts-pptx"` resolves through export conditions, so
the artifact you get depends on the runtime doing the resolving:

| the resolver sets | you get | `writeFile` |
| --- | --- | --- |
| `node` | `dist/node.js` | writes to disk via `node:fs` |
| `browser` (bundlers, `--conditions=browser`) | `dist/browser.js` | triggers a download |
| neither — Deno, Bun, edge workers | `dist/index.js` | throws `runtime/file-output-unavailable` |

Types resolve through the same condition as the code, so what TypeScript shows
you is what that runtime actually has.

The third row is the runtime-agnostic build. Authoring is identical to the other
two, and everything that hands bytes back to you — `write()`, `stream()`,
`toParts()` — works normally; a worker that returns a `.pptx` in a response body
needs nothing else. What it cannot do is *place a file for you*: there is no
filesystem and no DOM, so `writeFile()` throws an `UnsupportedFeatureError` naming
the two entries that can, instead of failing on a missing `document` deep inside
the call. Live-DOM `tableToSlides` is likewise browser-only; the DOM-agnostic form
is the free `tableToSlides` on `@shbernal/ts-pptx/html`.

Import `@shbernal/ts-pptx/node` or `@shbernal/ts-pptx/browser` directly whenever
you want a specific build regardless of how conditions resolve.

## What "Browser" Is Tested To Mean

The browser build is exercised in CI, by the `browser` job in `.github/workflows/ci.yml`
(`pnpm run test:browser` — Playwright, headless Chromium, driving `demos/vite-demo`).
It is not "supported by construction".

Two claims, kept separate on purpose:

- **The browser is a supported *runtime*.** A real browser runs the emission core
  and produces a `.pptx` you can download. The stronger form of that assertion is
  what CI actually checks: the demo imports the same showcase module the Node
  target builds, and the deck the browser assembles is compared **part for part**
  against the Node-built one. They are byte-identical. So every serializer, the zip
  writer, part ordering and relationship numbering are runtime-invariant — not by
  inspection, by comparison.
- **Browser *layout* is not an oracle this library answers to.** Nothing in that
  job depends on a rendered page. Real `offsetWidth` after layout, the resolved CSS
  cascade, and fonts as the browser chose them remain out of active scope — see
  [Project Target](project-target.md). In particular, `tableToSlides()` runs
  anywhere there is a DOM, and only *measurement* degrades without a layout engine:
  `offsetWidth` is `0`, column widths fall back to computed CSS widths and then to
  an equal split, and `data-pptx-width` / `data-pptx-min-width` let you pin them.

A layout difference between two browsers is therefore not a defect in this
package's browser support. A `.pptx` a browser builds differently from Node is.

Three adapter functions are not yet covered by that job, and the docs should not
imply otherwise: `loadMedia`, `createSvgPngPreview`, and `loadFontData`. The
showcase the demo builds draws every asset rather than loading one, so it never
reaches them.

## Dropped Compared To Upstream

### CommonJS

CommonJS is not a supported package target.

Unsupported:

```js
const TsPptx = require("@shbernal/ts-pptx")
```

The package does not ship:

- `dist/pptxgen.cjs.js`
- a CJS export condition
- a CJS-specific Node demo target

Modern Node.js versions can sometimes load ESM packages through `require()` as a
runtime interop feature. That behavior is not this package's maintained API. The
package smoke test verifies the actual contract: no CJS artifacts and no
`require` export condition.

### IIFE And Global Browser Bundle

The IIFE/global browser build is not supported.
`@shbernal/ts-pptx/browser` is an ESM browser entry, not a
`window.TsPptx` global.

Unsupported:

```html
<script src="pptxgen.bundle.js"></script>
<script>
  const pptx = new TsPptx()
</script>
```

The package does not ship:

- `dist/pptxgen.bundle.js`
- `dist/pptxgen.bundle.js.map`
- `dist/pptxgen.min.js`
- `dist/pptxgen.min.js.map`

Classic CDN script tags and `window.TsPptx` are legacy upstream workflows,
not the supported package target for this project.
The legacy upstream browser demo for that workflow is not included in this
repository.

## Artifact Name Changes

The old named ESM artifacts `dist/pptxgen.es.js` and `dist/pptxgen.js` are not
shipped. Use the package exports instead of direct artifact paths.

The maintained browser integration target is a module-aware app such as
`demos/vite-demo`.
