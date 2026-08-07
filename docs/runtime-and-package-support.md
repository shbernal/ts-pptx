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
(`pnpm run test:browser` — Playwright, headless Chromium). It is not "supported by
construction". Three fixtures: `demos/vite-demo` for the bundled path a real consumer
takes, a static server handing the browser the shipped `dist/browser.js` unbundled
for the runtime adapter itself, and a page that renders a real `<table>` so
`tableToSlides` reads a measured `offsetWidth`.

Two claims, kept separate on purpose:

- **The browser is a supported *runtime*.** A real browser runs the emission core
  and produces a `.pptx` you can download. The stronger form of that assertion is
  what CI actually checks: the demo imports the same showcase module the Node
  target builds, and the deck the browser assembles is compared **part for part**
  against the Node-built one. They are byte-identical. So every serializer, the zip
  writer, part ordering and relationship numbering are runtime-invariant — not by
  inspection, by comparison.
- **Browser *layout* is not an oracle this library answers to.** The resolved CSS
  cascade and fonts as the browser chose them remain out of active scope — see
  [Project Target](project-target.md). `tableToSlides()` runs anywhere there is a
  DOM, and only *measurement* is lost without a layout engine: `offsetWidth` is
  `0`, column widths fall back to computed CSS widths and then to an equal split,
  and `data-pptx-width` / `data-pptx-min-width` let you pin them.

  Losing the measurement is not the same as losing precision, and this page used
  to say "degrades" as though it were. `offsetWidth` is the border box; computed
  `width` is the content box. Padding alone is enough to make the two disagree —
  the `html-table` fixture is built to, at 1:1 measured against 2:1 from CSS — so
  one table converted in Chromium and under happy-dom can emit different column
  *proportions*, not the same proportions coarsened. Where both runtimes must
  agree on a column, state it with `data-pptx-width`.

  One part of the job does now drive a rendered page, and the line it holds is
  worth stating exactly. A `<table>` is laid out in Chromium and converted, and the
  lane asserts that the measured `offsetWidth` is what sizes the emitted columns —
  that the measurement is *taken and honoured*, proportionally, with
  `data-pptx-width` still overriding it. It asserts nothing about whether that
  measurement is *correct*, or whether Firefox would agree. The library's contract
  is "we use what your DOM reports"; it is not "your DOM reports what PowerPoint
  will draw".

A layout difference between two browsers is therefore not a defect in this
package's browser support. A `.pptx` a browser builds differently from Node is.

### The Runtime Adapter, Function By Function

Everything that differs between Node and the browser lives in one four-function
`RuntimeAdapter`. All four now run in a real Chromium, and what each is checked
against is worth stating precisely, because "covered" is a weaker word than what
these actually assert:

| adapter function | what the browser lane checks |
| --- | --- |
| `writeFile` | the object-URL `<a download>` fires and the downloaded bytes unzip to a real OPC package |
| `loadMedia` | a fetched image lands in the package as **the same bytes** Node reads off disk — and as the same bytes as the source file. A 404 fails the export with `media/fetch-failed` as the cause of `media/load-failed` |
| `createSvgPngPreview` | the `<canvas>` rasterizer emits a real PNG where Node can only stub a placeholder; an undecodable SVG and a zero-dimension SVG each fail rather than shipping a blank fallback |
| `loadFontData` | a font fetched over HTTP measures to the same baked `fontScale` and embeds the same `/ppt/fonts/` bytes as one read off disk. A 404 rejects with `font/fetch-failed` |

Two of those are cross-runtime comparisons run through the byte-identity gate's
own machinery, so "the same bytes" means the same thing here as it does there.

The one place the two runtimes are *expected* to disagree is
`createSvgPngPreview`: Node has no rasterizer, so it writes a fixed placeholder
into the PNG fallback rel where a browser draws the artwork. That is a documented
divergence rather than a bug, and the lane asserts its exact shape — one changed
part, and the browser's is a real PNG — so it cannot quietly become a different
divergence.

### Which Browsers The Lane Runs

Chromium, and only Chromium. This is a decision, not an oversight, and it is
recorded here so it does not get re-opened every time CI time is discussed.

The adapter surface above is `fetch`, `FileReader`, `<canvas>`, object URLs and
`<a download>`. None of those is a corner of the platform where engines are known
to disagree, and no divergence has been reported against this package or observed
while building the lane. A Firefox and WebKit matrix would therefore triple the
job to keep re-answering a question nothing has asked.

Add an engine when there is something concrete to add it for: a reported
difference, or a new adapter function that touches an API with a real
cross-engine history. Not pre-emptively.

### What The Lane Does Not Cover

Two gaps, stated rather than implied:

- **Live-DOM layout**, as above — deliberate, and the subject of
  [Project Target](project-target.md).
- **Two arms of `createSvgPngPreview`**: a missing 2d context and a
  `toDataURL` that throws. Neither is reachable in a browser that has a working
  canvas and is drawing a same-origin data URI; reaching them means stubbing DOM
  constructors, which asserts about the stub. The lane's own coverage floor
  accounts for them (see [Testing](testing.md#browser-lane)).

## What `/math` Costs In A Browser, And Why It Stays Node-Only

`@shbernal/ts-pptx/math` is Node-only, permanently, unless a real consumer asks
otherwise. The decision is recorded here so it is not re-litigated per release.

`src/math.ts` loads its two optional peers (`temml`, `mathml2omml`) through
`node:module`'s `createRequire`. That is what keeps `latexToOmml()` and
`mathmlToOmml()` **synchronous**. A browser has no `createRequire`, and the only
browser-compatible replacement is a dynamic `import()`, which makes both
functions async — a breaking change to a published API, paid by every existing
caller, to serve a use case nobody has raised.

The subpath is already documented as Node-only at the top of the module. If a
browser consumer does turn up, the answer is an additional `/math/async`
subpath, not a change to this one.

Nothing else in the package has this problem: `src/runtime/node.ts` is the only
other file importing `node:*`, and it is contained behind the `RuntimeAdapter`.
(`dist/zip.js` also carries a lazy `import('node:fs/promises')`, which a bundler
will warn about; it is on the read-a-package-from-a-path branch only and never
executes on the write path.)

## Bundle Size

`scripts/bundle-size-ratchet.mjs` freezes a budget for the browser entry and its
chunks, gzipped, and `pnpm run check:package` enforces it. Read the number as a
**growth detector, not a download size**: `dist/` is unminified, and every real
browser consumer runs it through a bundler that minifies before serving, so what
anyone actually downloads is well under the figure the gate prints. What the gate
is for is the step change — a dependency reaching the browser entry, or a chunk
split going wrong.

`pnpm run bundle-size:list` prints the per-chunk breakdown; the budget lives in
`scripts/bundle-size-budget.json` and is raised or lowered deliberately with
`pnpm run bundle-size:freeze`.

## Using The Browser Entry Without A Bundler

Supported environments assume a bundler, and that remains the maintained target.
But `dist/browser.js` does load in a browser as-is, over a plain
`<script type="module">`, provided you resolve the two bare specifiers it reaches
— which is exactly what the adapter harness does
(`test/browser/harness/index.html`):

```html
<script type="importmap">
  {
    "imports": {
      "fflate": "/node_modules/fflate/esm/browser.js",
      "opentype.js": "/node_modules/opentype.js/dist/opentype.mjs"
    }
  }
</script>
```

`opentype.js` is a *dynamic* import inside the measure/fit chunk — nothing
requests it until a font is registered, so an app that never calls
`registerFontMetrics` or `embedFont` will not notice its absence until it does.

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
