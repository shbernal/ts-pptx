---
doc-schema-version: 1
title: "Runtime and package support"
summary: "Supported imports, how every runtime loads the one ESM build, and shipped artifacts."
read_when:
  - Changing package exports or runtime support
  - Explaining how the ESM build loads in Node, a browser, or from CommonJS
  - Updating shipped artifact policy
doc_type: "reference"
---

# Runtime and package support

This project ships one ESM package for TypeScript and modern JavaScript applications.
Node, bundlers, browsers and `require()` callers all load that one build.

## Supported surface

Use the package export:

```ts
import TsPptx from "pptx-ts"
```

The package publishes:

- `dist/index.js` and `dist/index.d.ts` as the default ESM package entry. It also
  exports the public enums, shared types, layout constants, and unit helpers.
  Under Node, and in a browser bundle, the same bare import resolves to
  `dist/node.js` or `dist/browser.js` instead. See
  [Which build the bare import gives you](#which-build-the-bare-import-gives-you).
- `dist/inspect.js` and `dist/inspect.d.ts` for low-level PPTX package
  inspection, slide/object extraction, and geometry helpers.
- `dist/measure.js` and `dist/measure.d.ts` for headless text-measurement and
  autofit helpers (see [Measured text fit](measured-text-fit.md)).
- `dist/read.js` and `dist/read.d.ts` for opening, editing, and round-tripping
  an existing `.pptx` (see [PPTX read and round-trip](reference/pptx-read.md)).
- `dist/script.js` and `dist/script.d.ts` for turning an existing `.pptx` into
  TypeScript source that rebuilds it through the write API (see
  [PPTX to script](reference/pptx-to-script.md)).
- `dist/math.js` and `dist/math.d.ts` for LaTeX/MathML → OMML conversion (see
  [Math and LaTeX](math-latex.md)).
- `dist/zip.js` and `dist/zip.d.ts` for the shared OPC/zip package plumbing
  used by `read` and `inspect`.
- `dist/html.js` and `dist/html.d.ts` for converting an existing HTML `<table>`
  into slides. One artifact serves both runtimes: there is deliberately no
  `browser`/`node` condition split, because the entry works with whatever DOM
  the caller has (see [HTML tables to slides](html-tables.md)).
- `dist/families.js` and `dist/families.d.ts` for the construct families a
  `createPresentation({ use })` call names (see [Bundle size](bundle-size.md)).
- `dist/node.js` and `dist/node.d.ts` for explicit Node.js consumers.
- `dist/browser.js` and `dist/browser.d.ts` for explicit browser consumers.
- package `exports` entries for `.`, `./inspect`, `./measure`, `./read`,
  `./script`, `./math`, `./zip`, `./html`, `./families`, `./node`, and
  `./browser`.

Supported environments are modern module-aware environments:

- Node.js `>=24`.
- Vite, Rolldown, Rollup, Webpack, and similar modern bundlers.
- React, Angular, Electron, and other app frameworks that consume ESM packages.
- Browser applications when the app is built around ESM or a bundler.

Supported package imports:

```ts
import TsPptx, { ShapeType } from "pptx-ts"
import { inspectPptx } from "pptx-ts/inspect"
import { measureText } from "pptx-ts/measure"
import { Presentation } from "pptx-ts/read"
import { readModelToIr, printScript } from "pptx-ts/script"
import { latexToOmml } from "pptx-ts/math"
import { tableToSlides } from "pptx-ts/html"
import { charts, tables } from "pptx-ts/families"
import pptxgenNode from "pptx-ts/node"
import pptxgenBrowser from "pptx-ts/browser"
```

## Paying for what you author

Every entry publishes two ways to start a deck. `TsPptx` is composed with every
construct family the library has, so it authors everything and costs everything.
`createPresentation({ use })` is composed with the core tier plus the families you
name, and a family you do not name is not in your bundle:

```ts
import { createPresentation } from "pptx-ts"
import { charts } from "pptx-ts/families"

const pres = createPresentation({ use: [charts] })
pres.addSlide().addChart(data, { type: ChartType.bar })
```

Both write the same deck for the same slides, part for part. Which families are
core and which you ask for, what each one costs, and the diagnostic a method
you did not compose raises are in [Bundle size](bundle-size.md).

## Which build the bare import gives you

`import TsPptx from "pptx-ts"` resolves through export conditions, so
the artifact you get depends on the runtime doing the resolving:

| the resolver sets | you get | `writeFile` |
| --- | --- | --- |
| `node` | `dist/node.js` | writes to disk via `node:fs` |
| `browser` (bundlers, `--conditions=browser`) | `dist/browser.js` | triggers a download |
| neither: Deno, Bun, edge workers | `dist/index.js` | throws `runtime/file-output-unavailable` |

Types resolve through the same condition as the code, so what TypeScript shows
you is what that runtime actually has.

The third row is the runtime-agnostic build. Authoring is identical to the other
two, and everything that hands bytes back to you (`write()`, `toBytes()`,
`toParts()`) works normally; a worker that returns a `.pptx` in a response body
needs nothing else. What it cannot do is *place a file for you*: there is no
filesystem and no DOM, so `writeFile()` throws an `UnsupportedFeatureError` naming
the two entries that can, instead of failing on a missing `document` deep inside
the call. Live-DOM `tableToSlides` is likewise browser-only; the DOM-agnostic form
is the free `tableToSlides` on `pptx-ts/html`.

Import `pptx-ts/node` or `pptx-ts/browser` directly whenever
you want a specific build regardless of how conditions resolve.

## What "browser" is tested to mean

Nothing here is supported by construction. The `browser` job in
`.github/workflows/ci.yml` runs `pnpm run test:browser`, which is Playwright driving
headless Chromium over three fixtures. The site's own demos page (`www/demos/`), for
the bundled path a real consumer takes. A static server handing the browser the
shipped `dist/browser.js` unbundled, for the runtime adapter itself. And a page
rendering a real `<table>`, so `tableToSlides` reads a measured `offsetWidth`.

Two claims, kept separate on purpose:

- **The browser is a supported *runtime*.** A real browser runs the emission core
  and produces a `.pptx` you can download. CI checks the stronger form of that.
  The demo imports the same showcase module the Node target builds, and the deck
  the browser assembles is compared part for part against the Node-built one.
  They come out byte-identical. Every serializer, the zip writer, part ordering
  and relationship numbering are runtime-invariant by comparison, not by
  inspection.
- **Browser *layout* is not an oracle this library answers to.** The resolved CSS
  cascade, and fonts as the browser chose them, stay out of active scope (see
  [Project target](project-target.md)). `tableToSlides()` runs anywhere there is a
  DOM. Only *measurement* is lost without a layout engine. `offsetWidth` reads `0`,
  column widths fall back to computed CSS widths and then to an equal split, and
  `data-pptx-width` / `data-pptx-min-width` pin them.

  Losing the measurement is not the same as losing precision, whatever this page
  used to say. `offsetWidth` is the border box. Computed `width` is the content
  box. Padding alone makes the two disagree, and the `html-table` fixture is built
  to do exactly that, at 1:1 measured against 2:1 from CSS. So one table converted
  in Chromium and under happy-dom can emit different column *proportions*, not the
  same proportions coarsened. Where both runtimes have to agree on a column, state
  it with `data-pptx-width`.

  One part of the job does drive a rendered page, and the line it holds is worth
  stating exactly. Chromium lays a `<table>` out, the conversion runs, and the lane
  asserts that the measured `offsetWidth` is what sizes the emitted columns.
  Proportionally, with `data-pptx-width` still overriding it. It asserts nothing
  about whether that measurement is *correct*, or whether Firefox would agree. The
  contract is "we use what your DOM reports". It is not "your DOM reports what
  PowerPoint will draw".

A layout difference between two browsers is therefore not a defect in this
package's browser support. A `.pptx` a browser builds differently from Node is.

### The runtime adapter, function by function

Everything that differs between Node and the browser lives in one four-function
`RuntimeAdapter`. All four run in a real Chromium. "Covered" undersells what the
lane asserts, so here is each one exactly:

| adapter function | what the browser lane checks |
| --- | --- |
| `writeFile` | the object-URL `<a download>` fires and the downloaded bytes unzip to a real OPC package |
| `loadMedia` | a fetched image lands in the package as **the same bytes** Node reads off disk, and as the same bytes as the source file. A 404 fails the export with `media/fetch-failed` as the cause of `media/load-failed` |
| `createSvgPngPreview` | the `<canvas>` rasterizer emits a real PNG where Node can only stub a placeholder. An undecodable SVG and a zero-dimension SVG each report `media/svg-preview-failed` rather than shipping a blank fallback, by `path` or inline: the export fails under the default `onMediaError`, and under `'placeholder'` it warns, keeps the SVG's bytes, and writes the same placeholder fallback Node does |
| `loadFontData` | a font fetched over HTTP measures to the same baked `fontScale` and embeds the same `/ppt/fonts/` bytes as one read off disk. A 404 rejects with `font/fetch-failed` |

Two of those are cross-runtime comparisons run through the byte-identity gate's
own machinery, so "the same bytes" means the same thing here as it does there.

The one place the two runtimes are *expected* to disagree is
`createSvgPngPreview`: Node has no rasterizer, so it writes a fixed placeholder
into the PNG fallback rel where a browser draws the artwork. That is a documented
divergence rather than a bug, and the lane asserts its exact shape (one changed
part, and the browser's is a real PNG), so it cannot quietly become a different
divergence.

### Which browsers the lane runs

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

### What the lane does not cover

Two gaps, stated rather than implied:

- **Live-DOM layout**, as above: deliberate, and the subject of
  [Project target](project-target.md).
- **Two arms of `createSvgPngPreview`**: a missing 2d context and a
  `toDataURL` that throws. Neither is reachable in a browser that has a working
  canvas and is drawing a same-origin data URI; reaching them means stubbing DOM
  constructors, which asserts about the stub. The lane's own coverage floor
  accounts for them (see [Testing](testing.md#browser-lane)).

## What `/math` costs in a browser, and why it stays Node-only

`pptx-ts/math` is Node-only, permanently, unless a real consumer asks
otherwise. The decision is recorded here so it is not re-litigated per release.

`src/math.ts` loads its two optional peers (`temml`, `mathml2omml`) through
`node:module`'s `createRequire`. That is what keeps `latexToOmml()` and
`mathmlToOmml()` **synchronous**. A browser has no `createRequire`, and the only
browser-compatible replacement is a dynamic `import()`, which makes both
functions async: a breaking change to a published API, paid by every existing
caller, to serve a use case nobody has raised.

The subpath is already documented as Node-only at the top of the module. If a
browser consumer does turn up, the answer is an additional `/math/async`
subpath, not a change to this one.

Nothing else in the package has this problem: `src/runtime/node.ts` is the only
other file importing `node:*`, and it is contained behind the `RuntimeAdapter`.
(`dist/zip.js` also carries a lazy `import('node:fs/promises')`, which a bundler
will warn about; it is on the read-a-package-from-a-path branch only and never
executes on the write path.)

## Using the browser entry without a bundler

Supported environments assume a bundler, and that remains the maintained target.
Even so, `dist/browser.js` loads in a browser as-is, over a plain
`<script type="module">`, as long as you resolve the two bare specifiers it reaches.
That is exactly what the adapter harness does
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

`opentype.js` is a *dynamic* import inside the measure/fit chunk: nothing
requests it until a font is registered, so an app that never calls
`registerFontMetrics` or `embedFont` will not notice its absence until it does.

## One build, and everything that loads it

One ESM build ships, and it is what every consumer gets: Node, bundlers, browsers,
and CommonJS callers alike. This section is what that means for each of them, and
what upstream artifact names are gone.

### `require()` from CommonJS

`require("pptx-ts")` works. Node loads ESM through `require()` from 22.12 onward,
and this package floors at Node `>=24`, so the interop is always available on a
supported runtime:

```js
const { default: TsPptx, ShapeType } = require("pptx-ts")
const pptx = new TsPptx()
```

Every subpath loads the same way. The one difference from a package that ships a
separate CJS build is that `require()` returns a module *namespace*, so the class
arrives on `.default` rather than as the export itself. Destructure it, as above, and
the rest of the API reads identically.

This is tested, not incidental. `pnpm run test:package` asserts both halves of the
contract: that one build ships (no `dist/pptxgen.cjs.js`, no `require` export
condition, no CJS-specific demo target), and that `require()` resolves every published
subpath with its default and named exports intact. The interop has one failure mode
worth naming. A top-level await anywhere in an entry's chunk graph makes `require()` of
that entry throw while every ESM suite stays green, and that assertion is what catches
it.

### A browser, with a `<script>` tag

`pptx-ts/browser` is an ES module, so a browser loads it from
`<script type="module">` and binds it to whatever name the `import` gives it. An ESM
CDN resolves the dependency graph and serves it in one request:

```html
<script type="module">
  import TsPptx from "https://esm.sh/pptx-ts/browser"

  const pptx = new TsPptx()
  pptx.addSlide().addText("Built in the browser", { x: 1, y: 1, w: 8, h: 1 })
  await pptx.writeFile({ fileName: "example.pptx" })
</script>
```

jsDelivr serves the same build from `https://cdn.jsdelivr.net/npm/pptx-ts/+esm`, which
resolves to `dist/browser.js` through the same export condition a bundler uses.
[Using the browser entry without a bundler](#using-the-browser-entry-without-a-bundler)
above covers the self-hosted equivalent, where you supply the import map yourself.

The classic-script form upstream shipped, a `dist/pptxgen.bundle.js` assigning a
`window.TsPptx` global, has no equivalent here: the module script above is the whole
replacement for it, and nothing in the package writes to `window`. The named artifacts
`dist/pptxgen.bundle.js`, `dist/pptxgen.min.js`, their source maps, and the older ESM
names `dist/pptxgen.es.js` and `dist/pptxgen.js` are all gone. Reach the package
through its exports rather than an artifact path, which is what keeps a file rename
from being a breaking change.

The maintained browser integration target is a module-aware app such as the site's own
demos page (`www/demos/`).
