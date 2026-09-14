---
doc-schema-version: 1
title: "Where it runs"
summary: "Every pptx-ts entry point and what it is for, which runtimes load the one ESM build and how, what writeFile does in each, and the one Node-only subpath."
read_when:
  - Choosing which pptx-ts entry point to import
  - Loading the package from CommonJS, a browser page, Deno, Bun or an edge worker
  - Working out why writeFile behaves differently in another runtime
doc_type: "reference"
---

# Where it runs

ts-pptx ships one ESM build. Node, bundlers, browsers and `require()` all load that build, and no
office application is involved anywhere.

## Runtimes

| Runtime | Load it with | `writeFile` |
| --- | --- | --- |
| Node.js 24 or later | `import`, or [`require()`](#require-from-commonjs) | writes the file to disk |
| A browser app built with a bundler (Vite, Rollup, Webpack and others) | `import` | downloads the file |
| A browser page with no build step | [`<script type="module">`](#a-browser-with-a-script-tag) from an ESM CDN | downloads the file |
| Deno, Bun, edge workers | `import` | [throws](#deno-bun-and-edge-workers); use `toBytes()` instead |

CI runs the browser build in a real Chromium, and compares the deck it assembles part for part with
the one Node builds from the same code. They are identical.

## Entry points

| Subpath | Import | What it is for | Runs in | Guide |
| --- | --- | --- | --- | --- |
| `pptx-ts` | `import TsPptx from "pptx-ts"` | Writing decks: `TsPptx`, `createPresentation`, enums, types and unit helpers | Every runtime, through [the build for that runtime](#which-build-the-bare-import-gives-you) | [Your first deck](first-deck.md), [API reference](../reference/api/index.md) |
| `pptx-ts/node` | `import TsPptx from "pptx-ts/node"` | The Node build, whatever the resolver's conditions | Node | [API reference](../reference/api/index.md) |
| `pptx-ts/browser` | `import TsPptx from "pptx-ts/browser"` | The browser build, whatever the resolver's conditions | Browsers | [API reference](../reference/api/index.md) |
| `pptx-ts/families` | `import { charts } from "pptx-ts/families"` | Construct families for `createPresentation({ use })` | Every runtime | [Smaller bundles](../bundle-size.md) |
| `pptx-ts/read` | `import { Presentation } from "pptx-ts/read"` | Opening, editing and saving an existing deck | Every runtime. Loading from a file path needs Node | [Read and edit a deck](../reading/read-and-edit.md) |
| `pptx-ts/inspect` | `import { inspectPptx } from "pptx-ts/inspect"` | Reporting what a package holds without building the full model | Every runtime | [Inspect a package](../reference/pptx-inspection.md) |
| `pptx-ts/script` | `import { printScript, readModelToIr } from "pptx-ts/script"` | Turning a deck into the TypeScript that rebuilds it | Every runtime | [PPTX to script](../reference/pptx-to-script.md) |
| `pptx-ts/measure` | `import { measureText } from "pptx-ts/measure"` | Measuring text against real font metrics | Every runtime | [Text that fits](../text-fit.md) |
| `pptx-ts/html` | `import { tableToSlides } from "pptx-ts/html"` | Converting an HTML `<table>` into slides | Anywhere with a DOM: a browser, or Node with a DOM library such as happy-dom | [HTML tables to slides](../html-tables.md) |
| `pptx-ts/math` | `import { latexToOmml } from "pptx-ts/math"` | Turning LaTeX and MathML into native equations | [Node only](#math-is-node-only) | [Math equations](../math-latex.md) |
| `pptx-ts/zip` | `import { readZip } from "pptx-ts/zip"` | The zip reading and writing that `read` and `inspect` share | Every runtime | None |

## Which build the bare import gives you

`import TsPptx from "pptx-ts"` resolves through export conditions, so the file you get depends on the
runtime doing the resolving:

| The resolver sets | You get | `writeFile` |
| --- | --- | --- |
| `node` | `dist/node.js` | writes to disk through `node:fs` |
| `browser` (bundlers targeting the browser, `--conditions=browser`) | `dist/browser.js` | starts a download |
| neither: Deno, Bun, edge workers | `dist/index.js` | throws `runtime/file-output-unavailable` |

Types resolve through the same condition as the code, so what TypeScript shows you is what that
runtime has. To get one build whatever the conditions, import `pptx-ts/node` or `pptx-ts/browser`
directly.

## `require()` from CommonJS

```js
const { default: TsPptx, ShapeType } = require("pptx-ts")

const pptx = new TsPptx()
```

Node loads ES modules through `require()` from version 22.12, and the package requires Node 24, so
this works on every supported version. `require()` returns the module's namespace rather than a
separate CommonJS build, so the class arrives on `.default`. Every subpath loads the same way.

## A browser, with a `<script>` tag

`pptx-ts/browser` is an ES module. An ESM CDN resolves its dependencies and serves it to a module
script:

```html
<script type="module">
  import TsPptx from "https://esm.sh/pptx-ts/browser"

  const pptx = new TsPptx()
  pptx.addSlide().addText("Built in the browser", { x: 1, y: 1, w: 8, h: 1 })
  await pptx.writeFile({ fileName: "example.pptx" })
</script>
```

jsDelivr serves the same build from `https://cdn.jsdelivr.net/npm/pptx-ts/+esm`.

### Using the browser entry without a bundler

To serve the files yourself instead of from a CDN, load `dist/browser.js` from a module script and
map the two bare specifiers it imports:

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

`opentype.js` is imported only when a font is first registered, so a page that never calls
`registerFontMetrics` or `embedFont` will not notice a missing entry for it until it does.

## Deno, Bun and edge workers

These runtimes load the runtime-agnostic build, `dist/index.js`. Authoring is the same as anywhere
else, and everything that hands bytes back (`toBytes()`, `write()`, `toParts()`) works. A worker that
returns a deck in a response body needs nothing more.

What this build cannot do is place a file for you. It has no filesystem and no page to download onto,
so `writeFile()` throws an `UnsupportedFeatureError` with the code `runtime/file-output-unavailable`,
naming the two entries that can write a file.

## Math is Node-only

`pptx-ts/math` loads its two optional dependencies through Node's `createRequire`, which keeps
`latexToOmml()` and `mathmlToOmml()` synchronous, and a browser has no `createRequire`.
