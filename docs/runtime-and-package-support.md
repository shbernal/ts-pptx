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
- `dist/node.js` and `dist/node.d.ts` for explicit Node.js consumers.
- `dist/browser.js` and `dist/browser.d.ts` for explicit browser consumers.
- package `exports` entries for `.`, `./inspect`, `./measure`, `./read`,
  `./script`, `./math`, `./zip`, `./node`, and `./browser`.

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
import pptxgenNode from "@shbernal/ts-pptx/node"
import pptxgenBrowser from "@shbernal/ts-pptx/browser"
```

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
