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
import pptxgen from "@shbernal/pptxgenjs"
```

The package publishes:

- `dist/index.js` and `dist/index.d.ts` as the default ESM package entry.
- `dist/core.js` and `dist/core.d.ts` for public enums, shared types, layout
  constants, and unit helpers.
- `dist/inspect.js` and `dist/inspect.d.ts` for low-level PPTX package
  inspection, slide/object extraction, and geometry helpers.
- `dist/node.js` and `dist/node.d.ts` for explicit Node.js consumers.
- `dist/browser.js` and `dist/browser.d.ts` for explicit browser consumers.
- `dist/standalone.js` and `dist/standalone.d.ts` as a browser ESM entry that
  bundles JSZip.
- package `exports` entries for `.`, `./core`, `./inspect`, `./node`,
  `./browser`, and `./standalone`.

Supported environments are modern module-aware environments:

- Node.js `>=24`.
- Vite, Rolldown, Rollup, Webpack, and similar modern bundlers.
- React, Angular, Electron, and other app frameworks that consume ESM packages.
- Browser applications when the app is built around ESM or a bundler.

Supported package imports:

```ts
import pptxgen from "@shbernal/pptxgenjs"
import { ShapeType } from "@shbernal/pptxgenjs/core"
import { inspectPptx } from "@shbernal/pptxgenjs/inspect"
import pptxgenNode from "@shbernal/pptxgenjs/node"
import pptxgenBrowser from "@shbernal/pptxgenjs/browser"
import pptxgenStandalone from "@shbernal/pptxgenjs/standalone"
```

## Dropped Compared To Upstream

### CommonJS

CommonJS is not a supported package target.

Unsupported:

```js
const pptxgen = require("@shbernal/pptxgenjs")
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
`@shbernal/pptxgenjs/standalone` is an ESM browser entry, not a
`window.PptxGenJS` global.

Unsupported:

```html
<script src="pptxgen.bundle.js"></script>
<script>
  const pptx = new PptxGenJS()
</script>
```

The package does not ship:

- `dist/pptxgen.bundle.js`
- `dist/pptxgen.bundle.js.map`
- `dist/pptxgen.min.js`
- `dist/pptxgen.min.js.map`

Classic CDN script tags and `window.PptxGenJS` are legacy upstream workflows,
not the supported package target for this project.
The legacy upstream browser demo for that workflow is not included in this
repository.

## Artifact Name Changes

The old named ESM artifacts `dist/pptxgen.es.js` and `dist/pptxgen.js` are not
shipped. Use the package exports instead of direct artifact paths.

The maintained browser integration target is a module-aware app such as
`demos/vite-demo`.
