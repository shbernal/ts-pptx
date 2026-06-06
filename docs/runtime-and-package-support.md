# Runtime And Package Support

This project ships an ESM-only package for TypeScript and modern JavaScript
applications.

## Supported Surface

Use the package export:

```ts
import pptxgen from "pptxgenjs"
```

The package publishes:

- `dist/pptxgen.js` as the runtime ESM entry.
- `types/pptxgen.d.ts` as the TypeScript declaration entry.
- `exports["."].import` for ESM consumers.
- `exports["."].types` for declaration consumers.
- `main` and `module` pointing at the ESM runtime entry for compatibility with
  tools that still inspect those fields.

Supported environments are modern module-aware environments:

- Node.js `>=24`.
- Vite, Rollup, Webpack, and similar modern bundlers.
- React, Angular, Electron, and other app frameworks that consume ESM packages.
- Browser applications when the app is built around ESM or a bundler.

## Dropped Compared To Upstream

### CommonJS

CommonJS is not supported.

Unsupported:

```js
const pptxgen = require("pptxgenjs")
```

The package does not ship:

- `dist/pptxgen.cjs.js`
- a CJS export condition
- a CJS-specific Node demo target

The package smoke test intentionally verifies that `require("pptxgenjs")` does
not resolve.

### IIFE And Global Browser Bundle

The IIFE/global browser build is not supported.

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

Direct CDN script tags and `window.PptxGenJS` are legacy upstream workflows, not
the supported package target for this project.

## Artifact Name Changes

The old named ESM artifact `dist/pptxgen.es.js` is not shipped. The ESM runtime
artifact is `dist/pptxgen.js`.

## Legacy Browser Demo

The repository still contains `demos/browser` from the upstream browser-demo
lineage. Treat it as legacy reference material until it is either modernized or
removed. The maintained browser integration target is a module-aware app such
as `demos/vite-demo`.
