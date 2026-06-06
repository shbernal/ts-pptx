# PptxGenJS

PptxGenJS generates PowerPoint `.pptx` files from TypeScript and modern
JavaScript. This maintained project targets ESM package consumers, typed
application code, reproducible package verification, and agent-assisted OOXML
development.

## Project Target

- Generate standards-based PowerPoint `.pptx` packages without requiring
  PowerPoint at runtime.
- Support TypeScript-first workflows with checked declarations and modern
  bundler resolution.
- Ship a small, explicit ESM package boundary for Node.js, Vite, React,
  Angular, Electron, and similar modern toolchains.
- Keep OOXML changes grounded in fixtures, schema validation, and PowerPoint
  compatibility evidence.
- Make the repository practical for human and agent-driven maintenance.

## Install

```bash
pnpm add pptxgenjs
```

```bash
npm install pptxgenjs
```

```bash
yarn add pptxgenjs
```

## Quick Start

```ts
import pptxgen from "pptxgenjs"

const pptx = new pptxgen()
const slide = pptx.addSlide()

slide.addText("Hello from PptxGenJS", {
  x: 1,
  y: 1,
  w: 8,
  h: 1,
  fontSize: 24,
  color: "363636",
})

await pptx.writeFile({ fileName: "example.pptx" })
```

## What It Can Generate

- Slides, layouts, masters, sections, notes, and metadata.
- Text, tables, shapes, images, SVGs, charts, and media.
- Browser-downloadable, streamed, buffered, Blob, base64, or file outputs,
  depending on the runtime.
- OOXML that is intended to open cleanly in Microsoft PowerPoint and other
  `.pptx` consumers such as Keynote, LibreOffice Impress, and Google Slides
  import.

## Runtime And Package Support

The package is ESM-only.

Supported package surface:

- `import pptxgen from "pptxgenjs"`
- `import { ShapeType } from "pptxgenjs/core"`
- `import pptxgen from "pptxgenjs/node"`
- `import pptxgen from "pptxgenjs/browser"`
- `import pptxgen from "pptxgenjs/standalone"`
- generated runtime and declaration artifacts under `dist/`
- Node.js `>=24`
- modern bundlers and module-aware app frameworks

Dropped compared to upstream:

- No CommonJS support: no `require("pptxgenjs")`, no CJS export condition, and
  no `dist/pptxgen.cjs.js`. Modern Node.js may provide `require()` interop for
  ESM, but it is not a maintained API.
- No IIFE/global browser bundle: no `window.PptxGenJS` classic script API, no
  `dist/pptxgen.bundle.js`, and no `dist/pptxgen.min.js`.

The old named ESM artifact `dist/pptxgen.es.js` is also no longer shipped. Use
the package exports instead of direct artifact paths.

See [runtime and package support](docs/runtime-and-package-support.md) for the
complete support contract.

## Documentation

- [Documentation index](docs/README.md)
- [Project target](docs/project-target.md)
- [Runtime and package support](docs/runtime-and-package-support.md)
- [Development guide](docs/development.md)
- [Testing guide](docs/testing.md)
- [Agent development guide](docs/agent-development.md)
- [OOXML agent context](docs/ooxml-agent-context.md)
- [Legacy autoloop workflow](docs/legacy-autoloop.md)

## Repository Development

This repository uses `pnpm` and requires Node.js `>=24`.

```bash
pnpm install
pnpm run build
pnpm run typecheck
pnpm run test:unit
```

OOXML serialization changes should also add or update a schema fixture and run:

```bash
./tools/ooxml-validator/install.sh
pnpm run test:schema
```

Package-boundary changes should run:

```bash
pnpm run build
pnpm run package:lint
pnpm run pack:check
pnpm run test:package
```

## Demos

- `demos/node` exercises Node.js ESM generation and stream output.
- `demos/vite-demo` exercises a modern React, TypeScript, and Vite app.

## Relationship To Upstream

This project builds on PptxGenJS by Brent Ely and contributors. The modernized
package target is intentionally narrower than upstream in order to simplify the
runtime contract and keep maintenance focused.

## License

Copyright (c) 2015-present Brent Ely and PptxGenJS contributors.

[MIT](LICENSE)
