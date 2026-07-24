# ts-pptx

ts-pptx generates PowerPoint `.pptx` files from TypeScript and modern
JavaScript. This project targets ESM package consumers, typed application code,
reproducible package verification, and agent-assisted OOXML development.

> **Lineage.** ts-pptx descends from [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS)
> (MIT), detached at its v6.0.0 (June 2025) and developed independently since. It
> does not track, merge from, or mirror the original project, and pursues its own
> Node-first direction (see _Project Target_). The original work remains under its
> MIT license — see _License_ below.

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
pnpm add @shbernal/ts-pptx
```

```bash
npm install @shbernal/ts-pptx
```

```bash
yarn add @shbernal/ts-pptx
```

## Quick Start

```ts
import TsPptx from "@shbernal/ts-pptx"

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

## What It Can Generate

- Slides, layouts, masters, sections, notes, and metadata.
- Text, tables, shapes, images, SVGs, charts, and media.
- Browser-downloadable, streamed, buffered, Blob, base64, or file outputs,
  depending on the runtime.
- OOXML that is intended to open cleanly in Microsoft PowerPoint and other
  `.pptx` consumers such as Keynote, LibreOffice Impress, and Google Slides
  import.

## Scope And Contributions

This project is **Node-first**: it generates and is tested without a browser or any
office application. Two areas are out of *active* maintenance scope — not because
they lack merit, but because there is no in-house use case driving them, so the
maintainer generally will not pick up bugs or feature requests there:

- **Live-DOM / browser-layout features**, such as `tableToSlides()`, which scrapes
  a rendered HTML `<table>` (its on-screen column widths and computed CSS) and only
  works in a real browser. The in-memory `addTable(rows, options)` API is the
  supported, fully-tested way to build tables.
- **Third-party office-suite interop quirks** that appear only after a file is
  round-tripped through another application (for example, copy/paste inside WPS
  Office, then opening in PowerPoint) when the generated package is itself valid
  OOXML. The supported bar is that output opens cleanly in Microsoft PowerPoint.

**Contributions in these areas are welcome** — issues and pull requests are
encouraged even though the maintainer is not actively developing them. See
[`docs/project-target.md`](docs/project-target.md) for the full scope statement
and suggested testing approaches.

## Runtime And Package Support

The package is ESM-only.

Supported package surface:

- `import TsPptx from "@shbernal/ts-pptx"`
- `import { ShapeType } from "@shbernal/ts-pptx/core"`
- `import { inspectPptx } from "@shbernal/ts-pptx/inspect"`
- `import { measureText } from "@shbernal/ts-pptx/measure"`
- `import { Presentation } from "@shbernal/ts-pptx/read"`
- `import { latexToOmml } from "@shbernal/ts-pptx/math"`
- `import TsPptx from "@shbernal/ts-pptx/node"`
- `import TsPptx from "@shbernal/ts-pptx/browser"`
- `import TsPptx from "@shbernal/ts-pptx/standalone"`
- generated runtime and declaration artifacts under `dist/`
- Node.js `>=24`
- modern bundlers and module-aware app frameworks

Deliberately not supported:

- No CommonJS: no `require("@shbernal/ts-pptx")`, no CJS export condition, and no
  `.cjs` artifact. Modern Node.js may provide `require()` interop for ESM, but it
  is not a maintained API.
- No IIFE/global browser bundle: no classic-script global, no `dist/*.bundle.js`,
  and no `dist/*.min.js`.

Use the package exports rather than direct `dist/` artifact paths.

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
- [Backlog workflow](docs/backlog-workflow.md)

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

ts-pptx builds on the original work of Brent Ely and the gitbrent/PptxGenJS
contributors. The modernized package target is intentionally narrower than the
original in order to simplify the runtime contract and keep maintenance focused.

## License

Copyright (c) 2015-2022 Brent Ely.
Modifications copyright (c) 2026 shbernal.

[MIT](LICENSE)
