# ts-pptx

[![npm](https://img.shields.io/npm/v/%40shbernal%2Fts-pptx)](https://www.npmjs.com/package/@shbernal/ts-pptx)
[![weekly downloads](https://img.shields.io/npm/dw/%40shbernal%2Fts-pptx.svg?label=npm%20downloads&logo=npm)](https://www.npmjs.com/package/@shbernal/ts-pptx)
[![total downloads](https://img.shields.io/npm/dt/%40shbernal%2Fts-pptx.svg?label=npm%20total%20downloads&logo=npm)](https://www.npmjs.com/package/@shbernal/ts-pptx)
[![CI](https://github.com/shbernal/ts-pptx/actions/workflows/ci.yml/badge.svg)](https://github.com/shbernal/ts-pptx/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

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

### Installing an unreleased commit

Any commit is installable directly from GitHub, without waiting for a release:

```bash
npm install github:shbernal/ts-pptx#<commit-sha>
```

`master` (`github:shbernal/ts-pptx`) works too, but pin the sha — a branch spec
re-resolves to whatever is at the head of it when the lockfile is next written.

`dist/` is not committed, so this builds the package on install: your package
manager clones the repo, installs this package's `devDependencies`, and runs its
`prepare` script. That makes the install slow (a couple of minutes) and heavier
than a registry install, and it needs a working Node toolchain. It is meant for
trying a fix before it ships, not for production dependencies.

Note that `pres.version` reports the version in `package.json` at that commit, so
several different commits report the same number. The sha in your `package.json`
is what identifies the build.

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

- **Live-DOM / browser-layout features** — anything whose answer comes from a
  *rendered* page: real `offsetWidth` after layout, the resolved cascade, fonts as
  the browser actually chose them. (Converting an HTML `<table>` is *not* in this
  category — see [HTML tables → slides](#html-tables--slides) below. Only real
  measurement needs a browser.)
- **Third-party office-suite interop quirks** that appear only after a file is
  round-tripped through another application (for example, copy/paste inside WPS
  Office, then opening in PowerPoint) when the generated package is itself valid
  OOXML. The supported bar is that output opens cleanly in Microsoft PowerPoint.

**Contributions in these areas are welcome** — issues and pull requests are
encouraged even though the maintainer is not actively developing them. See
[`docs/project-target.md`](docs/project-target.md) for the full scope statement
and suggested testing approaches.

## Hit A Bug? There Is A Skill For That

Most code that uses this library is written by an agent, and an agent that hits a
library defect will usually route around it silently — so the defect is never
reported and never fixed. `ts-pptx-upstream` is a skill that turns that moment into
a filed issue with a minimal reproduction, which is what becomes a permanent
regression test here. It ships inside the package, so it is already on disk:

```bash
# Name the skill and the runtimes, and take the defaults: this is the form that
# completes unattended, which is how an agent will be running it.
npx skills add ./node_modules/@shbernal/ts-pptx -s '*' -a claude-code -a codex -a universal -y

npx skills add shbernal/ts-pptx   # same flags, straight from the repo instead of node_modules
```

Drop the flags for an interactive prompt if you are at a terminal yourself. Do not reach
for `--all` to avoid the prompt — it installs into every runtime the CLI knows about,
around seventy of them, and leaves an `agent/` directory at your repository root for
runtimes nobody there uses. Name the ones you have.

Two things about living with the installed copy. It is a copy, so **a version bump does
not update it** — re-run the command above, or `npx skills update ts-pptx-upstream`, in
the same commit as the bump. And if you track it, track `skills-lock.json` and ignore the
copy: `skills experimental_install` restores the file from that lock but creates none of
the runtime links, so it is a record, not a restore command.

It covers triage (is this ours, your deck's, or out of scope?), reducing a failure to
a script that builds its own deck, and — because presentations carry client names and
unreleased numbers — never uploading one to a public tracker. Once the reproduction
stands on its own it files without interrupting you, and tells you the issue number
afterwards. It also covers the far end of the cycle, which is the half that usually
rots: when a release lands, finding every workaround it retires and deleting them.

You do not need it to report something: <https://github.com/shbernal/ts-pptx/issues>
is open, and errors the library knows are its own fault print that link themselves.
See [errors](docs/errors.md#which-failures-are-worth-reporting) for which failures are
worth a report.

## Runtime And Package Support

The package is ESM-only.

Supported package surface:

- `import TsPptx, { ShapeType } from "@shbernal/ts-pptx"` (enums, shared types, and
  layout helpers ship from the main entry)
- `import { inspectPptx } from "@shbernal/ts-pptx/inspect"`
- `import { measureText } from "@shbernal/ts-pptx/measure"`
- `import { Presentation } from "@shbernal/ts-pptx/read"`
- `import { latexToOmml } from "@shbernal/ts-pptx/math"`
- `import { tableToSlides } from "@shbernal/ts-pptx/html"`
- `import TsPptx from "@shbernal/ts-pptx/node"`
- `import TsPptx from "@shbernal/ts-pptx/browser"`
  — the bare specifier resolves to one of these two by export condition, so Node
  and bundled browser apps get the right build without naming it. A runtime that
  sets neither (Deno, Bun, edge workers) gets a runtime-agnostic build that
  authors and exports bytes normally but has no `writeFile` destination; see
  [Runtime And Package Support](docs/runtime-and-package-support.md#which-build-the-bare-import-gives-you).
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

## HTML Tables → Slides

`@shbernal/ts-pptx/html` reproduces an existing HTML `<table>` as a PowerPoint
table, auto-paging across as many slides as its rows need. It works in the
browser and under Node with any DOM implementation, from one artifact.

```ts
import { TsPptx } from '@shbernal/ts-pptx'
import { tableToSlides } from '@shbernal/ts-pptx/html'
import { Window } from 'happy-dom'

const win = new Window()
win.document.body.innerHTML = '<table id="report">…</table>'

const pptx = new TsPptx()
tableToSlides(pptx, win.document.getElementById('report'))
await pptx.writeFile({ fileName: 'report.pptx' })
```

Pass the element itself and no global DOM is consulted at all. To pass a string
id instead, say which document it belongs to:

```ts
tableToSlides(pptx, 'report', { document: win.document })
```

In a browser, `options.document` defaults to the global `document`, so
`tableToSlides(pptx, 'report')` is enough. The equivalent method form,
`pptx.tableToSlides('report', options)`, remains on the browser build and
delegates to the same implementation.

**Column widths need a layout engine.** In a browser the columns are sized from
each cell's rendered `offsetWidth`, reproducing the table's real proportions.
Nothing outside a browser lays a table out, so `offsetWidth` is `0` there and
the conversion falls back in two steps: it uses the computed CSS `width`s when the
stylesheet states them for every column in one unit (all `px` or all `%`), and
an equal split when it does not.

The first step is a *fallback*, not a graceful loss of precision: `offsetWidth` is
the border box and computed `width` the content box, so padding alone can put the
two bases in different proportions. The same table can therefore come out with
different column widths in a browser and outside one.

To pin widths regardless of runtime, annotate the `<thead>` header cells — these
win outright on every path:

```html
<thead>
  <tr>
    <th data-pptx-width="2.5">Name</th>
    <th data-pptx-min-width="1">Qty</th>
  </tr>
</thead>
```

Everything else behaves the same wherever it runs: cell text (with `<br>` kept
as a line break), `colspan`/`rowspan`, computed colors, weight, alignment,
padding and borders, and auto-paging.

## Documentation

The full documentation site is published at
**<https://shbernal.github.io/ts-pptx/>**, including the generated API reference.

- [Documentation index](docs/README.md)
- [Project target](docs/project-target.md)
- [Runtime and package support](docs/runtime-and-package-support.md)
- [Development guide](docs/development.md)
- [Testing guide](docs/testing.md)
- [Agent development guide](docs/agent-development.md)
- [OOXML agent context](docs/ooxml-agent-context.md)
- [Evidence and fixtures](docs/evidence-and-fixtures.md)

## Repository Development

This repository uses `pnpm` and requires Node.js `>=24`.

```bash
pnpm install
pnpm run verify
```

OOXML serialization changes should also add or update a schema fixture in
`test/schema-cases.js`. The schema suite needs the validator installed once:

```bash
./tools/ooxml-validator/install.sh
```

Package-boundary changes should run:

```bash
pnpm run check:package
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
