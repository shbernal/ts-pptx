---
doc-schema-version: 1
title: "Project Target"
summary: "Current goals, non-goals, and maintenance posture for the PptxGenJS fork."
read_when:
  - Deciding whether a feature fits this fork
  - Updating package target or compatibility policy
  - Explaining current project goals
doc_type: "decision"
---

# Project Target

PptxGenJS generates PowerPoint `.pptx` packages from TypeScript and modern
JavaScript. The project target is a maintained, ESM-first library for
applications that need to create presentations programmatically.

## Goals

- Generate `.pptx` packages without requiring PowerPoint at runtime.
- Keep the public package boundary explicit and easy to verify.
- Provide TypeScript declarations that work in modern app code.
- Support Node.js `>=24` and modern bundler-driven front-end applications.
- Preserve broad OOXML feature coverage: slides, text, tables, charts, images,
  SVGs, media, and masters.
- Make OOXML changes testable through regression tests, schema fixtures, and
  package-level smoke tests.
- Support agent-driven maintenance by documenting local evidence, validation
  commands, and OOXML research paths.

## Non-Goals

- Shipping a CommonJS build.
- Shipping a standalone IIFE/global browser build.
- Supporting direct CDN script tags as the primary browser story.
- Rebuilding the upstream release matrix around every historical artifact name.
- Treating generated `dist/` outputs as hand-edited source.

## Out Of Active Scope (Contributions Welcome)

The project is **Node-first**: the generator runs and is tested without a browser
or any office application. Two domains fall outside what the maintainer actively
develops, because there is no in-house use case driving them. They are **not
rejected on merit** — they are simply not on the maintenance roadmap, and the
maintainer will generally not pick up bugs or feature requests in these areas.
**Pull requests that fix or extend them are welcome** (ideally with the testing
approach noted below), and the OOXML-emission core they build on is fully
supported.

- **Live-DOM / browser-layout features.** Anything that reads a *rendered* web
  page rather than in-memory data — most notably `tableToSlides()`, which scrapes
  a live `<table>` and copies its rendered column widths (`offsetWidth`) and
  computed CSS styles (`window.getComputedStyle`). These only work in a real
  browser and cannot be reproduced in the Node test suite. New browser-rendering
  features are out of active scope. The in-memory `addTable(rows, opts)` path is
  the supported, fully-tested way to build tables.

  *Contributor note:* the established pattern (see `resolveHtmlColWidth` in
  `src/gen-tables.ts` and `test/regression/html-table-col-width.test.js`) is to
  extract the DOM-independent logic into pure helpers and unit-test those with
  synthetic inputs; full-fidelity layout repros need a headless browser
  (Playwright/Puppeteer), which is not currently a project dependency.

- **Third-party office-suite interop quirks.** Bugs that only appear after a file
  is round-tripped through another application (for example, copy/paste inside WPS
  Office, then opening the result in PowerPoint) are out of active scope when the
  generated package is itself valid OOXML and the corruption is introduced by the
  other application. The supported compatibility bar is that output opens cleanly
  in Microsoft PowerPoint; cleanly opening in Keynote, LibreOffice Impress, and
  Google Slides import is a best-effort goal.

  *Contributor note:* a worked repro that pins the defect to invalid OOXML the
  library *itself* emits (independent of the other application) turns one of these
  into an in-scope correctness bug.

## Maintenance Posture

The repository should be understandable to a maintainer or an agent starting
from a clean checkout:

- package support is documented in `docs/runtime-and-package-support.md`;
- development commands are documented in `docs/development.md`;
- verification commands are documented in `docs/testing.md`;
- OOXML source-of-truth lookup is documented in `docs/ooxml-agent-context.md`.
