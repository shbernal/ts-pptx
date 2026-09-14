---
doc-schema-version: 1
title: "Project target"
summary: "Current goals, non-goals, and maintenance posture for this ts-pptx project."
read_when:
  - Deciding whether a feature fits this project
  - Updating package target or compatibility policy
  - Explaining current project goals
doc_type: "decision"
---

# Project target

ts-pptx generates PowerPoint `.pptx` packages from TypeScript and modern
JavaScript. One ESM library, maintained, for applications that build decks in code.

This is an independent derivative of [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS),
detached at its v4.0.1 in June 2025. Nothing syncs or mirrors from upstream any
more, and the direction here is its own. The code descends from that codebase and
keeps the original MIT copyright. It is not a drop-in continuation of the upstream
release line.

## Goals

- Generate `.pptx` packages without requiring PowerPoint at runtime.
- Keep the public package boundary explicit and easy to verify.
- Provide TypeScript declarations that work in modern app code.
- Support Node.js `>=24`, modern bundler-driven front-end applications, and a
  browser loading the module straight from an ESM CDN.
- Support the browser as a runtime, and prove it rather than assert it. CI drives
  the browser build and its runtime adapter against a real Chromium, then compares
  the deck a browser assembles part-for-part against the Node-built one. The claim
  is about *emission* and stops there. The Live-DOM bullet under Out of active
  scope draws the line.
- Preserve broad OOXML feature coverage: slides, text, tables, charts, images,
  SVGs, media, and masters.
- Make OOXML changes testable through regression tests, schema fixtures, and
  package-level smoke tests.
- Support agent-driven maintenance by documenting local evidence, validation
  commands, and OOXML research paths.

## Non-goals

- Shipping more than one build. The published ESM artifact is what Node, bundlers,
  browsers and `require()` callers all load; see [Runtime and package support](runtime-and-package-support.md#one-build-and-everything-that-loads-it).
- Reproducing the upstream release matrix: its historical artifact names, its
  `window.TsPptx` global, or compatibility with its build system. A browser reaches
  this package as a module, from a bundler or an ESM CDN.
- Treating generated `dist/` outputs as hand-edited source.

### What stays in the consumer

The bullets above are about the *shape* of the package. These are different. Each
one encodes a specific consumer's brand, content, or deck workflow. None is a
candidate here at any priority, and the line holds even when the code looks
perfectly generic. Don't raise them.

- Brand guidance, workflow-specific scripts, and consumer content.
- A consumer's default font choice.
- Icon-set policy, imports, aliases, and provenance manifests.
- External stock-asset sourcing helpers.
- Lint quality thresholds, annotated screenshots, and human-review artifacts.
- Slide semantics manifests as agent-facing design-intent contracts.
- Greenfield deck eval prompts, scorecards, and generator-adapter behavior.
- LibreOffice/ImageMagick rendering orchestration for local visual QA.

What this package owes such a consumer is sound primitives. The composition on top
of them belongs to the consumer. A generic PPTX gap uncovered while building one
*is* in scope, though. See
[Agent development guide → Promoting a downstream need](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/agent-development.md#promoting-a-downstream-need)
for how it moves across.

## Out of active scope (contributions welcome)

The project is Node-first. The generator runs and is tested with no office
application anywhere near it, and the Node suite is where nearly all of it is
proven. That says where the evidence lives. It is not a hedge about the browser,
which is a supported runtime with its own CI lane (see
[Runtime and package support](runtime-and-package-support.md#what-browser-is-tested-to-mean)).

Two domains sit outside what the maintainer actively develops, for the plain
reason that no in-house use case drives them. Neither is rejected on merit. They
are just off the roadmap, so expect bugs and feature requests in these areas to
sit. Pull requests that fix or extend them are welcome, ideally with the testing
approach noted below, and the OOXML-emission core underneath them is fully
supported.

- **Live-DOM / browser-layout features.** Anything whose *answer* comes from a
  rendered page: real `offsetWidth` after layout, the resolved cascade, fonts as
  the browser actually chose them. Reproducing those faithfully needs a real
  browser, so features that depend on them are out of active scope.

  Keep this apart from "the browser is a supported runtime". Every report lands
  in one bucket or the other. Running the library in a browser is supported and
  tested. Committing that its output *matches how a browser laid something out*
  is not, and the reason is that there is no oracle for it. Every other gate in
  this repo has one, whether schema validation, byte identity, or a PowerPoint
  render. "Renders differently in Firefox" has none, so correctness would turn
  into a judgement call per report. A `.pptx` a browser builds differently from
  Node is a defect. A layout difference between two browsers is not.

  Two triage rules, both learned the hard way.

  **A report is only a live-DOM report when the disagreement sits upstream of
  PowerPoint.** A construct PowerPoint renders as intended and another viewer
  does not belongs to the third-party bullet below. A browser-layout oracle
  would not answer it, because nothing about the rendered page is in dispute.
  The worked case is `upstream-issue-1402`, bullet indentation in LibreOffice
  and OnlyOffice.

  **A report arriving in browser clothes is not yet a layout report.** Ask what
  the browser actually hands the code path before filing one here.
  `gitbrent/PptxGenJS#1200`, `tableToSlides` auto-paging overflow, sat out of
  scope on the reasoning that its sizing input needed a browser. Then the
  headless repro got built. It reproduced, and the bug was arithmetic:
  the pager dropped one row's cell margins at every page break, through
  `addTable(rows, { autoPage })`, with no DOM in sight. The browser supplies
  column widths to that path and nothing else the vertical arithmetic reads.

  Revisiting the exclusion costs a headless-browser layout oracle plus an engine
  matrix, maintained indefinitely. It also needs a real consumer whose use case
  `data-pptx-width` and `data-pptx-min-width` cannot serve. Without that, the
  answer is no. The failure space is otherwise unbounded, sitting at the
  intersection of CSS layout, font fallback, and PowerPoint's own table layout.

  HTML `<table>` → slides left that category. `tableToSlides` ships as a free
  function on `pptx-ts/html`, runs under Node with any DOM implementation, and is
  covered end to end against happy-dom
  (`test/regression/html/html-to-slides-node.test.js`). The one thing it cannot do
  without a browser is *measure*, since `offsetWidth` reads `0` where nothing laid
  the table out. Column widths then fall back to the computed CSS widths, then to
  an equal split, and `data-pptx-width` and `data-pptx-min-width` pin them. That
  fallback is the whole scope boundary. Everything but real measurement works
  anywhere.

  Call it a fallback and not a degradation, because the two bases measure
  different boxes. `offsetWidth` is the border box. Computed `width` is the
  content box. Padding alone separates them, so one table can emit different
  column *proportions* on either side of a layout engine. A different answer, not
  the same answer with less precision. The fixture behind
  `test/browser/table-widths.spec.mjs` exists to show exactly that, 1:1 measured
  against 2:1 from CSS. Where both runtimes have to agree on a column, say so with
  `data-pptx-width`.

  The in-memory `addTable(rows, opts)` path remains the way to build a table
  from data you already hold; converting an existing HTML table is what the
  `/html` entry is for.

  *Contributor note:* the established pattern is to extract the DOM-independent
  decision into a pure helper and unit-test it with synthetic inputs. The
  originals are `resolveHtmlColWidth` / `htmlBorderToProps`
  (`test/regression/html/html-table-col-width.test.js`,
  `html-table-border-width.test.js`); the portability work added
  `pickColWidthBasis`, `parseCssWidthBasis`, `parseCssPx`, `cssColorToHex` and
  `readCellText` (`test/regression/html/html-table-portable-basis.test.js`), and the
  HTML-vs-pptx grid reconciliation added `measureGridColumns` / `extendColBasis`
  (`test/regression/html/html-table-grid.test.js`). Follow it. Those helpers are
  the reason the flow could be made portable at all. A full-fidelity *layout*
  repro still needs a headless browser, Playwright or Puppeteer, neither of which
  is a project dependency. A DOM-only repro no longer does.

- **Third-party office-suite interop quirks.** Some bugs only appear after a file
  round-trips through another application. Copy and paste inside WPS Office, then
  open the result in PowerPoint. Those are out of active scope when the generated
  package is valid OOXML and the other application introduced the corruption. The
  supported bar is that output opens cleanly in Microsoft PowerPoint. Keynote,
  LibreOffice Impress, and Google Slides import are best effort.

  *Contributor note:* a worked repro that pins the defect to invalid OOXML the
  library *itself* emits (independent of the other application) turns one of these
  into an in-scope correctness bug.

## Escape hatches

An escape hatch is any API that lets a caller step around the library's own
abstractions. This project ships several on purpose. One rule decides whether a
proposed one is acceptable:

> An escape hatch is fine when it bypasses a **convenience**, and needs a much
> higher bar when it bypasses a **guarantee**.

The guarantee here is *the bytes we author are valid OOXML that PowerPoint opens
cleanly*. Everything else is convenience. Unit conversion, autocomplete on a colour
string, a curated enum subset. A hatch through any of those costs the caller
nothing but their own care.

That one rule explains the read/write asymmetry already in the codebase.

- **Write path**: the library authors the bytes, so the guarantee applies. Only
  narrow, typed, validated hatches. Where a hatch takes uninterpreted input
  (`ShapeGuide.formula`, `shapeAdjust`, a `"<n>emu"` coordinate), it is guarded
  the way the rest of the write path is: warn and skip, warn and fall back, or
  throw, never silently emit a degenerate result (see the API Evolution Policy
  in `AGENTS.md`).
- **Read path**: the library never authored the bytes, so it offers no such
  guarantee. That makes one deep raw hatch acceptable. `part.dom` plus `element_`
  on the read model give direct DOM access at every level. The promise the read
  path *does* make, that untouched parts round-trip byte-identically, survives it,
  because reserialization only touches parts the caller marked dirty. The
  obligation is the caller's: call `markDirty()`. Every class exposing `element_`
  exposes it too. See `docs/reference/pptx-read.md`.

The worked rejection is the caller-provided XML transform hook. A generic
write-side hook over the emitted XML bypasses the guarantee itself, and leaves the
library unable to claim anything about its own output. Where a request like that
has a real need behind it, the answer is a typed, validated primitive, not a wider
hatch.

## Maintenance posture

The repository should be understandable to a maintainer or an agent starting
from a clean checkout:

- package support is documented in `docs/runtime-and-package-support.md`;
- development commands are documented in `docs/contributing/development.md`;
- verification commands are documented in `docs/contributing/testing.md`;
- OOXML source-of-truth lookup is documented in `docs/contributing/ooxml.md`.
