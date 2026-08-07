---
doc-schema-version: 1
title: "Project Target"
summary: "Current goals, non-goals, and maintenance posture for this ts-pptx project."
read_when:
  - Deciding whether a feature fits this project
  - Updating package target or compatibility policy
  - Explaining current project goals
doc_type: "decision"
---

# Project Target

ts-pptx generates PowerPoint `.pptx` packages from TypeScript and modern
JavaScript. The project target is a maintained, ESM-first library for
applications that need to create presentations programmatically.

This is an **independent derivative** of [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS),
detached at its v6.0.0 (June 2025). Upstream tracking is retired: the project does
not sync from or mirror the original and sets its own direction. It descends from
the original codebase and retains the original MIT copyright; it is not a
drop-in-compatible continuation of the upstream release line.

## Goals

- Generate `.pptx` packages without requiring PowerPoint at runtime.
- Keep the public package boundary explicit and easy to verify.
- Provide TypeScript declarations that work in modern app code.
- Support Node.js `>=24` and modern bundler-driven front-end applications.
- Support the **browser as a runtime**, and prove it rather than assert it: the
  browser build and its runtime adapter are exercised in CI against a real
  Chromium, and the deck a browser assembles is compared part-for-part against
  the Node-built one. This is a claim about *emission*, and it stops there —
  see the Live-DOM bullet under Out Of Active Scope for where the line is.
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

The project is **Node-first**: the generator runs and is tested without any
office application, and the Node suite is where nearly all of it is proven. That
is a statement about where the evidence lives, not a hedge about the browser —
the browser is a supported runtime with its own CI lane (see
[Runtime And Package Support](runtime-and-package-support.md#what-browser-is-tested-to-mean)).
Two domains fall outside what the maintainer actively develops, because there is
no in-house use case driving them. They are **not
rejected on merit** — they are simply not on the maintenance roadmap, and the
maintainer will generally not pick up bugs or feature requests in these areas.
**Pull requests that fix or extend them are welcome** (ideally with the testing
approach noted below), and the OOXML-emission core they build on is fully
supported.

- **Live-DOM / browser-layout features.** Anything whose *answer* comes from a
  rendered page: real `offsetWidth` after layout, the resolved cascade, fonts as
  the browser actually chose them. Reproducing those faithfully needs a real
  browser, so features that depend on them are out of active scope.

  This is a different claim from "the browser is a supported runtime", and the
  two are worth keeping apart, because a report lands in one bucket or the
  other. Running the library in a browser is supported and tested. Committing
  that the library's output *matches how a browser laid something out* is not,
  and there is no oracle for it — every other gate in this repo has one (schema
  validation, byte identity, a PowerPoint render); "renders differently in
  Firefox" has none, so correctness would become a judgement call per report.
  A `.pptx` a browser builds differently from Node is a defect; a layout
  difference between two browsers is not.

  Two triage rules, both learned rather than assumed. **A report is only a
  live-DOM report when the disagreement sits upstream of PowerPoint.** A
  construct PowerPoint renders as intended and another viewer does not belongs
  to the third-party bullet below — a browser-layout oracle would not answer it,
  because nothing about the rendered page is in dispute (`upstream-issue-1402`,
  bullet indentation in LibreOffice and OnlyOffice, is the worked case). And
  **a report arriving in browser clothes is not yet a layout report:** ask what
  the browser actually supplies to the code path before filing one here.
  `gitbrent/PptxGenJS#1200`, `tableToSlides` auto-paging overflow, sat out of
  scope on the reasoning that the sizing input driving it could not be exercised
  without a browser. The headless repro was eventually built, it reproduced, and
  the bug was arithmetic — the pager dropped one row's cell margins at every
  page break, reproducible through `addTable(rows, { autoPage })` with no DOM at
  all. The browser supplies column widths to that path and nothing the vertical
  arithmetic reads.

  Revisiting the exclusion has a stated entry cost — a headless-browser layout
  oracle plus an engine matrix, maintained indefinitely — and a prerequisite: a
  real consumer whose use case cannot be served by `data-pptx-width` /
  `data-pptx-min-width`. Absent that the answer is no. The failure space is
  otherwise unbounded, being the intersection of CSS layout, font fallback, and
  PowerPoint's own table layout.

  HTML `<table>` → slides is **not** in that category any more. `tableToSlides`
  is a supported, tested, portable path: it ships as a free function on
  `ts-pptx/html`, runs under Node with any DOM implementation, and is covered
  end-to-end against happy-dom (`test/regression/html-to-slides-node.test.js`).
  What it cannot do without a browser is *measure* — `offsetWidth` is `0` where
  nothing laid the table out — so column widths fall back to the computed CSS
  widths, then to an equal split, and `data-pptx-width` /
  `data-pptx-min-width` are there to pin them. That fallback is the scope
  boundary: everything except real measurement works anywhere.

  A *fallback*, deliberately, and not a *degradation* — the two bases do not
  measure the same box. `offsetWidth` is the **border box**; computed `width` is
  the **content box**. Padding alone is enough to separate them, so one table can
  emit different column *proportions* on either side of a layout engine — a
  different answer, not the same answer less precisely. The fixture behind
  `test/browser/table-widths.spec.mjs` is built to show it: 1:1 measured against
  2:1 from CSS. Where both runtimes have to agree on a column, state it with
  `data-pptx-width`.

  The in-memory `addTable(rows, opts)` path remains the way to build a table
  from data you already hold; converting an existing HTML table is what the
  `/html` entry is for.

  *Contributor note:* the established pattern is to extract the DOM-independent
  decision into a pure helper and unit-test it with synthetic inputs. The
  originals are `resolveHtmlColWidth` / `htmlBorderToProps`
  (`test/regression/html-table-col-width.test.js`,
  `html-table-border-width.test.js`); the portability work added
  `pickColWidthBasis`, `parseCssWidthBasis`, `parseCssPx`, `cssColorToHex` and
  `readCellText` (`test/regression/html-table-portable-basis.test.js`), and the
  HTML-vs-pptx grid reconciliation added `measureGridColumns` / `extendColBasis`
  (`test/regression/html-table-grid.test.js`). Follow
  it — those helpers are why the flow could be made portable at all. A
  full-fidelity *layout* repro still needs a headless browser
  (Playwright/Puppeteer), which is not a project dependency; a DOM-only repro no
  longer does.

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

## Escape Hatches

An escape hatch is any API that lets a caller step around the library's own
abstractions. This project ships several, deliberately, and the rule for whether
a proposed one is acceptable is:

> An escape hatch is fine when it bypasses a **convenience**, and needs a much
> higher bar when it bypasses a **guarantee**.

The guarantee here is *the bytes we author are valid OOXML that PowerPoint opens
cleanly*. Everything else — unit conversion, autocomplete on a colour string, a
curated enum subset — is convenience, and a hatch through it costs the caller
nothing but their own care.

That single rule produces the read/write asymmetry the codebase already has:

- **Write path** — the library authors the bytes, so the guarantee applies. Only
  narrow, typed, validated hatches. Where a hatch takes uninterpreted input
  (`ShapeGuide.formula`, `shapeAdjust`, a `"<n>emu"` coordinate), it is guarded
  the way the rest of the write path is: warn and skip, warn and fall back, or
  throw — never silently emit a degenerate result (see the API Evolution Policy
  in `AGENTS.md`).
- **Read path** — the library never authored the bytes, so no such guarantee is
  on offer. One deep raw hatch is therefore acceptable: `part.dom` plus
  `element_` on the read model gives direct DOM access at every level. The
  promise the read path *does* make — untouched parts round-trip
  byte-identically — survives it, because reserialization is scoped to parts the
  caller explicitly marked dirty. The obligation that comes with the hatch is
  that the caller must call `markDirty()`; every class exposing `element_` also
  exposes it. See `docs/reference/pptx-read.md`.

The worked precedent for a rejection is the caller-provided XML transform hook
(`docs/backlog.yml`, `upstream-issue-1282`, dismissed under
`escape-hatch-footgun`): a generic write-side hook over the emitted XML bypasses
the guarantee itself and leaves the library unable to make any claim about its
own output. A concrete need behind such a request is met with a typed, validated
primitive instead — not by widening the hatch.

## Maintenance Posture

The repository should be understandable to a maintainer or an agent starting
from a clean checkout:

- package support is documented in `docs/runtime-and-package-support.md`;
- development commands are documented in `docs/development.md`;
- verification commands are documented in `docs/testing.md`;
- OOXML source-of-truth lookup is documented in `docs/ooxml-agent-context.md`.
