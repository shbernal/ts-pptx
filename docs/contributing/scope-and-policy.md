---
doc-schema-version: 1
title: "Scope and design policy"
summary: "What ts-pptx leaves out on purpose, what stays in a consumer, the two areas outside active development, and the rule for escape hatches."
read_when:
  - Deciding whether a feature or a report fits this project
  - Triaging a report about browser layout or another office suite
  - Adding, widening or removing an escape hatch
doc_type: "decision"
---

# Scope and design policy

This page records what the project leaves out and the rules for deciding. What a user needs to
know first, what the library does, where it runs and where it came from, is on the site's
[Introduction](../getting-started/introduction.md).

## Non-goals

- Shipping more than one build. The published ESM artifact is what Node, bundlers, browsers and
  `require()` callers all load; see
  [Runtime and package support](../runtime-and-package-support.md#one-build-and-everything-that-loads-it).
- Reproducing the upstream release matrix: its historical artifact names, its `window.TsPptx`
  global, or compatibility with its build system. A browser reaches this package as a module,
  from a bundler or an ESM CDN.
- Treating generated `dist/` outputs as hand-edited source.

### What stays in the consumer

The bullets above are about the *shape* of the package. These are different. Each one encodes a
specific consumer's brand, content, or deck workflow. None is a candidate here at any priority, and
the line holds even when the code looks perfectly generic. Don't raise them.

- Brand guidance, workflow-specific scripts, and consumer content.
- A consumer's default font choice.
- Icon-set policy, imports, aliases, and provenance manifests.
- External stock-asset sourcing helpers.
- Lint quality thresholds, annotated screenshots, and human-review artifacts.
- Slide semantics manifests as agent-facing design-intent contracts.
- Greenfield deck eval prompts, scorecards, and generator-adapter behavior.
- LibreOffice/ImageMagick rendering orchestration for local visual QA.

What this package owes such a consumer is sound building blocks. The composition on top of them
belongs to the consumer. A generic PPTX gap uncovered while building one *is* in scope, though. See
[Agent development guide → Promoting a downstream need](agent-development.md#promoting-a-downstream-need)
for how it moves across.

## Out of active scope (contributions welcome)

The project is Node-first. The generator runs and is tested with no office application near it,
and the Node suite carries nearly all of its evidence. The browser is still a supported runtime
with its own CI lane; see
[what "browser" is tested to mean](../runtime-and-package-support.md#what-browser-is-tested-to-mean).

Two areas sit outside what the maintainer actively develops, because no in-house use case drives
them. Neither is rejected on merit. Reports there tend to wait, pull requests are welcome, and the
OOXML emission underneath both is fully supported.

| Area | Supported | Not actively developed | Why there is no oracle | What makes a report in scope |
| --- | --- | --- | --- | --- |
| Browser layout | Running the library in a browser, and converting an HTML `<table>` with `tableToSlides` under any DOM | Output that has to match how a browser laid a page out: `offsetWidth` after layout, the resolved cascade, the fonts a browser picked | Every gate in this repository has one: schema validation, byte identity, a PowerPoint render. "Renders differently in Firefox" has none, so each report would become a judgement call | A `.pptx` a browser builds differently from Node, or a defect that reproduces with no layout engine at all |
| Other office suites | Output that opens cleanly in desktop PowerPoint. Keynote, LibreOffice Impress and Google Slides import on a best-effort basis | Breakage that appears only after another application round-trips a file, such as a copy and paste in WPS Office followed by opening the result in PowerPoint | The package this library wrote is valid OOXML, so no check on it can see what the other application did to it | A repro that pins the defect to invalid OOXML the library itself emits |

Two rules for triage:

- **A report is only a browser-layout report when the disagreement sits upstream of PowerPoint.**
  A construct PowerPoint renders as intended and another viewer does not belongs to the office
  suites row, since nothing about a rendered page is in dispute.
- **A report that arrives from a browser is not yet a layout report.** Ask what the browser
  actually hands the code path. The browser supplies `tableToSlides` its column widths and nothing
  else, so a table defect that reproduces headless through `addTable` is an ordinary bug.

A fix in the browser-layout area follows one pattern: extract the decision that does not need a
DOM into a pure helper, and unit-test that helper with synthetic inputs.

## Escape hatches

An escape hatch is any API that lets a caller step around the library's own abstractions. The
project ships several on purpose, and one rule decides whether a proposed one is acceptable:

> An escape hatch is fine when it bypasses a **convenience**, and needs a much higher bar when it
> bypasses a **guarantee**.

The guarantee is that *the bytes the library authors are valid OOXML that PowerPoint opens
cleanly*. Everything else is convenience: unit conversion, autocomplete on a colour string, a
curated enum subset. A hatch through a convenience costs the caller nothing but their own care.

That rule is why the write path and the read path get different hatches:

| | Bypasses a convenience | Bypasses a guarantee |
| --- | --- | --- |
| **Write path**: the library authors the bytes | Accepted when narrow, typed and validated, as `ShapeGuide.formula`, `shapeAdjust` and a `"<n>emu"` coordinate are. Bad input warns and skips, warns and falls back, or throws, never emits a degenerate result (the API evolution policy in `AGENTS.md`) | Rejected. A caller-provided transform over the emitted XML would leave the library unable to claim anything about its own output |
| **Read path**: the library did not author the bytes | Accepted, and deep: `part.dom` and `element_` give direct DOM access at every level of the read model | Nothing to bypass: there is no validity promise about bytes the library did not write. The promise it does make survives the hatch. Untouched parts round-trip byte for byte, because only parts marked with `markDirty()` are reserialized, and marking them is the caller's job. See [PPTX read and round-trip](../reference/pptx-read.md) |

When a request for a wider write-side hatch has a real need behind it, the answer is a typed,
validated feature, not a wider hatch.

## Maintenance posture

The repository should make sense to a maintainer or an agent starting from a clean checkout. OOXML
changes are testable through regression tests, schema fixtures and package-level smoke tests, and
the evidence, commands and research paths are written down:

- package support: `docs/runtime-and-package-support.md`;
- development commands: `docs/contributing/development.md`;
- verification commands: `docs/contributing/testing.md`;
- OOXML lookup and what counts as evidence: `docs/contributing/ooxml.md`.
