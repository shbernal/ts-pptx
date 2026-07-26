# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **`markDirty()` on every read-model class that exposes `element_`.** `element_`
  hands out the live DOM node, but an edit through it was silently discarded on
  `save()` unless the caller reached the owning part themselves
  (`shape.slide.part.markDirty()` — three hops, undocumented). The obligation now
  sits on the same object as the hatch: `Slide`, `Shape`, `TextFrame`,
  `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `Placeholder`,
  `NotesPlaceholder`, `Theme`, `Chart`, `ChartAxis`, `ChartSeries`, `ChartEx`,
  `ChartExAxis`, `ChartExSeries`, and `ResolvedTableStyle`. `Shape.markDirty()`
  was `protected` and is now public.
- **`Slide.element_`** — the `p:sld` root, filling the one missing rung in the
  `element_` ladder. Slide-level DOM access was previously only reachable as
  `slide.part.dom`.
- Guide formulas passed to a `custGeom` shape's `guides` option now have their
  leading operation checked against the 17 operations ECMA-376 §20.1.9.11
  defines. An unknown operation (e.g. `{ name: 'w2', formula: 'bogus 1 2' }`)
  previously emitted schema-shaped but semantically dead geometry whose first
  feedback was a PowerPoint repair prompt; it now warns and skips the guide,
  matching the existing degenerate-entry behaviour. Operands are still passed
  through uninterpreted.
- The project's escape-hatch policy is now written down in
  [project target](docs/project-target.md) — the convenience-vs-guarantee rule
  and why the read path gets a deep raw hatch while the write path does not.

### Changed

- **Breaking (internal constructors):** `ChartAxis`, `ChartSeries`, and
  `ChartExAxis` now take the owning chart part (respectively the owning `ChartEx`)
  as a second constructor argument, so `markDirty()` can reach it. These are
  obtained from `Chart.axes` / `Chart.series` / `ChartEx.axes`; only code that
  hand-constructed them is affected.
- **Breaking (structural type):** `ResolvedTableStyle` gained a required
  `markDirty(): void` member. Only code that builds the object literal itself
  (rather than reading it from `table.resolvedStyle`) is affected.

## [1.0.0] - 2026-07-24

Initial public release of ts-pptx — an ESM-first, TypeScript-first library for
generating PowerPoint `.pptx` files from Node.js and modern JavaScript
toolchains.

ts-pptx descends from [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS)
(MIT) and has been developed independently since; see the README for lineage and
the [project target](docs/project-target.md) for scope. It ships its own API and
makes no backwards-compatibility guarantee with the original project.

### Added

- Slide authoring: slides, layouts, masters, sections, speaker notes, and
  presentation metadata.
- Content: text, tables, shapes, connectors, groups, images, SVGs, charts,
  media, and OLE objects.
- Outputs: file, stream, buffer, Blob, base64, and browser download, depending
  on the runtime.
- A `.pptx` read model for opening, inspecting, and round-tripping existing
  decks (`@shbernal/ts-pptx/read`).
- Standalone text measurement and table-fit helpers
  (`@shbernal/ts-pptx/measure`).
- Native equation authoring from LaTeX or MathML (`@shbernal/ts-pptx/math`).
- Explicit ESM package boundary with typed subpath exports for `inspect`,
  `measure`, `read`, `math`, `node`, and `browser`.

[1.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v1.0.0
