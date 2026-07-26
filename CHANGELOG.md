# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **`Presentation.appendSlides` now carries speaker notes.** A generator slide
  authored with `addNotes` previously lost its notes entirely when spliced onto a
  loaded deck — `extractSlides` emitted no notes part, so the append path had
  nothing to wire. It now emits a `notesSlide` per notes-bearing slide, wired back
  to the slide it annotates, with the notes body's own hyperlink relationships
  preserved (`rId1` = notes master, `rId2` = slide, hyperlinks from `rId3`).

  A notes slide must bind to a notes master, and a template usually has none — a
  deck authored without speaker notes carries no `notesMaster` part at all — so the
  generator's own notes master (and the theme its `.rels` requires) rides along in
  `ExtractedSlides.notesMaster` and is installed **only** when the destination deck
  has none. A destination that already has a notes master keeps it, so its notes
  styling wins; this matches the existing `importNotes` policy. `ExtractedSlide`
  gains an optional `notes` field for the same reason.

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
- **Four aggregate checks**, replacing the practice of hand-composing four or
  five scripts per iteration. `verify` (~45s) is the three typechecks +
  `backlog:validate` + the whole test suite; `verify:full` (~65s) adds
  `package:lint`, `test:package`, and `test:demos`. `check:static` and
  `check:package` are the two halves CI runs as separate jobs. `verify` and
  `verify:full` omit `lint`/`format:check` by design — the git hooks own those.
- **A `dist/` freshness guard** (`scripts/ensure-dist.mjs`) that every test,
  typecheck, and package script now starts with. It rebuilds only when `src/` or
  a build config is newer than `dist/`, and is a ~0.1s no-op otherwise. This
  replaces both halves of the old pattern — the unconditional
  `pnpm run build &&` prefix and the `:fast` twins that skipped it — so there is
  no longer a stale-`dist/` footgun to reason about.

### Changed

- **Breaking (internal constructors):** `ChartAxis`, `ChartSeries`, and
  `ChartExAxis` now take the owning chart part (respectively the owning `ChartEx`)
  as a second constructor argument, so `markDirty()` can reach it. These are
  obtained from `Chart.axes` / `Chart.series` / `ChartEx.axes`; only code that
  hand-constructed them is affected.
- **Breaking (structural type):** `ResolvedTableStyle` gained a required
  `markDirty(): void` member. Only code that builds the object literal itself
  (rather than reading it from `table.resolvedStyle`) is affected.
- **Breaking (internal type):** the internal `ShadowPropsInternal.opacity`
  field (the derived shadow alpha) is renamed to `_alpha`, clearing it of the
  removed public `opacity` shadow input. A stray `opacity` from an untyped/legacy
  caller is now inert (it lands on a field nothing reads) instead of being
  actively stripped — no behaviour change for supported inputs; use
  `transparency` (0–100). Only code reading the internal shape off a corrected
  shadow is affected.
- `pnpm run byte-identity:baseline` now refuses to run when `src/gen/` has
  uncommitted changes (override: `--allow-dirty`). A baseline frozen after the
  refactor has begun records the very bytes it exists to detect, so every later
  `check` passes trivially. The error names the workaround it is closing —
  `git stash` on a dirty tree, which risks unrelated work to a pop conflict.
- The three `tsc` projects are now `incremental`, keeping their build state in
  the gitignored `.tmp/` (a distinct `tsBuildInfoFile` per project). Warm
  `typecheck` drops from ~3.4s to ~1.3s; cold runs are no slower, so CI is
  unaffected.
- The OOXML schema fixtures now run concurrently, taking that suite from ~50s to
  ~10s — cheap enough that `verify` runs it on every iteration instead of
  reserving it for `verify:full`.
- The root build configs (`eslint.config.mjs`, `vitest.config.ts`,
  `tsdown.config.ts`, `tsdown.dev.config.ts`) are now linted and typechecked.
  They previously matched no ESLint `files` block and no tsconfig `include`, so
  nothing checked them at all.
- A missing OOXML validator now **fails** the read suite under `CI` instead of
  silently skipping a few hundred schema assertions. Locally it still skips, but
  prints a notice — a green local run no longer reads as a complete one.
- CI runs the static checks once rather than once per Node version, and adds a
  Windows leg for the package and demo scripts, which are the only exercise the
  Windows-specific subprocess handling in `scripts/script-utils.mjs` gets. The
  publish workflow now reuses the CI workflow instead of keeping its own copy of
  the gate, which had already drifted out of sync.

### Fixed

- **A theme-indexed picture background reported the wrong part.** When a
  `p:bgRef`'s `fmtScheme` entry is an `a:blipFill` — the third `bgFillStyleLst`
  slot in several stock Office themes (Ion, Facet, …) — the fill element comes out
  of the *theme* part, so its `r:embed` is scoped to the theme's relationships.
  `SlideBackground.resolvedFill` resolved it against the owning
  slide/layout/master's relationships instead, which does not fail loudly: the
  same id usually exists there and points at something else entirely. On an
  Ion-themed deck `resolvedFill.partName` read `/ppt/slideLayouts/slideLayout1.xml`
  where the image is `/ppt/media/image1.jpeg`. Affects `Slide.background`,
  `SlideMaster.background`, and `SlideLayout.background`.

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
