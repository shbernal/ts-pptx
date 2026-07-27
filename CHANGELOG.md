# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **A deck IR (`@shbernal/ts-pptx/script`), the read half of turning an existing
  `.pptx` back into source.** `readModelToIr(presentation)` walks a deck read
  through `ts-pptx/read` and returns a serializable description of the write-API
  calls that would rebuild it — `{ slideSize, props, slides, assets, fidelity }`,
  where each slide holds `{ method, args }` calls whose `args` are literal
  write-API option objects. Geometry is carried as exact `"<n>emu"` strings
  wherever the option is `Coord`-typed, and as six-decimal inches for the three
  that are not (`colW`, `rowH`, `margin`) — the proven minimum for an EMU-exact
  round-trip.

  It is a new subsystem rather than part of `ts-pptx/read` because it needs both
  the read model and the write option types, and because the read subpath is
  documented as isomorphic (bytes in, bytes out); a converter whose output is
  source text would break that guarantee for its consumers.

  Every construct that cannot survive is a `FidelityNote` on the IR rather than a
  log line, carrying `{ slideNumber, shapeName, construct, disposition, cause,
  detail }`. `cause` distinguishes a missing read accessor (`unread`) from a
  missing write option (`unwritable`) from a structural limit (`unsupported`),
  which is what makes a note actionable. Because the notes are data, a round-trip
  check can exclude exactly the declared losses and treat any other difference as
  a defect — an undeclared loss fails, and a declared loss that actually survives
  is a stale note. Read `DeckIr.fidelity` before trusting a conversion: notable
  entries include theme-referenced outline width (`p:style/a:lnRef` resolves a
  colour but not a width or dash), embedded audio/video (only the poster frame is
  readable), and OMML equations.

- **`printScript(ir, options)` turns that IR into a runnable TypeScript module.**
  Returns `{ code, assets, notes }`: the module source, the image bytes it
  expects beside it, and the losses that apply to *this* output. The emitted
  script anchors on a template, and the template is the **source deck itself,
  unmodified** — `Presentation.fromTemplate` already strips a package's slides
  while leaving its masters, layouts, theme, and document properties
  byte-identical, so only slide content is regenerated and the deck's whole
  design survives untouched. Slides are emitted in source order; contiguous
  slides sharing a layout share one generator, since `appendSlides` binds one
  layout per call.

  `notes` is not simply `ir.fidelity`. A template-anchored output *rescues* some
  declared losses — all twelve document properties ride in the template, so the
  IR's `deck.docProps` note does not apply and is omitted rather than left to
  teach readers to skim. It also *adds* losses that belong to the tier rather
  than to the conversion: a slide's `p:cSld@name` reads fine and would survive a
  byte copy, but has no public write-API setter. The applicable set is
  reproduced as a header comment in `code`, so the artifact carries its own
  caveats, and it is the set a round-trip check should exclude from its diff.

  Binding is by layout name where that is unambiguous, because a name survives
  being re-pointed at a different template; a deck whose masters repeat a layout
  name falls back to gallery position, since `appendSlides` throws on an
  ambiguous name rather than choosing.

  Not yet implemented: a standalone (template-free) tier that rebuilds the
  chrome via `defineSlideMaster`.

- **A round-trip check for generated scripts: `canonicalDeckIr` + `diffDeckIr`,
  and `pnpm run script:roundtrip`.** Reads a deck, prints a script, runs it,
  reads the deck that came out, and diffs the two IRs using the printer's
  fidelity notes as the exclusion list — so a difference no note predicted is a
  defect. `diffDeckIr` returns `{ differences, undeclared, declared, added,
  unmatchedNotes }`; `undeclared` is the number to gate on.

  The comparison is on IRs rather than packages because the output can never be
  byte-identical (fresh rel ids, regenerated shape ids), so comparing bytes would
  fail for every deck and measure nothing. `canonicalDeckIr` removes the noise
  that is not loss — a value spelled out that means what its absence means
  (`bold: false`, `wrap: true`, the default `a:bodyPr` insets), and asset
  identity by content digest rather than by generated filename — and each such
  rule cites the OOXML default that makes it an equivalence rather than a
  convenience.

  Its reach is bounded, and knowing how bounded is the point. Both IRs come from
  the same reader, so a construct the read path cannot see is absent from both
  and compares equal; and the converter need not be injective, so two source
  constructs mapping to one call also compare equal. It detects **asymmetry**.
  Mutation testing says so concretely: of twelve deliberate converter defects it
  catches six, and the six that survive are exactly the symmetric ones. Read a
  clean run as "nothing the converter can distinguish was lost", and pair it with
  `pnpm run read:census` and the IR unit tests, whose expectations come from
  `src/types/*.ts` rather than from the converter.

### Fixed

- **Nine converter defects that no static check and no execution check could
  catch**, all found by the round-trip comparison above and all producing a
  script that typechecked, ran, and wrote a plausible deck.
  - **Every paragraph bullet was wrong.** `Paragraph.bullet` reports a *tagged*
    string (`'none'`, `'char:<glyph>'`, `'autoNum:<type>'`) and the mapper read
    it as a bare glyph — so an explicit `a:buNone` became a literal `n` bullet
    (`'none'.codePointAt(0)`), a real character bullet became `c`, and a numbered
    list became `a`. The numbering scheme was also passed as `style` rather than
    `numberType`, so every numbered list fell back to the default scheme, and the
    glyph was emitted with fewer than the four hex digits the write path requires,
    which made it substitute its own default.
  - **Placeholders were emitted with no geometry at all.** A shape with no
    transform of its own reported no frame, and omitting `x`/`y`/`w`/`h` does not
    leave the geometry to be inherited — an appended slide inherits nothing — it
    produces a zero-height box in the corner. Now resolved through the
    layout/master chain via `resolvedFrame`.
  - **A group's rotation and flips were applied twice**, because the children were
    already emitted in composed slide-absolute coordinates *and* the transform was
    repeated on the group. A 30° group came back at 60°; a flipped group came back
    unflipped, the double flip having cancelled.
  - **Image crops were sent to `sizing`**, which crops in displayed inches against
    the image's measured natural size, instead of to `crop`, which is `a:srcRect`
    emitted verbatim. Fractions read as inches shrank every cropped picture.
  - **Every text body was re-anchored to centre.** An unset `a:bodyPr/@anchor`
    means top in OOXML and centre in `addText`, so the converter now spells the
    anchor out rather than leaving it off.
  - **Every uncoloured run was repainted black**, since `addText` fills a run with
    no colour using `DEF_FONT_COLOR`. The inherited colour is now resolved and
    emitted (with a note that it is frozen against later theme edits), or, where
    nothing resolves it, declared as a colour that may be wrong.
  - **PowerPoint text boxes were demoted to auto shapes.** Text-box-ness is
    `p:cNvSpPr/@txBox`, not "has no preset geometry" — PowerPoint gives every text
    box an explicit `prstGeom rect`, so the old test misclassified all of them.
  - **An SVG picture was reduced to its raster fallback.** The vector part is now
    preferred, since `addImage` accepts SVG bytes and regenerates a fallback.
  - **A bulleted paragraph with more than one run was split in two**, because the
    write path treats a bullet on a run that does not open a line as a request for
    a new paragraph. Paragraph-level `bullet` now rides on the first run only.
  - Also declared rather than fixed, each being a real limit rather than a
    mistake: an automatic field (`a:fld` — slide number, date, footer) has no
    accessor and no `addText` expression; an explicit zero `a:spcBef`/`a:spcAft`
    is indistinguishable from unset on the write side; an auto-height table row
    comes back pinned to its content's height; and a paragraph that inherits its
    bullet cannot say "inherit" through the write API.

- **`line: { type: 'none' }` did nothing.** The value is documented on
  `ShapeFillProps.type` and accepted by the type checker, but `genXmlLineFill` had
  no branch for it and emitted no paint child — which is how a stroke says
  "inherit", not "none". A shape authored with an explicitly suppressed outline
  therefore *grew* the theme's border instead of losing it. It now emits
  `<a:noFill/>`.

- **`readModelToIr` mapped three constructs onto write-API shapes that do not
  exist.** All three produced an IR that typechecked and a script that failed at
  run time, because `IrValue` is deliberately loose enough that `tsc` cannot
  check an argument against the signature it is meant to satisfy.
  - `addChart` received the chart type as a third positional argument; the
    signature is `addChart(data, options & { type })`, so the type now rides in
    the options object.
  - `addConnector` received a bounding box (`x`/`y`/`w`/`h`) and a nested `line`
    object. It takes two endpoints and flat stroke options, so connectors now
    emit `x1`/`y1`/`x2`/`y2` derived from the box **and its `a:flipH`/`a:flipV`
    flags** — without the flips every up- or leftward connector is silently
    mirrored — plus `color`/`width`/`dashType`/arrowheads. A gradient or
    translucent connector stroke has no flat spelling and is now noted.
  - A slide bound to a layout whose name is shared by another layout in the deck
    made `appendSlides` throw. `SlideIr.layout` now carries the gallery index and
    whether the name is unique.

- **Three shapes could vanish from a conversion with no fidelity note**, against
  the contract that a dropped shape is never silent: an auto shape with neither
  text nor geometry of its own (an unfilled placeholder, whose outline comes from
  the layout), and a group whose every child was dropped. Both are now declared
  (`shape.empty`, `group.empty`). The third path — an unrecognised shape kind —
  turned out to be unreachable, which the type checker proves.

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
