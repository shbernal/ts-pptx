---
doc-schema-version: 1
title: "ts-pptx vs PptxGenJS"
summary: "What ts-pptx 3.7.0 and pptxgenjs 4.0.1 each emit, how many of their decks validate, what each costs to bundle and install, and how the two projects are run."
read_when:
  - Choosing between ts-pptx and pptxgenjs
  - Checking whether a construct is emitted by one library or by both
  - Weighing what ts-pptx gives up against what it adds
doc_type: "overview"
---

<!-- GENERATED FILE. Do not edit by hand.
     Regenerate with `pnpm run comparison:render`.
     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->

# ts-pptx vs PptxGenJS

ts-pptx 3.7.0 and pptxgenjs 4.0.1 were measured on 2026-09-15 by building the same 22 deck
intents with each library and reading the bytes that came out. ts-pptx descends from
pptxgenjs, detached at its v4.0.1 ([lineage](getting-started/introduction.md#lineage)), so
every difference below comes from running both rather than from either one describing
itself.

## Before you choose

- **Node.js 24 or later.** pptxgenjs declares no engine floor and runs on much older
  releases.
- **One ESM build.** `require("pptx-ts")` works through the ESM interop Node has had since
  22.12, and a browser loads the package as a module. pptxgenjs also ships CommonJS and a
  classic-script bundle that defines a global. See [where it
  runs](getting-started/runtime.md).
- **Not a drop-in continuation of the upstream release line.** The API is close by
  descent, not by contract, and it has moved since. Moving code across is a port, not an
  upgrade: [porting from PptxGenJS](comparison-syntax.md) lists the calls that change.
- **No SmartArt on the write side, in either library.** It is not a difference between
  them, but it is a real gap in both.
- **Adoption is not close.** pptxgenjs was downloaded 10,900,438 times in the last month,
  against 1,680 for ts-pptx. That gap buys answers that already exist, examples written by
  people other than the maintainer, and good odds that a bug on a common path was hit by
  someone else first. If that outweighs the differences below, use pptxgenjs.

## Scorecard

|  | ts-pptx | pptxgenjs |
|---|---|---|
| Intents emitted | 21 of 22 | 10 of 22 |
| Decks with no schema error | 21 of 21 | 0 of 10 |
| Hello world, bundled and gzipped | 103.3 KiB | 123.3 KiB |
| Runtime dependencies, transitive | 3 | 18 |
| Installed size, with dependencies | 10.7 MiB | 6.6 MiB |

The bundled size is what a browser program fetches before its first line runs. [How the
comparison was measured](comparison-method.md#package-hygiene) has every install and
bundle figure, and the programs behind them.

## Construct coverage

<CoverageMatrix />

Of 22 intents, ts-pptx emits 21 and pptxgenjs emits 10.

Emitted by both: Text run, Table, Raster image, Bar chart, External hyperlink,
User-defined slide master, Sections, Speaker notes, Preset-geometry shape and Slide
background colour.

Emitted by ts-pptx only: Slide transition, Build animation on a shape, Embedded OLE
object, 3D model, Embedded font face, Connector between shapes, Inline equation, Gradient
shape fill, 3D bevel on a table cell, Funnel chart (chartEx) and Slide Zoom tile.

No intent is emitted by pptxgenjs and not by ts-pptx.

Emitted by neither: SmartArt diagram (write side).

Every intent a library does not emit is one it has no API for, rather than one it tried
and got wrong.

[How the comparison was measured](comparison-method.md#construct-coverage) has the token
each intent is read for, the part it is read from, and the notes on individual rows.

## Schema validity

<ValidityBars />

ts-pptx built 21 of the 22 decks, and all 21 validated cleanly. It had no API, and so no
deck, for the remaining one.

pptxgenjs built 10 of the 22 decks, and none validated cleanly (12 errors in all). It had
no API, and so no deck, for the remaining 12.

Every deck went through the Open XML SDK validator (3.5.1) at the `Microsoft365`
conformance target. The decks a library could not build are counted because a library with
fewer decks has fewer decks to be wrong in. [How the comparison was
measured](comparison-method.md#schema-validity) lists each distinct error.

## Generation time

<TimingRatio />

Each cell is the ts-pptx median divided by the pptxgenjs median for the same deck, so a
figure below 1× means ts-pptx took less time.

| Deck | Compressed | Stored |
|---|---|---|
| Hello world | 0.49× | 1.33× |
| Text deck | 0.65× | 1.26× |
| Table deck | 0.65× | 1.06× |
| Chart deck | 0.69× | 0.95× |
| Full deck | 0.79× | 1.31× |
| 50 slides | 0.80× | 1.21× |
| 200 slides | 0.75× | 1.21× |
| 500 slides | 0.78× | 1.23× |

The milliseconds, the machine they were taken on, and why the two settings point in
opposite directions are in [how the comparison was
measured](comparison-method.md#generation-time).

## Reading decks

pptxgenjs generates decks and does not read them. ts-pptx also reads:

- [Inspection](reference/pptx-inspection.md) reports what a package contains without
  parsing it into a model.
- [Reading](reading/read-and-edit.md) loads a deck into an object model, edits it in
  place, and writes the package back out.
- [Deck to script](reference/pptx-to-script.md) turns a deck into the TypeScript that
  rebuilds it, reporting what it could not express rather than dropping it.

## Adoption and project health

How the two projects are run, kept apart from everything above because stars and downloads
measure history as much as merit.

|  | ts-pptx | pptxgenjs |
|---|---|---|
| Repository | [shbernal/ts-pptx](https://github.com/shbernal/ts-pptx) | [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS) |
| Default branch | `master` | `master` |
| Last commit on the default branch | 2026-09-10 | 2025-06-26 |
| Last npm publish | 2026-08-29 | 2025-06-26 |
| Downloads, last month | 1,680 | 10,900,438 |
| Stars | 2 | 6,165 |
| Open issues | 0 | 232 |
| Open pull requests | 0 | 65 |
| Source lines | 67,678 | 10,125 |
| Test lines | 76,333 | 0 |
| Test suite | 13 test scripts, 363 spec files under `test/` | no test script, no spec file, no test directory |
| Statement coverage | 95.55% (node lane only) | no automated suite |

[How the comparison was measured](comparison-method.md#project-health) says how each of
these figures is taken.

## More

- [How the comparison was measured](comparison-method.md): the corpus, every full table,
  and how each figure was taken.
- [Porting from PptxGenJS](comparison-syntax.md): the calls that change between the
  libraries, and the code behind every row.
