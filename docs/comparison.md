---
doc-schema-version: 1
title: "Comparison With PptxGenJS"
summary: "What ts-pptx 3.7.0 and pptxgenjs 4.0.1 each emit, what validates, what each costs to install, and how the two projects are run."
read_when:
  - Choosing between ts-pptx and pptxgenjs
  - Checking whether a construct is emitted by one library or by both
  - Weighing what ts-pptx gives up against what it adds
doc_type: "reference"
---

<!-- GENERATED FILE. Do not edit by hand.
     Regenerate with `pnpm run comparison:render`.
     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->

# Comparison With PptxGenJS

ts-pptx is an independent derivative of
[gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS), detached at its v4.0.1 (see
[project target](project-target.md)). Descending from a project is a poor reason to be
trusted over it, so every difference below was produced by running both libraries and
reading what came out.

Measured on 2026-09-04: ts-pptx 3.7.0 built from this repository, against pptxgenjs 4.0.1
installed from npm (published 2025-06-26).

## What this measures, and how

The corpus is 22 deck intents. Each one states an intent ("a slide that enters with a push
transition"), and each library expresses that intent in its own idiom. Transcribing one
library's calls into the other is how a comparison gets rigged, so the two arms of a probe
deliberately do not have to look alike. Both decks are then opened, and the part the probe
names is read for the token it names.

Four outcomes are possible, per probe per library:

| Outcome | Meaning |
|---|---|
| emitted | the token is present in the named part |
| absent | an API exists, and the output does not carry the token |
| no API | nothing in the public surface expresses the intent |
| error | the build threw |

`no API` is the only one of the four that is a claim rather than a reading, so it is
checked rather than trusted: that library's shipped bundle is searched for the token, and
a hit fails the measurement run unless the corpus carries a written reason for it. Those
reasons are printed under the table they belong to.

Two things a reader should price in. The corpus is ours, so it was chosen by an interested
party. It is kept honest in two specific ways: it carries a probe neither library can
satisfy, and the set of probes upstream emits and ts-pptx does not is reported below even
when it is empty, so an empty set is a stated result rather than something a reader has to
infer from a gap. A pull request that adds a probe is welcome, including one ts-pptx
fails.

Every number on this page comes from `scripts/comparison/snapshot.json`, which is
refreshed on release cadence and carries the date above. Nothing here is edited by hand.

## What ts-pptx gives up

- **No CommonJS build.** pptxgenjs ships one, so it runs unchanged on Node versions and
  toolchains ts-pptx cannot serve at all. `require('pptx-ts')` does work, through the ESM
  interop Node has had since 22.12, which every Node ts-pptx supports has. See [runtime
  and package support](runtime-and-package-support.md).
- **No global bundle, and no CDN script tag.** pptxgenjs can be dropped into a page with a
  `<script>` tag and used from a global. ts-pptx requires a bundler or a runtime that
  loads ES modules.
- **Node.js `>=24` only.** pptxgenjs declares no engine floor and runs much further back.
- **Not a drop-in continuation of the upstream release line.** The API is close by
  descent, not by contract, and it has moved since. Migrating is a port, not an upgrade.
- **No SmartArt on the write side.** Neither library generates it, so this is not a
  difference between them, but it is a real gap in both.
- **Adoption is not close.** pptxgenjs was downloaded 11,116,327 times in the last month,
  against 2,019 for ts-pptx. That gap buys real things: answers that already exist,
  examples written by people who are not the maintainer, and reasonable odds that a bug on
  a common path was hit by someone else first. Anyone who weighs those above the
  differences measured below should use pptxgenjs.

## Construct coverage

Of 22 probes, ts-pptx emitted 21 and pptxgenjs emitted 10.

The middle column is the token the harness looks for. It is the OOXML element in every
case but one, where the intent is speaker notes and the token is the note text itself; the
part each token has to appear in is recorded in the snapshot.

### Shared baseline

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Text run | `<a:t>` | emitted | emitted |
| Table | `<a:tbl>` | emitted | emitted |
| Raster image | `<p:pic>` | emitted | emitted |
| Bar chart | `<c:barChart>` | emitted | emitted |
| External hyperlink | `<a:hlinkClick` | emitted | emitted |
| User-defined slide master | `<p:ph` | emitted | emitted |
| Sections | `<p14:sectionLst` | emitted | emitted |
| Speaker notes | `probe note` | emitted | emitted |
| Preset-geometry shape | `<a:prstGeom` | emitted | emitted |
| Slide background colour | `<p:bg>` | emitted | emitted |

### Motion

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Slide transition | `<p:transition` | emitted | no API |
| Build animation on a shape | `<p:timing>` | emitted | no API |

### Embedding

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Embedded OLE object | `<p:oleObj` | emitted | no API |
| 3D model | `am3d:model3d` | emitted | no API |
| Embedded font face | `<p:embeddedFontLst>` | emitted | no API |

### Shapes

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Connector between shapes | `<p:cxnSp>` | emitted | no API |

### Text

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Inline equation | `<a14:m` | emitted | no API |

### Fills

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Gradient shape fill | `<a:gradFill` | emitted | no API |

- Gradient shape fill, pptxgenjs: appears only inside the bundled Office theme XML, which
  no API parameterises.

### Tables

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| 3D bevel on a table cell | `<a:cell3D` | emitted | no API |

### Charts

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Funnel chart (chartEx) | `<cx:chart>` | emitted | no API |

### Navigation

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| Slide Zoom tile | `pslz:sldZm` | emitted | no API |

### Diagrams

| Intent | Looked for | ts-pptx | pptxgenjs |
|---|---|---|---|
| SmartArt diagram (write side) | `<dgm:relIds` | no API | no API |

The shared baseline is the control group. A corpus holding only constructs one side cannot
produce would prove that the corpus was chosen, not that the libraries differ, so 10 of
the probes are ones both libraries are expected to pass. A failure there fails the
measurement run instead of becoming a row on this page.

No probe in this corpus is emitted by pptxgenjs and not by ts-pptx.

Emitted by neither library: SmartArt diagram (write side).

## Schema validity

**This validates the decks this corpus builds, not either library in general.** A deck no
probe builds is not covered by any of it, and a library can be perfectly correct on
everything these probes never touch.

Every deck the corpus built was passed through the Open XML SDK validator (3.5.1) at the
`Microsoft365` conformance target: the same oracle, and the same target, that this
project's own `test:schema` suite uses.

|  | ts-pptx | pptxgenjs |
|---|---|---|
| Decks validated | 21 | 10 |
| Decks with no error | 21 | 0 |
| Errors | 0 | 12 |
| Intents with no deck to validate | 1 | 12 |

The last row is the denominator a validity count needs. A library that builds fewer decks
has fewer decks to be wrong in, and reading the error counts without it would reward not
having an API.

There is no warning column. This validator reports a single severity, so a zero in a
second column would be a number nobody measured.

### What failed in the pptxgenjs decks

Distinct diagnostics rather than a raw error total. One fault repeated across every deck,
and that many unrelated faults, are different facts about a library, and a total on its
own cannot tell them apart.

| Diagnostic | Part | Decks |
|---|---|---|
| `Sch_UnexpectedElementContentExpectingComplex` | `/ppt/presentation.xml` | 10 |
| `Sch_UnexpectedElementContentExpectingComplex` | `/ppt/charts/chart1.xml` | 1 |
| `Sch_AttributeValueDataTypeDetailed` | `/ppt/presentation.xml` | 1 |

- `/ppt/presentation.xml`: `The element has unexpected child element
  'http://schemas.openxmlformats.org/presentationml/2006/main:notesMasterIdLst'. List of
  possible elements expected:
  <http://schemas.openxmlformats.org/presentationml/2006/main:notesSz>.`
- `/ppt/charts/chart1.xml`: `The element has unexpected child element
  'http://schemas.openxmlformats.org/drawingml/2006/chart:axId'.`
- `/ppt/presentation.xml`: `The attribute 'id' has invalid value
  '{c20c63a9-c643-1013-9457-821c050dc52c}'. The Pattern constraint failed. The expected
  pattern is \{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}.`

## Package hygiene

What a consumer gets. Each library was installed on its own into an empty directory,
upstream from the registry and ts-pptx from a pack of this working tree, so nothing here
is measured against a development checkout with its dependencies hoisted flat.

|  | ts-pptx | pptxgenjs |
|---|---|---|
| Installed size, with dependencies | 10.2 MB | 6.7 MB |
| Installed size, the package alone | 5.5 MB | 2.5 MB |
| Runtime dependencies, transitive | 3 | 18 |
| Runtime dependencies, direct | `@xmldom/xmldom`, `fflate`, `opentype.js` | `@types/node`, `https`, `image-size`, `jszip` |
| Entry points | `.`, `./inspect`, `./measure`, `./read`, `./script`, `./math`, `./zip`, `./html`, `./node`, `./browser` | `.` |
| Module formats | esm | cjs, esm |
| `engines.node` | `>=24` | not declared |
| Hello world, first chunk | 144 kB | 123 kB |
| Hello world, every chunk | 210 kB | 123 kB |

The hello world program is identical in intent on both sides and written in each library's
own idiom: one slide, one text box, then export. It is bundled with esbuild for the
browser, minified, and gzipped at level 9, following the conventions
`scripts/bundle-size-ratchet.mjs` documents, with one difference that matters. The ratchet
never bundles, so it cannot drop unreachable code and its figures are an upper bound on
what the package ships; this bundles and does tree-shake, because a consumer's build is
precisely the thing being compared here. **The two sets of numbers will not agree, and
neither is wrong.**

Two figures rather than one, because code splitting is on. The first chunk is what the
program pays to start; every chunk is what it can reach. They differ for ts-pptx because
font metrics load `opentype.js` behind a dynamic import that only runs once a font is
registered, and a bundler that can defer that will. Charging the program for a chunk it
may never fetch, and hiding bytes it might, are both misleading, so both are printed.

ts-pptx installs larger than pptxgenjs despite carrying fewer dependencies. Its `dist/`
ships unminified, and a large share of that weight is documentation comments that no
consumer build keeps, which is why the bundled figures above are much closer together than
the installed ones.

## The read side

pptxgenjs generates decks. It does not read them, and it does not claim to. So there is
nothing to compare here and no table: this is a capability one library has, which is a
different statement from one library being better at something both do.

ts-pptx also reads:

- [Inspection](reference/pptx-inspection.md) reports what a package contains without
  parsing it into a model.
- [Reading](reference/pptx-read.md) loads a deck into an addressable object model, edits
  it in place, and writes the package back out.
- [Deck to script](reference/pptx-to-script.md) turns an existing deck into runnable
  TypeScript, reporting what it could not express rather than dropping it silently.

If you only generate decks, none of this is a reason to choose either library.

## Project health

Separate from everything above, and on purpose. These figures describe how the two
projects are run, not what either one emits. Stars and downloads measure adoption,
adoption measures history as much as merit, and none of it belongs in the same table as a
construct a library does or does not write.

|  | ts-pptx | pptxgenjs |
|---|---|---|
| Repository | [shbernal/ts-pptx](https://github.com/shbernal/ts-pptx) | [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS) |
| Default branch | `master` | `master` |
| Last commit on the default branch | 2026-09-04 | 2025-06-26 |
| Last npm publish | 2026-08-29 | 2025-06-26 |
| Downloads, last month | 2,019 | 11,116,327 |
| Stars | 2 | 6,114 |
| Open issues | 0 | 230 |
| Open pull requests | 0 | 64 |
| Source lines | 61,442 | 10,125 |
| Test lines | 64,581 | 0 |
| Test suite | 13 test scripts, 304 spec files under `test/` | no test script, no spec file, no test directory |
| Statement coverage | 95.33% (Node and browser lanes merged) | no automated suite |

The last commit on the default branch is reported rather than the repository's last push,
which the same API offers and which counts activity on any branch. The two disagree for
pptxgenjs by several months, and reporting the later one would say something the default
branch does not support.

Line counts come from the same walk on both sides: every code file under `src/`, raw lines
with comments and blanks included, and test lines are spec files plus anything under a
test directory, counted once each. No normalisation makes two libraries formatted to
different rules comparable, and a large part of the ts-pptx figure is the documentation
comments the bundled sizes above shed. Read it as an order of magnitude for how much there
is to maintain, and as nothing at all about whether it is good.

The empty pptxgenjs test row is what this walk can see, and it is not the same claim as
untested. That repository documents a manual, demo-driven process instead, which nothing
measured here can weigh. The row is about an automated suite, and the coverage figure
beside it exists for ts-pptx only because there is a suite to instrument.

ts-pptx is published under two names carrying the same bytes, `pptx-ts` and
`@shbernal/ts-pptx`. The download figure above is their sum (`pptx-ts` 295,
`@shbernal/ts-pptx` 1,724), because either name alone understates the total, and the
canonical name alone happens to understate it by most.

The pptxgenjs row shows no npm release since 2025-06-26 and no commit on `master` since
2025-06-26. That is what the two APIs report, and it is all this page says about it: from
outside, a stable library that has stopped needing changes looks exactly like one between
maintainers, and this measurement cannot tell them apart. It is worth weighing either way,
next to 230 open issues and 64 open pull requests.
