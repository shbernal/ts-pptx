---
doc-schema-version: 1
title: "How the comparison was measured"
summary: "The corpus, the four outcomes and every full table behind the comparison of ts-pptx 3.7.0 with pptxgenjs 4.0.1: validation, installs, bundles, generation time and project health, and how each was taken."
read_when:
  - Checking a number on the comparison page
  - Adding a probe or a program to the comparison corpus
  - Re-measuring the comparison for a release
doc_type: "reference"
---

<!-- GENERATED FILE. Do not edit by hand.
     Regenerate with `pnpm run comparison:render`.
     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->

# How the comparison was measured

Every figure on [ts-pptx vs PptxGenJS](comparison.md) comes from
`scripts/comparison/snapshot.json`, which is refreshed on release cadence, and nothing on
either page is edited by hand. This page is the method behind those figures and the full
tables they summarise.

Measured on 2026-09-15: ts-pptx 3.7.0 built from this repository, against pptxgenjs 4.0.1
installed from npm (published 2025-06-26).

## The corpus

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

## Construct coverage

Of 22 probes, ts-pptx emitted 21 and pptxgenjs emitted 10.

"Looked for" is the token the harness reads for, in the part named beside it. It is the
OOXML element in every case but one, where the intent is speaker notes and the token is
the note text itself. [Porting from PptxGenJS](comparison-syntax.md) prints the calls
behind every row.

### Shared baseline

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Text run | `<a:t>` | `ppt/slides/slide1.xml` | emitted | emitted |
| Table | `<a:tbl>` | `ppt/slides/slide1.xml` | emitted | emitted |
| Raster image | `<p:pic>` | `ppt/slides/slide1.xml` | emitted | emitted |
| Bar chart | `<c:barChart>` | `ppt/charts/chart1.xml` | emitted | emitted |
| External hyperlink | `<a:hlinkClick` | `ppt/slides/slide1.xml` | emitted | emitted |
| User-defined slide master | `<p:ph` | `ppt/slides/slide1.xml` | emitted | emitted |
| Sections | `<p14:sectionLst` | `ppt/presentation.xml` | emitted | emitted |
| Speaker notes | `probe note` | `ppt/notesSlides/notesSlide1.xml` | emitted | emitted |
| Preset-geometry shape | `<a:prstGeom` | `ppt/slides/slide1.xml` | emitted | emitted |
| Slide background colour | `<p:bg>` | `ppt/slides/slide1.xml` | emitted | emitted |

### Motion

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Slide transition | `<p:transition` | `ppt/slides/slide1.xml` | emitted | no API |
| Build animation on a shape | `<p:timing>` | `ppt/slides/slide1.xml` | emitted | no API |

### Embedding

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Embedded OLE object | `<p:oleObj` | `ppt/slides/slide1.xml` | emitted | no API |
| 3D model | `am3d:model3d` | `ppt/slides/slide1.xml` | emitted | no API |
| Embedded font face | `<p:embeddedFontLst>` | `ppt/presentation.xml` | emitted | no API |

### Shapes

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Connector between shapes | `<p:cxnSp>` | `ppt/slides/slide1.xml` | emitted | no API |

### Text

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Inline equation | `<a14:m` | `ppt/slides/slide1.xml` | emitted | no API |

### Fills

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Gradient shape fill | `<a:gradFill` | `ppt/slides/slide1.xml` | emitted | no API |

- Gradient shape fill, pptxgenjs: appears only inside the bundled Office theme XML, which
  no API parameterises.

### Tables

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| 3D bevel on a table cell | `<a:cell3D` | `ppt/slides/slide1.xml` | emitted | no API |

### Charts

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Funnel chart (chartEx) | `<cx:chart>` | `ppt/charts/chartEx1.xml` | emitted | no API |

### Navigation

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| Slide Zoom tile | `pslz:sldZm` | `ppt/slides/slide2.xml` | emitted | no API |

### Diagrams

| Intent | Looked for | Part | ts-pptx | pptxgenjs |
|---|---|---|---|---|
| SmartArt diagram (write side) | `<dgm:relIds` | `ppt/slides/slide1.xml` | no API | no API |

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
  '{f07f2693-3e62-fdb6-0d4c-93b723073eea}'. The Pattern constraint failed. The expected
  pattern is \{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}.`

## Package hygiene

What a consumer gets. Each library was installed on its own into an empty directory,
upstream from the registry and ts-pptx from a pack of this working tree, so nothing here
is measured against a development checkout with its dependencies hoisted flat.

|  | ts-pptx | pptxgenjs | Difference |
|---|---|---|---|
| Installed size, with dependencies | 10.7 MiB | 6.6 MiB | +61% |
| Installed size, the package alone | 6.0 MiB | 2.5 MiB | +142% |
| Runtime dependencies, transitive | 3 | 18 | -83% |

The last column is ts-pptx measured against pptxgenjs, so a positive number is ours
costing more and a negative one is ours costing less. It is a percentage of the pptxgenjs
figure rather than a difference in bytes, because two of its rows are in mebibytes and the
third is a count, and a reader comparing them needs a number that does not change meaning
between rows.

ts-pptx installs larger than pptxgenjs despite carrying fewer dependencies. The largest
share of that weight is source maps: `dist/` ships a `.js.map` beside every module, and
each one embeds the original TypeScript. The unminified `.js` is the next largest share.
No consumer build keeps either, which is why the bundled figures below are much closer
together than the installed ones.

|  | ts-pptx | pptxgenjs |
|---|---|---|
| Runtime dependencies, direct | `@xmldom/xmldom`, `fflate`, `opentype.js` | `@types/node`, `https`, `image-size`, `jszip` |
| Entry points | `.`, `./inspect`, `./measure`, `./read`, `./script`, `./math`, `./zip`, `./html`, `./families`, `./node`, `./browser` | `.` |
| Module formats | esm | cjs, esm |
| `engines.node` | `>=24` | not declared |

### What a bundled program costs

Each row is a whole deck both libraries build: 5 consumer programs, from the smallest one
anyone writes up to one using every construct the shared baseline above shows both of them
emitting. Each is written in its own idiom on both sides, and the calls behind every row
are on [porting from PptxGenJS](comparison-syntax.md).

- **Hello world.** One slide with one text box.
- **Text deck.** A defined master, two sections, formatted and bulleted text, a hyperlink,
  a slide background and speaker notes.
- **Table deck.** A titled slide and a bordered table with a header row, fixed column
  widths and per-cell options.
- **Chart deck.** Three charts on three slides: a column chart with value labels, a
  two-series line chart, and a pie chart with percentages.
- **Full deck.** Every construct the shared baseline covers, in one deck: master,
  sections, background, text, hyperlink, notes, a preset shape, an image, a table and a
  chart.

| Program | ts-pptx | pptxgenjs | Difference |
|---|---|---|---|
| Hello world | 103.3 KiB | 123.3 KiB | -16% |
| Text deck | 103.6 KiB | 123.6 KiB | -16% |
| Table deck | 103.4 KiB | 123.5 KiB | -16% |
| Chart deck | 103.5 KiB | 123.5 KiB | -16% |
| Full deck | 103.8 KiB | 123.8 KiB | -16% |

The column is nearly flat, and that is the result. From hello world to full deck, ts-pptx
grows by 0.5 KiB and pptxgenjs by 0.5 KiB, which is about what the programs' own literals
weigh. Neither library splits along feature lines: importing either one costs almost
everything it will ever cost, and the deck written afterwards is close to free. So a hello
world was never a flattering measurement of either library, and a consumer weighing bundle
size is choosing between two roughly fixed costs rather than between two slopes.

Both columns construct the library the way every consumer of pptxgenjs constructs it, with
the class that carries everything. ts-pptx has a lower floor than that, reached by
composing a presentation from only the construct families a program uses, and [smaller
bundles](bundle-size.md) carries those figures. It is deliberately not a row here:
pptxgenjs has no counterpart to compose, so the cell beside it would be empty and the
percentage would be comparing two different programs.

Every program is identical in intent on both sides. Each is bundled with esbuild for the
browser, minified, and gzipped at level 9, following the conventions
`scripts/bundle-size-ratchet.mjs` documents, with one difference that matters. The ratchet
never bundles, so it cannot drop unreachable code and its figures are an upper bound on
what the package ships; this bundles and does tree-shake, because a consumer's build is
precisely the thing being compared here. **The two sets of numbers will not agree, and
neither is wrong.**

Code splitting is on, so each figure is the entry chunk: what the program pays before its
first line runs. Chunks a bundler defers behind a dynamic import are not counted, because
a program that never takes that path never fetches them.

Each program is run against both libraries before it is bundled. A bundler compiles a call
that does not exist, and a misspelled method comes out as a smaller bundle rather than as
an error, because the tree-shaker keeps less: running the programs first is what stops a
typo being published here as a saving.

## Generation time

How long each library takes to turn a deck into bytes: the same decks the bundle table
above weighs, plus three larger ones built for this measurement alone, because the largest
program up there is three slides and a clock has almost nothing to see in it. The calls
behind every row are on [porting from PptxGenJS](comparison-syntax.md).

**A `.pptx` is a zip, so the compression setting is not a detail of this measurement, it
is the measurement.** The two libraries do not default to the same one. ts-pptx deflates
unless told not to. pptxgenjs passes no compression option to JSZip on the `outputType`
path, and JSZip stores by default. Its own `compression` argument is honoured on the
stream and browser paths and ignored on the one in between, which is the path `writeFile`
takes in Node. Timing the two default calls against each other would compare deflating
with not deflating and report the difference as a library being slow, so both tables below
are matched pairs.

### Compressed

Both libraries asked for a compressed deck, which is what a consumer writing a file they
intend to keep gets. This is the table that matters.

| Deck | ts-pptx | pptxgenjs | Difference |
|---|---|---|---|
| Hello world | 11 ms | 23 ms | -51% |
| Text deck | 18 ms | 27 ms | -35% |
| Table deck | 14 ms | 22 ms | -35% |
| Chart deck | 29 ms | 42 ms | -31% |
| Full deck | 20 ms | 26 ms | -21% |
| 50 slides | 212 ms | 264 ms | -20% |
| 200 slides | 719 ms | 962 ms | -25% |
| 500 slides | 1941 ms | 2493 ms | -22% |

### Stored, the control

Both libraries asked not to compress. The zip drops out of the measurement, leaving each
library's own work: building the XML and assembling the package.

| Deck | ts-pptx | pptxgenjs | Difference |
|---|---|---|---|
| Hello world | 4.6 ms | 3.5 ms | +33% |
| Text deck | 5.5 ms | 4.3 ms | +26% |
| Table deck | 4.7 ms | 4.4 ms | +6% |
| Chart deck | 14 ms | 15 ms | -5% |
| Full deck | 12 ms | 9.4 ms | +31% |
| 50 slides | 142 ms | 117 ms | +21% |
| 200 slides | 533 ms | 442 ms | +21% |
| 500 slides | 1344 ms | 1092 ms | +23% |

The measurement is a median over repeated rounds, taken after a warm-up that is thrown
away, with the two libraries interleaved and the order alternated so that a machine which
slows down mid-run cannot hand either column a result it did not earn. Building the deck
and writing it are both inside the clock; constructing the presentation object is not.

**The milliseconds belong to the machine that took them and do not transfer; the ratios
mostly do.** These were taken on Intel(R) Core(TM) Ultra 5 235U (14 cores) under Node
v24.20.0 on win32, on 2026-09-15. Repeating a run on the same machine moves a difference
by a few points in either direction, so read the columns for their direction and rough
size rather than for their last digit.

## Project health

The [health table](comparison.md#adoption-and-project-health) describes how the two
projects are run, not what either one emits. Stars and downloads measure adoption,
adoption measures history as much as merit, and none of it belongs in the same table as a
construct a library does or does not write. This is how each figure in it is taken.

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
`@shbernal/ts-pptx`. The download figure is their sum (`pptx-ts` 356, `@shbernal/ts-pptx`
1,324), because either name alone understates the total, and the canonical name alone
happens to understate it by most.

The pptxgenjs column shows no npm release since 2025-06-26 and no commit on `master` since
2025-06-26. That is what the two APIs report, and it is all these pages say about it: from
outside, a stable library that has stopped needing changes looks exactly like one between
maintainers, and this measurement cannot tell them apart. It is worth weighing either way,
next to 232 open issues and 65 open pull requests.
