---
doc-schema-version: 1
title: "Bundle size"
summary: "The two ways to construct a presentation, which construct families you pay for, and the two size gates that keep the numbers honest."
read_when:
  - Choosing between TsPptx and createPresentation
  - Working out which construct families a program needs
  - Investigating a bundle-size regression or a failing size gate
  - Changing what a construct family reaches, or adding one
doc_type: "reference"
---

# Bundle size

Two questions sound like one and are not. *What does this package ship?* is answered
per published entry, and the answer is an upper bound. *What does my program
download?* is answered by bundling a real program and letting a bundler shake it.
This page covers the choice a consumer makes, then the two gates, one per question.

## Paying for what you author

Every entry publishes two ways to start a deck.

`TsPptx` is composed with every construct family the library has. It authors
everything and costs everything, and that will not change.

`createPresentation` is composed with the core tier plus the families you name. A
family you do not name is not in your bundle:

```ts
import { createPresentation } from "pptx-ts"
import { charts } from "pptx-ts/families"

const pres = createPresentation({ use: [charts] })
pres.addSlide().addChart(data, { type: ChartType.bar })
```

Both write the same deck for the same slides, part for part. What changes is reach,
not behaviour. There is no second write path here.

The families are values rather than strings, so naming one is what puts its code in
the module graph. A bundler needs no configuration to leave the rest out.

### What it saves

The figures below come from `pnpm run bundle-tier:list`, which is the same
measurement `bundle-tier:check` gates. Each row is a bundled browser program,
minified and gzipped. `initial` is what the program fetches before its first line
runs; `total` adds every chunk it can reach (see [Two figures, not
one](#two-figures-not-one)).

| program | initial | total |
| --- | --- | --- |
| `createPresentation()`, one text box | 59.4 kB | 138.2 kB |
| `createPresentation({ use: [charts] })`, plus a chart | 80.5 kB | 159.3 kB |
| `new TsPptx()`, one text box | 98.7 kB | 177.4 kB |
| `new TsPptx()`, plus a shape and an image | 98.9 kB | 177.6 kB |
| `new TsPptx()`, plus a chart, a table and a video | 99.0 kB | 177.7 kB |

Two things to read out of that table. A text-only program pays 59.4 kB composed
against 98.7 kB through the class, so composing saves 39.3 kB of blocking download,
two fifths of it. And the chart family costs 21.1 kB to the program that asks for it
and nothing to the program that does not.

The three `TsPptx` rows sit within 0.3 kB of each other, which is the point of the
control: the class links every family whatever the program calls, so its size tracks
the composition and not the call list.

### The core tier

Five families are composed for you, so `createPresentation()` with no argument still
writes a deck rather than a slide-shaped hole:

| family | what it authors |
| --- | --- |
| `text` | text boxes |
| `shapes` | preset shapes, and the `rect` / `line` / `roundRect` descriptors |
| `images` | raster and SVG pictures |
| `groups` | `addGroup`, and `groupObjects` over what is already on the slide |
| `notes` | speaker notes and the notes slides they are written to |

The line is a judgement, not a measurement. A family is core when leaving it out
would make the result something other than a deck. These five are cheap, and a
surface that made someone compose `text` before writing a word would be a worse
default than one that costs a few kilobytes more. All five are exported from
`pptx-ts/families` anyway, so you can list one to be explicit; listing it changes
nothing.

### Everything else is asked for

| family | what it authors |
| --- | --- |
| `charts` | every chart type, the embedded workbooks, the chartEx sidecars. The expensive one |
| `tables` | tables, including auto-paging onto continuation slides |
| `domTables` | `tableToSlides` off a rendered `<table>`. Needs a live DOM |
| `media` | embedded and online audio and video |
| `connectors` | lines drawn between two points |
| `oleObjects` | embedded OLE payloads that travel inside the `.pptx` |
| `models3d` | embedded `.glb` models with their preview picture |
| `zooms` | Slide, Section and Summary Zoom tiles |
| `comments` | review comments and the deck-wide author list |
| `animations` | preset build animations |
| `measure` | `measureText`, `overflowsBox` and `tableLayout` on the presentation |

`domTables` reads a browser's laid-out `<table>`, so it belongs to a program that has
a `document`. `createPresentation` from `pptx-ts/browser` takes it; the DOM-agnostic
form of the same conversion is the free `tableToSlides` on `pptx-ts/html`, which
costs a composed deck nothing.

`measure` supplies three presentation methods that measure without authoring
anything. The export-time autofit bake is not part of it: `fit:'shrink'` text is
baked on the ordinary write path, so a composed deck that never asks for a
measurement still gets its autofit. See [Measured text
fit](measured-text-fit.md#instance-methods-inchespoints-reuse-registered-metrics).

### When a family is missing

Calling a method whose family you did not compose raises `family/not-composed`,
naming the family, at the call. A method that is not composed still exists for this
reason: leaving it off would report the same condition as `TypeError:
slide.addChart is not a function`, which names neither the family nor the fix.

The types say it first. The slide `createPresentation` hands you carries only the
methods your families supply, and each of those methods answers that same slide, so
a chain cannot widen back to the full surface.

Child descriptors are the one thing the types cannot catch. The `{ chart: ... }`,
`{ image: ... }`, `{ text: ... }` forms that a slide master's `objects` and a group's
children are written with are rejected only when *no* family anywhere claims the key,
so a key some family claims type-checks whether or not you composed it. Both walks
warn instead, with `family/child-not-composed` naming the family to compose. A key no
family has ever claimed still reports `group/unrecognized-child`, which is the typo
it always was.

## What a program downloads

`scripts/bundle-tier-size.mjs` measures what one program pays. It bundles five
consumer programs against `dist/browser.js` with esbuild (minified, gzipped,
code-split) and freezes both figures each one produces in
`scripts/bundle-tier-budget.json`. `check:package` enforces it alongside the
entry-point budget.

The programs come in two shapes, and the pair is deliberate. The three `new TsPptx()`
rows are cumulative: `text` writes one slide with one text box, `text-shape-image`
adds a shape and a base64 image, and `full` adds a chart, a table and an embedded
video. The two `composed` rows are not cumulative. They are the same program with and
without one family, so the difference between them is what that family costs. Real
programs rather than the smallest call the types accept, because a synthetic minimum
measures the type checker.

### Two figures, not one

`initial` is the entry chunk plus every chunk reachable from it by an `import`
statement: what a browser fetches before the first line runs. `total` is every chunk
the program can reach, which matters because font metrics load `opentype.js` through
a dynamic import that runs only on first font registration. Charging a program for a
chunk it may never fetch is as wrong as hiding bytes it might. There is no fair
single number, so there is no single number.

### The deferred poster

The library's own largest deferred chunk is the default video poster. When a caller
passes no `cover`, `addMedia` falls back to a play-button overlay of 19,312 base64
characters. `addMedia` is a class method, so nothing tree-shakes it. Naming the artwork
where the media object is defined therefore charged that payload to every consumer,
text-only ones included. It is resolved during the async media pass instead, from
`src/media/playbtn.ts`, which holds nothing else so that the chunker can give it a
chunk of its own. A program that writes no media never fetches it, and one that does
fetches it only at export.

The artwork used to be nearly four times that. It was a pasted blob, and a pasted blob
carries no evidence of whether its size is detail or noise: 74,380 characters for a
frame holding four flat colours, plus a stray olive border nobody had looked closely
enough to see. `scripts/gen-playbtn.mjs` draws it instead, from geometry measured off
the original, and `test/scripts/gen-playbtn.test.js` holds the module and the
generator together.

### Each program is run before it is weighed

esbuild resolves modules and does not care whether `slide.addChart` exists. A renamed
or dropped method would leave every program bundleable, the family that method reached
would fall out of the graph, and the number would go *down*, which is the shape of a
win. Running each program against `dist/` first makes that a failure instead.

`pnpm run bundle-tier:list` prints the per-chunk breakdown;
`pnpm run bundle-tier:freeze` re-baselines.

## What the package ships

`scripts/bundle-size-ratchet.mjs` freezes a budget for every entry point
`package.json` publishes and the chunks each one pulls in, minified and then gzipped,
and `pnpm run check:package` enforces it. Per entry rather than for the package as a
whole, because the question a consumer asks is what importing *one* subpath costs; the
shared chunks are counted once per entry that reaches them.

Minified, because `dist/` ships unminified and is close to half doc comments by weight,
none of which survives a consumer's build. Gating on the raw bytes made the number track
how much the code was *documented*. The refactor series between v3.7.0 and 147951de took
14.5 kB of code out of the browser closure and added 26.9 kB of comments explaining the
consolidations. The gate reported the net as a 10.2 kB regression. Minifying first takes
prose out of the measurement.

It is an upper bound rather than a download size, because a consumer's bundler also
tree-shakes across the closure and this deliberately does not. What the gate is for is
the step change: a dependency reaching the browser entry, or a chunk split going
wrong.

`pnpm run bundle-size:list` prints the per-chunk breakdown; the budget lives in
`scripts/bundle-size-budget.json` and is raised or lowered deliberately with
`pnpm run bundle-size:freeze`.

## Why both gates exist

They answer different questions and their numbers will not agree.

The entry-point ratchet never bundles, so it cannot see reachability. Making a
construct family unreachable for a program that never calls into it deletes not one
byte from `dist/`, and the ratchet does not move. The tier gate bundles, so a bundler
shakes it, and it is the only gate that moves when code stops being *reachable*
rather than stopping being *shipped*.

The reverse is just as true. A chunk that grows, or a dependency that arrives on the
browser entry, shows up in the ratchet whether or not any measured program reaches it.

Neither replaces the other. A change that improves one and leaves the other flat is
usually working as intended.

## Keeping it true

The whole design rests on one rule, and
[Architecture](architecture.md#the-rule-that-keeps-the-tiers-real) states it: a static
import from `slide.ts`, `gen/slide/object.ts` or `package/assemble.ts` into a family
module is what the tier budget is watching for. Break it and every composed program
pays for every family again, silently, while every test still passes.
