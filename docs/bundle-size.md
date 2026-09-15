---
doc-schema-version: 1
title: "Smaller bundles"
summary: "Starting a deck with createPresentation and only the construct families a program uses, what each family authors, and what a browser program downloads either way."
read_when:
  - Choosing between TsPptx and createPresentation
  - Working out which construct families a program needs
  - Cutting what a browser program downloads
doc_type: "guide"
---

# Smaller bundles

Every entry offers two ways to start a deck. `new TsPptx()` can author everything the library
supports, and a program built on it downloads all of it. `createPresentation({ use })` starts from a
small core and adds only the construct families you name. A family you do not name is not in your
bundle:

```ts
import { ChartType, createPresentation } from "pptx-ts"
import { charts } from "pptx-ts/families"

const pres = createPresentation({ use: [charts] })
pres.addSlide().addChart(
  [{ name: "Revenue", labels: ["Q1", "Q2"], values: [1.2, 1.5] }],
  { type: ChartType.bar, x: 1, y: 1, w: 8, h: 4 },
)
```

For the same calls, both write the same deck, part for part. Families are values rather than strings,
so naming one is what puts its code in your program, and a bundler needs no configuration to leave
the others out.

## What a program downloads

Each row is a browser program bundled with esbuild against the current build, minified and gzipped.
`initial` is what the browser fetches before the program's first line runs. `total` adds every chunk
the program can reach later, such as the font parser, which loads only when a font is first
registered.

| Program | initial | total |
| --- | --- | --- |
| `createPresentation()`, one text box | 59.4 KiB | 138.2 KiB |
| `createPresentation({ use: [charts] })`, plus a chart | 80.5 KiB | 159.3 KiB |
| `new TsPptx()`, one text box | 98.7 KiB | 177.4 KiB |
| `new TsPptx()`, plus a shape and an image | 98.9 KiB | 177.6 KiB |
| `new TsPptx()`, plus a chart, a table and a video | 99.0 KiB | 177.7 KiB |

A text-only program downloads 59.4 KiB composed against 98.7 KiB through the class, so composing saves
39.3 KiB before the first line runs. The chart family costs 21.1 KiB to the program that asks for it and
nothing to one that does not. The three `TsPptx` rows sit within 0.3 KiB of each other, because the
class carries every family whatever the program calls.

The [comparison page](comparison.md) reports 98.6 KiB for a hello world. That is a separate
measurement, of a different program, taken with the rest of the comparison on its snapshot date. The
figures here come from the repository's size gate and follow the current build; the
[testing guide](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/testing.md#size-gates)
describes how both are measured.

## The core tier

Five families are always composed, so `createPresentation()` with no argument still writes a deck:

| Family | What it authors |
| --- | --- |
| `text` | text boxes |
| `shapes` | preset shapes, and the `rect` / `line` / `roundRect` descriptors |
| `images` | raster and SVG pictures |
| `groups` | `addGroup`, and `groupObjects` over what is already on the slide |
| `notes` | speaker notes and the notes slides they are written to |

A family is core when leaving it out would make the result something other than a deck. All five are
exported from `pptx-ts/families` too, so you can list one to be explicit; listing it changes nothing.

## Everything else is asked for

| Family | What it authors |
| --- | --- |
| `charts` | every chart type, with their embedded workbooks. The largest family |
| `tables` | tables, including tables that continue onto new slides |
| `domTables` | `tableToSlides` from a `<table>` a browser has laid out. Needs a live DOM |
| `media` | embedded and online audio and video |
| `connectors` | lines drawn between two points or two shapes |
| `oleObjects` | embedded OLE objects that travel inside the `.pptx` |
| `models3d` | embedded `.glb` models with their preview picture |
| `zooms` | Slide, Section and Summary Zoom tiles |
| `comments` | review comments and the deck-wide author list |
| `animations` | preset build animations |
| `measure` | `measureText`, `overflowsBox` and `tableLayout` on the presentation |

`domTables` reads a `<table>` a browser has already laid out, so it belongs to a program that has a
`document`, and `createPresentation` from `pptx-ts/browser` takes it. The DOM-agnostic form of the
same conversion is the free `tableToSlides` on `pptx-ts/html`, which costs a composed deck nothing.

`measure` supplies three presentation methods that measure without authoring anything. Text with
`fit: 'shrink'` is still shrunk when the deck is written without it; see
[Text that fits](text-fit.md#measure-text-before-export).

## When a family is missing

Calling a method whose family you did not compose throws `family/not-composed`, naming the family, at
the call. The method still exists so that the error can say which family to add, rather than
`TypeError: slide.addChart is not a function`.

The types say it first. The slide `createPresentation` returns carries only the methods your families
supply, and each of those methods returns that same slide, so a chain cannot widen back to the full
set of methods.

Child descriptors are the one thing the types cannot catch. The `{ chart: ... }`, `{ image: ... }` and
`{ text: ... }` forms that a slide master's `objects` and a group's children use type-check whenever
some family claims the key, composed or not. Both walks warn with `family/child-not-composed`, naming
the family to compose. A key no family claims reports `group/unrecognized-child`, which is usually a
typo.
