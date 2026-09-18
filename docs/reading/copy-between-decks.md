---
doc-schema-version: 1
title: "Copy slides between decks"
summary: "Duplicate a slide, import slides or single shapes from other open decks, place them, pick a theme mode, carry speaker notes, rescale across slide sizes, and build one deck from pages of several decks in an all-or-nothing batch with pptx-ts/read."
read_when:
  - Choosing between cloneSlide, importSlide, importSlides, importShape and appendSlides
  - Building one deck from pages of several library decks
  - Deciding whether an imported slide keeps its source theme or takes this deck's
  - Importing from a deck with a different slide size
  - Carrying speaker notes with an imported slide
  - Lifting a table, chart or picture from another deck onto an existing slide
doc_type: "guide"
---

# Copy slides between decks

`deck.importSlide(source, index)` copies a slide out of another open deck, with the layout, master, theme and media it needs:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { Presentation } from 'pptx-ts/read'

const deck = await Presentation.load(await readFile('deck.pptx'))
const library = await Presentation.load(await readFile('library.pptx'))
deck.importSlide(library, 2, { at: 0 })
await writeFile('deck-merged.pptx', await deck.save())
```

The copy of the library's third slide becomes the first slide of `deck`. Parts of `deck` the import does not touch are saved byte for byte, and `library` does not change.

## Choose a method

| Method | Source | Copies | Theme mode | Speaker notes | Different slide size | All or nothing |
| --- | --- | --- | --- | --- | --- | --- |
| [`cloneSlide`](../reference/api/read/classes/Presentation.md#cloneslide) | a slide of this deck | one slide, sharing its layout, master, theme and media | none, the copy keeps the source slide's chrome | copied | not applicable | checks the index, then copies |
| [`importSlide`](../reference/api/read/classes/Presentation.md#importslide) | another loaded deck | one slide, with its layout, master, theme and media | `copy` by default, or `preserve`, `restyle` | dropped unless `importNotes: true` | throws unless `rescale` is set | yes |
| [`importSlides`](../reference/api/read/classes/Presentation.md#importslides) | one or more loaded decks | several slides, each at its final position | `copy` only | per request, with `importNotes` | per request, with `rescale` | yes, for the whole batch |
| [`importShape`](../reference/api/read/classes/Presentation.md#importshape), [`importShapes`](../reference/api/read/classes/Presentation.md#importshapes) | a slide of any loaded deck | one or more top-level shapes, onto a slide of this deck | `preserve` by default, or `restyle`, `copy` | not applicable | throws unless `rescale` is set | yes |
| [`importSlideMasters`](../reference/api/read/classes/Presentation.md#importslidemasters) | another loaded deck | slide masters with their layouts, and no slides | not applicable, each master brings its theme | not applicable | throws unless `requireEqualSize: false`, and never rescales | yes |
| [`appendSlides`](../reference/api/read/classes/Presentation.md#appendslides) | a `TsPptx` generator | the generator's slides, bound to one layout of this deck | none, the slides bind to this deck's layout | carried | throws | checks the layout, size and links first |
| [`Presentation.fromTemplate`](../reference/api/read/classes/Presentation.md#fromtemplate) | a `.pptx` or `.potx` file | a new deck with the file's chrome and no slides | not applicable | removed with the slides | not applicable | not applicable |

`importSlide` and `importShape` default to different theme modes. A slide import brings the source layout, master and theme with the slide. A lifted shape lands on a slide that keeps this deck's chrome, so by default it bakes in the colours its source theme gave it. [Build on a template](build-on-a-template.md) covers `fromTemplate` and `appendSlides`.

## Duplicate a slide

```ts
const copy = deck.cloneSlide(0, { at: 1 })
```

- The copy takes the slide's current XML, including edits made since the deck was loaded.
- It shares the slide's layout, master, theme and images with the source slide.
- It gets its own copy of each part the slide owns: its notes slide, its charts with their workbooks, its SmartArt diagrams and its OLE embeddings. PowerPoint refuses to open a deck in which two slides point at one chart or one diagram. [Read object model](../reference/read-object-model.md#owned-vs-shared) lists which parts a slide owns and which it shares.
- `at` is a zero-based position in deck order. An omitted, negative or out-of-range `at` appends. `copy.index` is the position the slide landed at.

## Import a slide

`deck.importSlide(source, index, options)` copies `source.slides[index]` into this deck and returns the new slide. [`ImportSlideOptions`](../reference/api/read/interfaces/ImportSlideOptions.md) lists every option.

- The slide comes with its layout, that layout's master and theme, and its media, charts and embeddings, all under new part names.
- Only the layouts that imported slides use come across. An imported master lists only those layouts.
- Imports from one source deck share what they copied. A second import from the same deck reuses the layout, master, theme and media the first one brought.
- The slide part and the parts it owns are copied on every import, so importing one page twice gives two independent slides.
- A jump link from the page to another slide must point at a page an earlier import from the same deck brought across. Otherwise the import throws `import/unresolved-slide-link`. Import the target page first, or import both in one [batch](#import-several-slides-in-one-batch).
- Before it changes the deck, the import checks the index, the slide sizes, the links, and each source part it will read. A refused import leaves the deck unchanged.

### Place the imported slide

`at` works as in `cloneSlide`. `0` puts the slide first, and an omitted or out-of-range `at` appends. Each insert shifts the slides after it, so to set every position in one step, use [`importSlides`](#import-several-slides-in-one-batch).

To put brand slides around slides you generate, build the interior with `TsPptx`, load its bytes, and import the cover and the closer:

```ts
import { TsPptx } from 'pptx-ts'

const pptx = new TsPptx()
pptx.layout = 'LAYOUT_WIDE'
pptx.addSlide().addText('Interior', { x: 1, y: 1, w: 6, h: 1 })

const report = await Presentation.load(await pptx.toBytes())
const brand = await Presentation.load(await readFile('brand-library.pptx'))
report.importSlide(brand, 0, { at: 0 }) // cover, first
report.importSlide(brand, 5) // closer, appended
await writeFile('report.pptx', await report.save())
```

### Reuse chrome the deck already has

When this deck already holds a part identical to one the import would copy, the slide binds to that part. The usual case is a deck opened with `fromTemplate` from the same file you import from:

```ts
const bytes = await readFile('brand.pptx')
const shell = await Presentation.fromTemplate(bytes)
const source = await Presentation.load(bytes)
shell.importSlide(source, 3) // uses the shell's own layout, master and theme
```

A part of this deck counts as identical only when all of these hold:

- It has the same part name and content type.
- It has the same bytes as it would be saved with, so a part you edited in this deck does not count.
- It has the same relationships, and every part they reach passes the same test.
- A master is registered in `presentation.xml`, and a layout is listed by a registered master.

Anything else is copied. A part reached from a copied part is always copied too, so copied chrome never points into this deck's own parts.

## Import several slides in one batch

`deck.importSlides(requests)` imports pages from one or more decks and places each at its position in the final slide list:

```ts
const [cover, closer] = deck.importSlides([
  { source: libraryA, sourceIndex: 2, outputIndex: 0, importNotes: true },
  { source: libraryB, sourceIndex: 0, outputIndex: deck.slides.length + 1 },
])
```

- `outputIndex` is a position in the deck as it will be after the batch, so it counts the other pages of the batch. Above, the deck gains two slides, and the closer goes last.
- The returned array follows the order of `requests`, whatever the output positions.
- Each request takes `importNotes`, `embedFonts` and `rescale`. [`ImportSlidesRequest`](../reference/api/read/interfaces/ImportSlidesRequest.md) lists them.
- A batch always imports in `copy` mode. For `theme`, `carryMasterGraphics` or `remapLiterals`, use `importSlide`.
- `embedFonts` on one request carries that source deck's fonts once for the whole batch. See [Carry fonts when importing slides](../embedded-fonts.md#carry-fonts-when-importing-slides).
- An empty `requests` array returns `[]` and changes nothing.

### Batch validation rules

The batch checks every rule before it changes the deck. When one fails, it throws and the deck stays byte-identical.

| Rule | Code when broken |
| --- | --- |
| `sourceIndex` is a non-negative integer that names a slide of its source | `slide/index-out-of-range` |
| `outputIndex` is a non-negative integer less than the final slide count, which is the current count plus the number of requests | `import/output-index-out-of-range` |
| no two requests share an `outputIndex` | `import/output-index-conflict` |
| the requests naming one source give the same `rescale`, where `true` and `'fit'` count as the same | `import/rescale-conflict` |
| a source whose layout and master an earlier call rescaled is not rescaled in the other mode | `import/rescale-conflict` |
| without `rescale`, each source has the deck's slide size | `import/slide-size-mismatch` |
| every deck declares a slide size | `import/slide-size-unknown` |
| a jump link on a selected page points at another selected page, or at a page an earlier import from that source brought across | `import/unresolved-slide-link` |
| every source part the copy will read exists and parses, including the notes and font parts the requests ask for | a `PackageReadError`, such as `package/part-missing` |

The last rule runs the whole copy once without writing anything. A jump link between two selected pages is rewritten to point at the new slide part, so importing page 3 of 10 does not pull in pages 1 and 2.

### Copy one page more than once

Name the same source page in several requests to get independent copies of it:

```ts
const [before, after] = deck.importSlides([
  { source: library, sourceIndex: 4, outputIndex: 0 },
  { source: library, sourceIndex: 4, outputIndex: 1 },
])
```

- Each copy has its own slide part, its own slide id, and its own copy of each part the page owns. The layout, master, theme and media are copied once and shared.
- Pages duplicated together are copied in rounds. The first copy of a linking page links to the first copy of its target, and the second copy to the second.
- A link into a page requested only once lands on that single copy.

## Choose a theme mode

`importSlide` takes `theme`. `importShape` and `importShapes` take the same three values with a different default, as [Theme modes for lifted shapes](#theme-modes-for-lifted-shapes) sets out.

| Mode | What the import does | Result |
| --- | --- | --- |
| `copy` | Copies the slide's layout, master and theme. | The slide looks as it did. The deck gains a master and a theme for each source deck. |
| `preserve` | Writes the values the source theme gave the slide onto the slide, then binds it to this deck's first layout. | The slide looks as it did, and the deck keeps one theme. Text takes this deck's theme fonts. |
| `restyle` | Binds the slide to this deck's first layout, leaves its theme references symbolic, and drops its colour map override (`p:clrMapOvr`). | The slide takes this deck's theme colours and fonts. A colour written as a literal RGB value keeps that value. |

- The first layout is the first layout of the first slide master, in the order `presentation.xml` lists them. There is no option to pick another. A deck with no master, or with no layout under it, throws.
- `carryMasterGraphics: true`, with `preserve` or `restyle`, copies the shapes of the source master and layout onto the slide behind its own content, placeholders excepted, with their media. Logos and drawn footers then survive the rebind. Under `preserve` those shapes are baked like the rest of the slide, and under `restyle` they take this deck's colours too.
- `remapLiterals: true`, with `restyle` only, rewrites each literal colour that equals a colour of the source theme as that theme colour, so it takes this deck's value. It also copies each table style the slide uses from the source deck, unless this deck defines that style id already.
- A restyled slide needs a visual check. A source `accent1` that was light on a dark background can land on a background of the same shade in this deck.

### What preserve does

1. It reads the source slide's theme, colour map, layout and master.
2. It copies the slide and points its layout relationship at this deck's first layout. With `carryMasterGraphics`, it adds the source master and layout shapes behind the slide's own shapes, master shapes first.
3. It writes the slide's effective background onto the slide as literal values, taking the slide's own background, else the source layout's, else the source master's.
4. It resolves each shape's `p:style` line, fill and effect references against the source theme into explicit line, fill and effect.
5. It gives each placeholder that has no `a:xfrm` of its own the position and size of the matching placeholder on the source layout, else the source master.
6. It gives each run that sets no colour, size, bold or italic of its own the values it inherited from the source placeholder, layout and master text styles. A run outside a placeholder takes them from the source deck's default text style. A run in a table cell takes its size from the source master's `p:otherStyle`, and its bold and italic too where the table style states none.
7. It rewrites each `a:schemeClr` as the `a:srgbClr` the source theme gives it, keeping transforms such as `lumMod` and `shade`.

Typefaces (`a:latin`) and `fontRef` references stay symbolic, so text takes this deck's theme fonts. Steps 3 to 7 use the source values read in step 1, so binding to this deck's layout in step 2 does not change them.

## Carry speaker notes

Notes live in a part of their own, so every import drops them unless you pass `importNotes: true`:

```ts
const imported = deck.importSlide(library, 0, { importNotes: true })
console.log(imported.notesText)
```

- The notes part is copied, and its link back to its slide points at the new slide.
- Media and hyperlinks in the notes come across like the slide's own.
- The notes take this deck's notes master when it has one, and the source deck's otherwise, as [One notes master per deck](read-and-edit.md#one-notes-master-per-deck) sets out.
- In `importSlides`, `importNotes` is set per request. When the deck has no notes master, the first request that carries notes installs its source's notes master. The dry run reads a source's notes master only when the batch would copy it.
- A page named twice in a batch gets a notes part for each copy.
- To write notes of your own on an imported slide, use [`addNotes`](read-and-edit.md#read-and-write-speaker-notes).

## Import from a deck with a different slide size

An import throws `import/slide-size-mismatch` when the two decks' slide sizes differ. `importSlide`, each `importSlides` request, `importShape` and `importShapes` take `rescale` to scale the copy onto this deck's slide instead:

| `rescale` | Scaling |
| --- | --- |
| `false`, or omitted | none. A size mismatch throws. |
| `'fit'`, or `true` | one factor for both axes, the smaller of the width and height ratios, with the leftover space split evenly on both sides. Circles stay round. |
| `'stretch'` | the width and height ratios, each on its own axis. Shapes distort. |

```ts
const legacy = await Presentation.load(await readFile('legacy-4x3.pptx'))
deck.importSlide(legacy, 0, { rescale: 'fit' })
```

- Only geometry scales. That is the position and size of each top-level shape, group and graphic frame, and the column widths and row heights of tables. Font sizes and line widths stay as authored, so text can overflow a box that shrank.
- In `copy` mode the imported layout and master scale too, so placeholders that inherit their geometry stay aligned. Each layout and master scales once, however many slides use it.
- In `preserve` and `restyle` mode the slide binds to this deck's layout, which already has the right size, so only the slide scales.
- Once an import has rescaled a source's layout and master, a later import from that source in the other mode throws `import/rescale-conflict`.
- `importShape` applies `left`, `top`, `width` and `height` after the rescale, so they win.
- Both decks must declare a slide size (`p:sldSz`). Without one, the import throws `import/slide-size-unknown`, with or without `rescale`.
- `appendSlides` and `importSlideMasters` do not rescale.

## Lift shapes from other decks

`deck.importShape(target, source, shapeIndex, options)` copies one shape from a slide of any open deck onto `target`, a slide of this deck, and returns the new shape. `importShapes` takes a list of indexes and returns the shapes in that order:

```ts
const host = deck.slides[0]
const comparison = libraryA.slides[38]
const icons = libraryB.slides[34]
if (host && comparison && icons) {
  deck.importShape(host, comparison, 2, { left: 457200, top: 1371600 })
  deck.importShapes(host, icons, [4, 5, 6], { rescale: 'fit', carryAnimation: true })
}
```

- `shapeIndex` counts `source.shapes`, the slide's top-level shapes. A shape inside a group comes across only with its group.
- Media, charts with their workbooks, and embeddings the shape uses come into this deck under new part names. Imports from one source deck copy a shared image once, and shapes of one call that share an image share one relationship. Each chart or diagram gets its own copy.
- Each lifted shape and group child gets a drawing id unused on the host slide, and connectors inside the lifted shapes stay attached.
- `left`, `top`, `width` and `height`, in EMU, replace the source position and size. They apply to every shape of the call, so shapes lifted in one call with a `left` all land at that `left`. Lift shapes in separate calls to place them apart.
- `at` is the z-order position among the host slide's shapes, where `0` is the back. An omitted or out-of-range `at` puts the shapes on top. `importShapes` inserts them in list order from `at`.
- A shape's build animation lives in the slide's timing, so it arrives static. `carryAnimation: true` copies its animation into the host slide's timing with the new ids.
- Embedded fonts do not come with shapes.
- The call checks the target, the indexes, the slide sizes, the position values, and each part it will copy, before it inserts the first shape.

[`ImportShapeOptions`](../reference/api/read/interfaces/ImportShapeOptions.md) lists every option.

### Theme modes for lifted shapes

- `preserve`, the default, bakes the shape's scheme colours and style references against the source theme, as a slide import does, and never touches a background. A lifted placeholder also gets its inherited position, size, colour, text size, vertical anchor and list style. It then loses its placeholder identity (`p:ph`), so it cannot inherit from, or clash with, a placeholder on the host slide.
- `restyle` leaves theme references symbolic and keeps `p:ph`, so the shape takes the host deck's theme.
- `copy` brings the XML across unchanged. Use it when the host deck has the source deck's theme.

## Invalid input

| Condition | Result | Code |
| --- | --- | --- |
| `index` or `sourceIndex` names no slide | throws `InvalidOptionError` | `slide/index-out-of-range` |
| `shapeIndex` names no top-level shape | throws `InvalidOptionError` | `shape/index-out-of-range` |
| the `importShape` target slide is not a slide of this deck | throws `InvalidOptionError` | `slide/foreign-target` |
| the slide sizes differ and `rescale` is not set | throws `InvalidOptionError` | `import/slide-size-mismatch` |
| a deck declares no slide size | throws `InvalidOptionError` | `import/slide-size-unknown` |
| `rescale` differs between requests naming one source, or from an earlier import of that source | throws `InvalidOptionError` | `import/rescale-conflict` |
| `outputIndex` is negative, not an integer, or past the final slide list | throws `InvalidOptionError` | `import/output-index-out-of-range` |
| two requests share an `outputIndex` | throws `InvalidOptionError` | `import/output-index-conflict` |
| a jump link on an imported page points at a page that is not imported | throws `InvalidOptionError` | `import/unresolved-slide-link` |
| `theme: 'preserve'` or `'restyle'` into a deck with no slide master | throws `InvalidOptionError` | `import/destination-missing-master` |
| `theme: 'preserve'` or `'restyle'` into a deck whose first master has no layout | throws `InvalidOptionError` | `import/destination-missing-layout` |
| a source part the copy needs is missing | throws `PackageReadError` | `package/part-missing` |
| `importShape` `left` or `top` is not a finite number | throws `InvalidOptionError` | `coord/non-finite` |
| `importShape` `width` or `height` is zero or less | throws `InvalidOptionError` | `coord/not-positive` |

Each of these is found before the deck changes.

## Limits

- `importSlides` imports in `copy` mode only.
- `preserve` and `restyle` always bind to the first layout of this deck's first master.
- `rescale` scales geometry, not font sizes or line widths.
- `appendSlides` and `importSlideMasters` never rescale.
- `restyle` recolours symbolic colours only, unless `remapLiterals` is set.
- A shape inside a group cannot be lifted without its group.
- A position override on `importShapes` applies to every shape of the call.
- Shape imports carry no embedded fonts. A slide import with `embedFonts` copies the source deck's whole font list.
- A jump link comes across only when its target page is imported first, or in the same batch.

## See also

- [Read and edit a deck](read-and-edit.md)
- [Build on a template](build-on-a-template.md)
- [Embedded fonts](../embedded-fonts.md)
- [Read object model](../reference/read-object-model.md)
- [Round-trip guarantee](../reference/round-trip.md)
- [Errors and warnings](../errors-and-warnings.md)
- API reference: [`Presentation`](../reference/api/read/classes/Presentation.md), [`ImportSlideOptions`](../reference/api/read/interfaces/ImportSlideOptions.md), [`ImportSlidesRequest`](../reference/api/read/interfaces/ImportSlidesRequest.md), [`ImportShapeOptions`](../reference/api/read/interfaces/ImportShapeOptions.md), [`ImportSlideMastersOptions`](../reference/api/read/interfaces/ImportSlideMastersOptions.md)
