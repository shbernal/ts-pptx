---
doc-schema-version: 1
title: "Read and edit a deck"
summary: "Open an existing .pptx with pptx-ts/read, find shapes by name, id or placeholder, replace text, change position and colour, add and remove text boxes and pictures, read and write speaker notes, replace a picture's image, and edit the XML directly, then save with untouched parts byte for byte."
read_when:
  - Opening a deck this library did not write, to read it or change it
  - Replacing the text of a named shape or placeholder
  - Adding a text box or picture to an existing slide, or deleting a shape
  - Reading or writing speaker notes
  - Replacing the image behind a picture
  - Changing XML that no typed setter covers
doc_type: "guide"
---

# Read and edit a deck

`pptx-ts/read` opens an existing `.pptx`, gives you its slides, shapes and text to read and change, and saves the result:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { Presentation } from 'pptx-ts/read'

const deck = await Presentation.load(await readFile('deck.pptx'))
for (const slide of deck.slides) {
  for (const shape of slide.shapes) console.log(slide.index, shape.shapeType, shape.name, shape.text)
}

const title = deck.slides[0]?.placeholder('title')
if (title) title.text = 'Quarterly review'
await writeFile('deck-edited.pptx', await deck.save())
```

`save()` writes again only the parts you changed. Every other part keeps the bytes it was loaded with. [Round-trip guarantee](../reference/round-trip.md#what-save-writes) states the fidelity contract, and [Read object model](../reference/read-object-model.md) lists what each class reads.

## Load and save

- `Presentation.load(input)` takes the package as a `Uint8Array`, an `ArrayBuffer`, a `Blob` or a `number[]`. A Node `Buffer` is a `Uint8Array`. A string is a file path, read from disk under Node.
- `deck.save()` resolves to a `Uint8Array`. Writing it to a file is up to you.
- A setter or a method changes the part in memory and marks it dirty. Reading a property never marks a part dirty.
- `save()` serializes each dirty part again and writes every other part with its original bytes.
- `deck.opc` is the package underneath, with its parts, content types and relationships. [Round-trip guarantee](../reference/round-trip.md#package-layer) covers it.
- Positions and sizes are in EMU, 914400 to the inch. [Layout units](../reference/layout-units.md) has the conversion helpers.

## Find a shape

| Call | Finds | Looks inside groups |
| --- | --- | --- |
| `slide.shapes` | every top-level shape, in document order | no |
| [`slide.shapeByName(name)`](../reference/api/read/classes/Slide.md#shapebyname) | the first top-level shape with that name (`p:cNvPr/@name`) | no |
| [`slide.shapeById(id)`](../reference/api/read/classes/Slide.md#shapebyid) | the first top-level shape with that drawing id (`p:cNvPr/@id`) | no |
| [`slide.shapeByIdDeep(id)`](../reference/api/read/classes/Slide.md#shapebyiddeep) | the shape with that drawing id, visiting each group before its children | yes |
| [`slide.placeholder(type, idx?)`](../reference/api/read/classes/Slide.md#placeholder) | the first top-level placeholder with that `p:ph/@type`, and with that `idx` when you pass one | no |

```ts
const slide = deck.slides[0]
if (!slide) throw new Error('the deck has no slides')

const caption = slide.shapeByName('Caption')
const insideGroup = slide.shapeByIdDeep(12)
const body = slide.placeholder('body', '1')
console.log(caption?.placeholder, insideGroup?.shapeType, body?.text)
```

- The finders return `undefined` when nothing matches.
- `placeholder()` returns an `AutoShape`, since only `p:sp` shapes can be placeholders.
- `shape.placeholder` is `{ type, idx }`, or `null` when the shape is not a placeholder. A `p:ph` with no `idx` reads as `idx: '0'`.
- To find a shape by name inside a group, walk `group.shapes`.
- `isAutoShape`, `isPicture`, `isGraphicFrame`, `isGroupShape` and `isConnector` narrow a shape to its class. `shape.shapeType` names the class as a string.
- Tables have their own editing members. See [Reading it back](../tables.md#reading-it-back).

## Replace text

Three setters replace text at three levels:

| Setter | Replaces | Keeps |
| --- | --- | --- |
| [`shape.text`](../reference/api/read/classes/Shape.md#text), [`textFrame.text`](../reference/api/read/classes/TextFrame.md#text) | every paragraph, with one paragraph holding one run | the first paragraph's properties, and its first run's character formatting |
| [`paragraph.text`](../reference/api/read/classes/Paragraph.md#text) | the runs, line breaks and fields of that paragraph, with one run | the paragraph's level, alignment and bullet, its first run's formatting, and every other paragraph |
| [`run.text`](../reference/api/read/classes/Run.md#text) | the run's text | the run's formatting |

```ts
const title = slide.placeholder('title')
if (title) title.text = 'New title'

const second = slide.placeholder('body', '1')?.textFrame?.paragraphs[1]
if (second) second.text = 'Second bullet, rewritten'
```

- The string goes into a single run. A `\n` in it does not start a new paragraph.
- Leading and trailing spaces are kept (`xml:space="preserve"`).
- Setting `text` on a shape with no text frame, such as a picture, throws `UnsupportedFeatureError` with code `shape/no-text-frame`.
- To change one run and leave its neighbours alone, set `run.text` through `textFrame.paragraphs[i].runs[j]`.

## Change position, size and colour

```ts
const callout = slide.shapeByName('Callout')
if (callout) {
  callout.left = 914400
  callout.top = 457200
  callout.width = 3657600
  callout.fillColor = '1F4E79'
  callout.lineColor = 'D4D4D4'
}

const lead = slide.placeholder('body', '1')?.textFrame?.paragraphs[0]?.runs[0]
if (lead) {
  lead.bold = true
  lead.fontSizePt = 28
  lead.color = '1F4E79'
}
```

- `left`, `top`, `width` and `height` are in EMU.
- `fillColor` and `lineColor` take a six-digit hex colour, and `fillSchemeColor` and `lineSchemeColor` take a theme colour name such as `accent2`.
- `fillColor = null` removes the shape's own fill, so it inherits one from its style or placeholder. `noFill()` writes an explicit empty fill instead.
- A run takes `bold`, `italic`, `underline`, `fontSizePt`, `fontName`, `color` and `schemeColor`.
- [`Shape`](../reference/api/read/classes/Shape.md) and [`Run`](../reference/api/read/classes/Run.md) in the API reference list every setter.

## Add and remove shapes

```ts
const box = slide.addTextBox({
  text: 'Draft',
  left: 914400,
  top: 457200,
  width: 4572000,
  height: 914400,
  name: 'Draft stamp',
})
const run = box.textFrame?.paragraphs[0]?.runs[0]
if (run) run.bold = true

slide.addPicture(await readFile('logo.png'), { left: 8229600, top: 457200, width: 914400, height: 914400 })

slide.shapeByName('Old caption')?.delete()
```

- [`addTextBox`](../reference/api/read/classes/Slide.md#addtextbox) adds a rectangle text box (`txBox="1"`) with one paragraph. `left`, `top`, `width` and `height` are required, in EMU. `text` defaults to an empty paragraph, and `name` to `TextBox <id>`.
- [`addPicture`](../reference/api/read/classes/Slide.md#addpicture) stores the image as a part under `/ppt/media/`, registers its content type, and adds an image relationship from the slide. `name` defaults to `Picture <id>`.
- `addPicture` recognises PNG, JPEG, GIF, BMP, TIFF and WebP from the bytes. For any other format, pass `extension` and `contentType`.
- A new shape gets a drawing id one above the highest on the slide, and goes in front of the other shapes.
- `shape.delete()` removes the shape from its slide or group, and the object is unusable afterwards. A deleted picture's relationship and media part stay in the package.
- The animations of a deleted shape, or of a shape inside a deleted group, are removed with it. A connector attached to one keeps its line, and that end is left unattached.

## Read and write speaker notes

A slide's notes live in a separate notes slide part. Three getters read it:

| Getter | Returns | Keeps formatting | `null` when |
| --- | --- | --- | --- |
| [`slide.notesText`](../reference/api/read/classes/Slide.md#notestext) | the notes body as a string, paragraphs joined by `\n` | no | the slide has no notes part |
| [`slide.notesTextFrame`](../reference/api/read/classes/Slide.md#notestextframe) | the notes body as a `TextFrame`, with its paragraphs, runs and hyperlinks | yes | the slide has no notes part, or its notes part has no body text frame |
| [`slide.notesSlide`](../reference/api/read/classes/Slide.md#notesslide) | the whole notes slide as a [`NotesSlide`](../reference/api/read/classes/NotesSlide.md), with its slide image, body and slide number placeholders | yes | the slide has no notes part |

- A notes part with an empty body gives `notesText === ''`. So does a notes part with no body text frame, for which `notesTextFrame` is `null`.
- Every slide `TsPptx` writes has a notes part, so its `notesText` is `''` rather than `null` when it has no notes. A slide imported without `importNotes` has no notes part.
- A notes run reports `resolvedColor`, `resolvedSizePt`, `resolvedFontFace` and `resolvedBold` through the deck's notes master and its theme.
- A notes hyperlink resolves `run.hyperlink.url` through the notes part's own relationships.

[`slide.addNotes(text)`](../reference/api/read/classes/Slide.md#addnotes) writes notes and returns the `NotesSlide`:

```ts
if (slide.notesText === null) slide.addNotes('Open with the headline number.\nThen the three drivers.')

const first = slide.notesTextFrame?.paragraphs[0]?.runs[0]
if (first) first.bold = true
```

- A `\n` starts a new paragraph. The runs have no formatting of their own, so style them afterwards through `notesTextFrame`.
- On a slide that has notes, `addNotes` replaces the body's paragraphs. The part's geometry and its other two placeholders stay as they were.
- On a slide with no notes part, it creates one. The slide gets a relationship to it, and the notes part gets a relationship to the notes master (`rId1`) and one back to the slide (`rId2`).

### One notes master per deck

A deck holds at most one notes master (`p:notesMasterIdLst` allows `0..1`). Every way of adding notes follows the same rule:

| Call | The deck has a notes master | The deck has none |
| --- | --- | --- |
| `slide.addNotes` | the new notes use it | a notes master is added, bound to a copy of this deck's theme |
| `importSlide`, `importSlides` with `importNotes` | the imported notes use it, and the source deck's notes master and theme are not copied | the source deck's notes master and its theme are copied, in a batch from the first request that carries notes |
| `appendSlides` | the appended notes use it | the generator's notes master and its theme are added |

A deck's own notes styling therefore wins whenever it has one, and no mix of these calls adds a second notes master.

## Replace the image of a picture

```ts
import { isPicture } from 'pptx-ts/read'

const picture = slide.shapes.find(isPicture)
picture?.setImage(await readFile('our-logo.png'), { contentType: 'image/png', fit: 'cover' })
```

[`picture.setImage(bytes, options)`](../reference/api/read/classes/Picture.md#setimage) stores the bytes as a new media part, adds a relationship from the slide, and points the picture's `a:blip/@r:embed` at it.

- `contentType` is required, such as `'image/png'`. The bytes are not inspected. `extension` defaults from the content type.
- The old media part is never changed or removed. Several pictures often share one media part, after an import for example, and each of the others keeps its image. The old part stays in the package even when no picture uses it any more.
- `fit` sets the crop (`a:srcRect`) against the picture's current frame:

| `fit` | Crop |
| --- | --- |
| omitted | unchanged, and so is the frame. A crop sized for the old image's aspect ratio stays and can distort the new image. |
| `'cover'` | fills the frame, cropping the axis that overflows |
| `'contain'` | fits the whole image inside the frame, leaving space at both ends of the short axis |
| `'stretch'` | removes the crop, so the whole image stretches to the frame |

- `'cover'` and `'contain'` need a frame with a width and height above zero, and read the image's size from its bytes. When they cannot measure it, the crop stays as it was and a warning says so.
- To point a picture at an image its slide already references, assign the relationship id, as in `picture.imageRelId = other.imageRelId`. No part is added, and the id must name an image relationship of that slide.

## Edit the XML directly

For anything the typed setters do not cover, change the element and mark its part dirty:

```ts
const watermark = slide.shapeByName('Watermark')
if (watermark) {
  watermark.element_.getElementsByTagName('p:cNvPr')[0]?.setAttribute('hidden', '1')
  watermark.markDirty()
}
```

- `element_` returns the live element behind a model object, and `markDirty()` marks the part that holds it.
- Both are on `Slide`, `SlideLayout`, `SlideMaster`, `Placeholder`, `NotesPlaceholder`, `Theme`, every shape class, `TextFrame`, `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `Chart`, `ChartAxis`, `ChartSeries`, `ChartEx`, `ChartExAxis`, `ChartExSeries`, `Diagram`, `DiagramNode`, `DiagramPoint`, and the `ResolvedTableStyle` that `table.resolvedStyle` returns.
- On a shape, text frame, paragraph or run, `markDirty()` marks the part that holds the shape: a slide, a layout or a master. On a chart object it marks the chart part, and on a resolved table style the deck's table styles part.
- Without `markDirty()`, `save()` writes the part's original bytes and the change is lost. Nothing throws or warns.
- `deck.opc.part(partName)` reaches a part that no model class covers. Change its `dom`, then call `part.markDirty()`.
- The trailing underscore makes every use of the hatch easy to find in your own code.

```ts
const part = deck.opc.part('/ppt/slides/slide1.xml')
const text = part?.dom.getElementsByTagName('a:t')[0]
if (part && text) {
  text.textContent = 'New title'
  part.markDirty()
}
```

## Invalid input

| Condition | Result | Code |
| --- | --- | --- |
| the `load` input is not a zip archive | throws `PackageReadError` | `zip/not-a-zip-archive` |
| the `load` input is none of the accepted types | throws `InvalidOptionError` | `zip/unsupported-input` |
| the archive has no `[Content_Types].xml` | throws `PackageReadError` | `package/not-an-opc-package` |
| a part has no content type | throws `PackageReadError` | `package/part-content-type-missing` |
| `text` set on a shape with no text frame | throws `UnsupportedFeatureError` | `shape/no-text-frame` |
| an `addTextBox` or `addPicture` position or size that is not a finite number | throws `InvalidOptionError` | `coord/non-finite` |
| an `addTextBox` or `addPicture` width or height of zero or less | throws `InvalidOptionError` | `coord/not-positive` |
| `addPicture` bytes in an unrecognised format, without `extension` and `contentType` | throws `InvalidOptionError` | `image/undeterminable-type` |
| `setImage` without `contentType` | throws `InvalidOptionError` | `image/missing-content-type` |
| `setImage` with `fit: 'cover'` or `'contain'` on a picture with no frame, or a zero-size frame | throws `InvalidOptionError`, picture unchanged | `image/fit-needs-extent` |
| `setImage` with `fit: 'cover'` or `'contain'` when the image size cannot be read | warns, crop unchanged | `image/unmeasurable-natural-size` |
| `addNotes` must add a notes master, and no theme is reachable from the slide | throws `PackageReadError`, no part added | `package/part-missing` |
| `addNotes` on a slide whose notes part has no body placeholder | throws `PackageReadError` | `package/part-has-no-root` |

## Limits

- Only `shapeByIdDeep` looks inside groups.
- The text setters write a single run, and a `\n` stays inside it.
- `addTextBox` adds a plain rectangle text box and `addPicture` a plain picture. Other shapes need an import or the XML.
- `shape.delete()` leaves the shape's relationships and media in the package, and `setImage` leaves the image it replaced.
- `addNotes` writes runs with no formatting.
- A change made through `element_` or `part.dom` is saved only after `markDirty()`.

## See also

- [Copy slides between decks](copy-between-decks.md)
- [Build on a template](build-on-a-template.md)
- [Tables](../tables.md)
- [Round-trip guarantee](../reference/round-trip.md)
- [Read object model](../reference/read-object-model.md)
- [Layout units](../reference/layout-units.md)
- [Errors and warnings](../errors-and-warnings.md)
- API reference: [`Presentation`](../reference/api/read/classes/Presentation.md), [`Slide`](../reference/api/read/classes/Slide.md), [`Shape`](../reference/api/read/classes/Shape.md), [`TextFrame`](../reference/api/read/classes/TextFrame.md), [`Paragraph`](../reference/api/read/classes/Paragraph.md), [`Run`](../reference/api/read/classes/Run.md), [`Picture`](../reference/api/read/classes/Picture.md), [`NotesSlide`](../reference/api/read/classes/NotesSlide.md), [`AddTextBoxOptions`](../reference/api/read/interfaces/AddTextBoxOptions.md), [`AddPictureOptions`](../reference/api/read/interfaces/AddPictureOptions.md)
