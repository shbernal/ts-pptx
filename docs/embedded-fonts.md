---
doc-schema-version: 1
title: "Embedded fonts"
summary: "Embed font files in a deck with embedFont(), carry a source deck's embedded fonts through importSlide(), importSlides(), importSlideMasters() and appendSlides(), name the typeface PowerPoint matches against, and read the embedded faces back."
read_when:
  - Embedding a font file so a deck renders on machines that do not have the font
  - Embedding the bold, italic or bold italic face of a family
  - Carrying embedded fonts through importSlide, importSlides, importSlideMasters or appendSlides
  - Working out why PowerPoint does not use an embedded font
  - Choosing between embedFont and registerFontMetrics
doc_type: "guide"
---

# Embedded fonts

`pptx.embedFont(options)` writes a font file into the deck, so PowerPoint can render text in that font on a machine that does not have it installed.

```ts
import { TsPptx } from 'pptx-ts'

const pptx = new TsPptx()
await pptx.embedFont({ path: 'fonts/Silkscreen-Regular.ttf', typeface: 'Silkscreen' })
pptx.addSlide().addText('Hello', { x: 1, y: 1, w: 6, h: 1, fontFace: 'Silkscreen' })
await pptx.writeFile({ fileName: 'embedded.pptx' })
```

`embedFont` returns a promise, so await it before writing the deck. Fonts belong to the deck, not to a slide, so the call can come before or after the slides that use the font.

## Options at a glance

`pptx.embedFont(options)`:

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `path` | `string` | none | The font file to read: a file path under Node, or an `http` or `https` URL. Wins over `data` when both are set. |
| `data` | `Uint8Array \| ArrayBuffer \| string` | none | The font bytes. A string is base64, bare or as a `data:` URL, and never a path. |
| `typeface` | `string` | required | The family name written to the deck. It must be the `fontFace` your text uses. |
| `style` | `'regular' \| 'bold' \| 'italic' \| 'boldItalic'` | `'regular'` | Which face of the family the bytes are. |

The import methods of `pptx-ts/read` take one option each:

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `ImportSlideOptions.embedFonts` | `boolean` | `false` | `importSlide` copies every font the source deck embeds. |
| `ImportSlidesRequest.embedFonts` | `boolean` | `false` | `importSlides` copies every font of that request's source deck, once per source. |
| `ImportSlideMastersOptions.embedFonts` | `boolean` | `false` | `importSlideMasters` copies every font the source deck embeds. |

`appendSlides` has no option: it always carries the fonts the generator embedded.

## Embed a font file

Pass the font as a file, a URL or bytes:

```ts
import { readFile } from 'node:fs/promises'

await pptx.embedFont({ path: 'fonts/Silkscreen-Regular.ttf', typeface: 'Silkscreen' })
await pptx.embedFont({ path: 'https://example.com/fonts/Inter-Regular.ttf', typeface: 'Inter' })
await pptx.embedFont({ data: await readFile('fonts/Lora-Regular.ttf'), typeface: 'Lora' })
await pptx.embedFont({ data: loraBase64, typeface: 'Lora', style: 'italic' })
```

- The Node entry point reads a path from disk and fetches an `http` or `https` URL. The browser and default entry points fetch every `path`.
- A `data` string is decoded as base64, with or without a `data:font/ttf;base64,` prefix.
- The file goes into the deck whole, as `ppt/fonts/fontN.fntdata`. No glyphs are removed.
- The deck is marked as carrying whole faces: `embedTrueTypeFonts="1"` and `saveSubsetFonts="0"` on `presentation.xml`. A deck with no embedded font keeps `saveSubsetFonts="1"` and no `embedTrueTypeFonts`.

## Embed bold and italic faces

A family has up to four faces. Embed each file under the same `typeface` with its own `style`:

```ts
await pptx.embedFont({ path: 'fonts/Silkscreen-Regular.ttf', typeface: 'Silkscreen' })
await pptx.embedFont({ path: 'fonts/Silkscreen-Bold.ttf', typeface: 'Silkscreen', style: 'bold' })
pptx.addSlide().addText([
  { text: 'Regular ', options: { fontFace: 'Silkscreen' } },
  { text: 'bold', options: { fontFace: 'Silkscreen', bold: true } },
], { x: 1, y: 1, w: 6, h: 1 })
```

- Calls with the same `typeface` share one entry in the deck's font list.
- The faces are written in the order `regular`, `bold`, `italic`, `boldItalic`, whatever order the calls came in.
- A second call with the same `typeface` and `style` replaces the bytes of the first.
- A different `typeface` makes a separate entry, even for files of the same family.

## Name the typeface your text uses

PowerPoint matches an embedded font to text by name. The `typeface` you pass is written to the deck's font list, and `fontFace` is written to each run, so the two strings must be equal:

```ts
await pptx.embedFont({ path: 'fonts/Silkscreen-Regular.ttf', typeface: 'Silkscreen' })
pptx.addSlide().addText('matches', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen' })
pptx.addSlide().addText('does not match', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen Regular' })
```

- PowerPoint does not use the embedded font for text whose `fontFace` differs from `typeface`.
- The library does not read the font file, so it cannot check the name, and a mismatch raises no warning. Use the family name the font file declares.

## Carry fonts when importing slides

Fonts are stored on the presentation, not on a slide, so a slide copied from another deck arrives without them. Ask for them with `embedFonts`:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { Presentation } from 'pptx-ts/read'

const deck = await Presentation.load(await readFile('deck.pptx'))
const library = await Presentation.load(await readFile('library.pptx'))

deck.importSlide(library, 0, { embedFonts: true })
deck.importSlides([
  { source: library, sourceIndex: 3, outputIndex: 1, embedFonts: true },
  { source: library, sourceIndex: 4, outputIndex: 2 },
])
deck.importSlideMasters(library, { embedFonts: true })
await writeFile('deck-merged.pptx', await deck.save())
```

- The source deck's whole font list comes across. The list does not record which slide uses which face, so the copy cannot be narrowed to one slide.
- In `importSlides`, one request that asks carries its source's fonts once for the whole batch, however many of that source's pages the batch names.
- The merge goes by `typeface` and face. A face the deck already embeds is kept, and the incoming copy of that face is not added. Importing the same slide twice adds each face once.
- When the deck already has an entry for the `typeface`, that entry keeps its own attributes, and only the missing faces are added to it.
- The import checks that every font file the source lists is in the source package before it changes anything. A missing one throws, and the deck is left byte-identical.
- `appendSlides` carries the fonts of the `TsPptx` you pass it, with the same merge rules.

## Register font metrics separately for text fit

`embedFont` and `registerFontMetrics` both take a font file and a family name, and they do different jobs. Neither one does the other's:

| | `embedFont` | `registerFontMetrics` |
| --- | --- | --- |
| Result | The font file is stored in the deck. | The library measures text in that font. Nothing is stored in the deck. |
| Used by | PowerPoint, when it opens the deck | `fit: 'shrink'` and `measureText`, when the deck is built |
| Arguments | `{ path \| data, typeface, style }` | `(face, source, { bold, italic, font })` |
| A string source | `path` is a path or URL. A `data` string is base64. | Always a path or URL. |
| Weight and style | `style` names the face | `bold` and `italic` flags |

A deck that fits text in an embedded font passes the file to both:

```ts
const bytes = await readFile('fonts/Silkscreen-Regular.ttf')
await pptx.embedFont({ data: bytes, typeface: 'Silkscreen' })
await pptx.registerFontMetrics('Silkscreen', bytes)
```

[Text that fits](text-fit.md) covers `registerFontMetrics`.

## Invalid input

| Condition | Result | Code |
| --- | --- | --- |
| `typeface` missing, not a string, or only spaces | throws `InvalidOptionError` | `font/missing-typeface` |
| `style` outside the four faces | throws `InvalidOptionError` | `font/invalid-style-slot` |
| neither `path` nor `data` | throws `InvalidOptionError` | `font/missing-source` |
| `path` not a string, and `data` not bytes or a string | throws `InvalidOptionError` | `font/missing-source` |
| a `data` string that is not base64, including `''` | throws `InvalidOptionError` | `font/invalid-base64` |
| a `path` file that cannot be read, under Node | throws `MediaError` | `font/read-failed` |
| a `path` URL that answers with an error status | throws `MediaError` | `font/fetch-failed` |
| a font file the source deck lists but does not contain, on an import with `embedFonts` | throws `PackageReadError`, deck unchanged | `package/part-missing` |

## Limits

- Faces are embedded whole. There is no subsetting.
- The bytes are not checked. An empty array or a file that is not a font is embedded as given.
- The font's embedding permissions (`OS/2.fsType`) are not read. Licensing is the caller's responsibility.
- The typeface name and the style are not read from the file. A wrong `typeface` or `style` is not detected.
- An import copies the source deck's whole font list, not only the faces its slides use.
- `importShape` and `importShapes` do not carry fonts.
- An import or `appendSlides` adds fonts without setting `embedTrueTypeFonts` on the destination deck. A deck built with `embedFont` has it set, but a template that never embedded a font does not.
- An open deck gains fonts only through the imports and `appendSlides`. No method adds a font file directly or removes one.

## Reading it back

`Presentation.embeddedFonts` lists the fonts any deck embeds, including one PowerPoint saved:

```ts
import { readFile } from 'node:fs/promises'
import { Presentation } from 'pptx-ts/read'

const deck = await Presentation.load(await readFile('deck.pptx'))
for (const font of deck.embeddedFonts) {
  for (const face of font.faces) {
    const bytes = deck.opc.part(face.partName)?.serialize()
    console.log(font.typeface, face.slot, face.partName, bytes?.length)
  }
}
```

- Each entry has `typeface`, `panose` (a string, or `null` when the deck declares none), and `faces`.
- Each face has `slot` and `partName`, the package path of its font file. Faces come in the order `regular`, `bold`, `italic`, `boldItalic`.
- The list is `[]` when the deck embeds no font. An entry with no `typeface`, and a face whose file reference does not resolve, are left out.
- A face PowerPoint saved as a subset holds only some of the font's glyphs, so its bytes can be smaller than the original font file.

[Read object model](reference/read-object-model.md#presentation) lists every member.

## See also

- [Text that fits](text-fit.md)
- [Read and edit a deck](reading/read-and-edit.md)
- [Errors and warnings](errors-and-warnings.md)
- API reference: [`TsPptx.embedFont`](reference/api/index/classes/TsPptx.md#embedfont), [`TsPptx.registerFontMetrics`](reference/api/index/classes/TsPptx.md#registerfontmetrics)
