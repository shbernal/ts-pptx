---
doc-schema-version: 1
title: "OLE embedded objects"
summary: "Embed a workbook, document, deck or other OLE payload in a slide with addOleObject(): the cover picture, how the payload kind and progId are chosen, sizing, and reading an object back."
read_when:
  - Embedding a spreadsheet, Word document or other file so it opens on double-click
  - Reproducing PowerPoint's Insert > Object > Create from File
  - Choosing the cover picture a slide shows for an embedded object
  - Working out which progId, part or content type an embedded payload gets
doc_type: "guide"
---

# OLE embedded objects

`slide.addOleObject()` embeds a file in the `.pptx` and places it on the slide as an OLE object, the kind PowerPoint's Insert > Object > Create from File makes. Double-clicking it in PowerPoint opens the embedded file in the application its `progId` names.

```ts
import TsPptx from 'pptx-ts'

const pptx = new TsPptx()
pptx.addSlide().addOleObject({
  path: 'assets/quarterly-budget.xlsx',
  cover: { path: 'assets/quarterly-budget.png' },
  x: 1,
  y: 1,
  w: 6,
  h: 3,
})
await pptx.writeFile({ fileName: 'budget.pptx' })
```

`data` or `path` is required, and everything else is optional. The payload's bytes are stored in the package, so the deck needs no file beside it. `addOleObject` returns the slide, so calls chain.

## Options at a glance

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `data` | `string` | none | The payload as base64, with or without a `data:` header |
| `path` | `string` | none | A file path or URL, read when the deck is written |
| `cover` | `{ path?: string; data?: string }` | a gray placeholder | The picture drawn for the object |
| `extn` | `string` | from `data`, `path` or `progId` | The payload kind |
| `progId` | `string` | the kind's default | The OLE server PowerPoint launches |
| `showAsIcon` | `boolean` | `false` | Sets PowerPoint's "Display as icon" flag |
| `imgW`, `imgH` | `number` | the frame's `w` and `h`, in EMU | The object's picture size, in EMU |
| `x`, `y` | `Coord` | `0` | Top-left corner |
| `w`, `h` | `Coord` | `4`, `3` | Size |
| `objectName` | `string` | `Object 1`, `Object 2`, ... | Selection Pane name |
| `altText` | `string` | none | Alt text |
| `objectLock` | `ObjectLockProps` | `{ noChangeAspect: true }` | Lock flags on the object's frame |

## Supply a cover picture

The library never opens the payload, so it cannot draw a picture of it. Pass one as `cover`, by `path` or as base64 `data` with a header, the way `addImage` takes an image:

```ts
slide.addOleObject({ path: 'budget.xlsx', cover: { data: `image/png;base64,${pngBase64}` } })
```

- With no `cover`, a 32 by 32 gray PNG is embedded and nothing is reported.
- A `cover.data` with no base64 header warns `preview-image/missing-base64-header`, and the gray PNG is embedded instead.
- The cover goes into the object's `mc:Fallback` branch only. A reader that does not take the `mc:Choice` branch draws the cover in place of the object. [Supply a preview picture](3d-models.md#supply-a-preview-picture) on the 3D models page explains the two branches.
- A cover used by two objects is stored once, like any other image.

## Choose the payload kind

The payload kind sets the part's extension, its content type, its relationship type and the default `progId`. The library takes the kind from the first of these that names one of the six Office kinds in the table:

1. `extn`, with or without a leading dot, in any case. An `extn` outside the table makes the payload a generic blob, and the later sources are not read.
2. The MIME type of a `data` value that starts with `data:`.
3. The extension of `path`, ignoring any query string or fragment.
4. `progId`, when it is one of the six defaults in the table.

When none of them names an Office kind, the payload is a generic blob.

| Kind | `progId` default | MIME type read from a `data:` URI |
| --- | --- | --- |
| `xlsx` | `Excel.Sheet.12` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `xlsm` | `Excel.SheetMacroEnabled.12` | `application/vnd.ms-excel.sheet.macroEnabled.12` |
| `docx` | `Word.Document.12` | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` |
| `docm` | `Word.DocumentMacroEnabled.12` | `application/vnd.ms-word.document.macroEnabled.12` |
| `pptx` | `PowerPoint.Show.12` | `application/vnd.openxmlformats-officedocument.presentationml.presentation` |
| `pptm` | `PowerPoint.ShowMacroEnabled.12` | `application/vnd.ms-powerpoint.presentation.macroEnabled.12` |
| generic blob | `Package` | none |

```ts
// Bare base64 has no MIME type and no file name, so progId decides the kind.
slide.addOleObject({ data: workbookBase64, progId: 'Excel.Sheet.12' })

// A path with no extension needs extn.
slide.addOleObject({ path: 'exports/budget', extn: 'xlsx' })
```

- An Office kind keeps its extension, as in `ppt/embeddings/oleObject-1-1.xlsx`, with a content type for that extension.
- A generic blob is stored as `.bin` with the content type `application/vnd.openxmlformats-officedocument.oleObject`, so an unknown extension never reaches `[Content_Types].xml`.
- A MIME header with no `data:` in front, such as `application/vnd.ms-excel.sheet.macroEnabled.12;base64,`, is not read for the kind. The bytes still load.
- An explicit `progId` is written as given. It picks the kind only when no earlier source did.
- Every object gets its own payload part, even when two objects carry the same bytes, so editing one object never changes the other.

## Size the object

```ts
slide.addOleObject({
  path: 'budget.xlsx',
  cover: { path: 'excel-icon.png' },
  showAsIcon: true,
  x: 1,
  y: 1,
  w: 1,
  h: 1,
})
```

- `w` and `h` default to 4 by 3 inches, and `x` and `y` to 0. The library does not open the payload, so it has no size to measure. Set all four.
- `imgW` and `imgH` are written to `p:oleObj` as the size of the object's picture, in EMU. Each defaults to the frame's width or height in EMU. A fraction is rounded.
- `showAsIcon: true` does not change what is drawn. The cover is still the picture, so give an icon picture to match.
- The frame is locked against aspect-ratio changes. Flags in `objectLock` are added to that lock, and `noChangeAspect: false` removes it.

## Invalid input

Throws happen inside the `addOleObject()` call, as `InvalidOptionError`. A file that fails to load is reported when the deck is written.

| Condition | Warns or throws | Code |
| --- | --- | --- |
| neither `data` nor `path` is set | throws | `ole/missing-source` |
| `imgW` or `imgH` is not a number from 0 to 2147483647 | throws | `ole/invalid-image-size` |
| `x`, `y`, `w` or `h` is not a finite number | throws | `coord/non-finite` |
| `w` or `h` is 0 | warns, and the zero is kept | `frame/zero-extent` |
| `cover.data` has no base64 header | warns, and the gray placeholder is embedded | `preview-image/missing-base64-header` |
| `path` or `cover.path` fails to load (when written) | throws `MediaError` | `media/load-failed` |
| the same, with `onMediaError: 'placeholder'` | warns, and a broken-image PNG is written in place of the file | `media/load-failed` |
| `objectName` is only whitespace, longer than 255 characters, or holds control characters | warns | `object-name/empty`, `object-name/too-long`, `object-name/control-characters` |

[Errors and warnings](errors-and-warnings.md) covers the error classes and how to route warnings.

## Limits

- Linked objects, which point at a file outside the package, cannot be authored.
- The library never opens or checks the payload. It draws no cover, measures no size, and does not notice bytes that do not match the kind.
- A generic blob's bytes are written unchanged. The library does not package a file into an OLE compound file.
- Only a `data:` URI's MIME type is read, and only for the six Office kinds.
- Two objects never share a payload part.
- `pptx-ts/read` has no typed accessor for an OLE object.

## Reading it back

```ts
import { readFile } from 'node:fs/promises'
import { Presentation } from 'pptx-ts/read'

const deck = await Presentation.load(await readFile('budget.pptx'))
for (const slide of deck.slides) {
  for (const shape of slide.shapes) {
    if (shape.shapeType === 'graphicFrame') console.log(shape.name)
  }
}
```

- Before export, [`slide.objects`](groups.md#list-what-a-slide-holds) lists the object with `type: "oleObject"` and `canGroup: false`.
- `pptx-ts/read` loads the object as a `graphicFrame` shape whose `name` is its `objectName`. There is no accessor for the payload, the `progId` or the cover. See [the preserve-only boundary](reference/pptx-read.md#preserve-only-boundary-what-the-read-model-does-not-decode).
- Saving a loaded deck keeps the payload and cover parts, and `importSlide` copies both into the target deck.
- [PPTX inspection](reference/pptx-inspection.md) reports the object as a `graphicFrame` element with `graphicKind: 'other'`.

## See also

- [3D models](3d-models.md)
- [Groups](groups.md), for `slide.objects`
- [Reading and round-tripping existing decks](reference/pptx-read.md)
- [Errors and warnings](errors-and-warnings.md)
- API reference: [`OleObjectProps`](reference/api/index/type-aliases/OleObjectProps.md), [`ObjectLockProps`](reference/api/index/interfaces/ObjectLockProps.md), [`SlideObjectInfo`](reference/api/index/interfaces/SlideObjectInfo.md)
