---
doc-schema-version: 1
title: "OLE Embedded Objects"
summary: "Embed a live Office document (workbook, document, deck) or any OLE payload into a slide with addOleObject(), so double-clicking it in PowerPoint opens the source in place."
read_when:
  - Embedding a spreadsheet, Word document, or other file so it opens on double-click
  - Reproducing PowerPoint's Insert > Object > Create from File
  - Choosing the preview picture a slide shows for an embedded object
  - Understanding which relationship type and content type an embedded payload gets
doc_type: "guide"
---

# OLE Embedded Objects

`slide.addOleObject()` embeds a file inside the `.pptx` and places it on the slide
as a live OLE object — PowerPoint's **Insert ▸ Object ▸ Create from File**.
Double-clicking it opens the source in place (Excel's grid, Word's page) rather
than launching a separate window on a separate file.

```js
const s = pptx.addSlide()
s.addOleObject({
	path: 'assets/quarterly-budget.xlsx',
	cover: { path: 'assets/quarterly-budget.png' },
	x: 1,
	y: 1,
	w: 6,
	h: 3,
})
```

The payload's bytes travel inside the package, so the deck stays self-contained.
Either `data` (base64, with or without a `data:...;base64,` header) or `path` is
required; everything else is optional.

> Linked objects — a `<p:link>` pointing at a file outside the package — are not
> supported. Only embedded payloads.

## The preview picture (`cover`)

An embedded object shows a *cached picture* of its source. This library is
Node-first: it never opens the payload, so it cannot render one. Supply your own
with `cover`, taking a `path` or base64 `data` just like `addImage`:

```js
s.addOleObject({ path: 'model.xlsx', cover: { path: 'model-preview.png' } })
```

Omit it and a neutral gray placeholder is embedded instead. That is usually
survivable — PowerPoint reads the `mc:Choice` branch, which carries no picture at
all, and draws the live object over it — but **every other consumer** (and
PowerPoint's own `mc:Fallback` path) shows exactly the placeholder. Ship a real
screenshot for any deck meant to read correctly outside PowerPoint.

## Payload kind: `extn` and `progId`

The payload's extension picks three things at once: the part's content type, its
relationship type, and the default `progId` (the OLE server PowerPoint launches).
It is resolved from, in order: an explicit `extn`, a `data:` URI's MIME type, the
`path`'s extension, and finally `progId`.

| Resolved extension | `progId` default | Part |
| --- | --- | --- |
| `xlsx` / `xlsm` | `Excel.Sheet.12` / `Excel.SheetMacroEnabled.12` | `ppt/embeddings/…​.xlsx` |
| `docx` / `docm` | `Word.Document.12` / `Word.DocumentMacroEnabled.12` | `ppt/embeddings/…​.docx` |
| `pptx` / `pptm` | `PowerPoint.Show.12` / `PowerPoint.ShowMacroEnabled.12` | `ppt/embeddings/…​.pptx` |
| anything else | `Package` | `ppt/embeddings/…​.bin` |

Those six are OPC packages (Office files are themselves ZIP archives) and keep
their own extension. Everything else is embedded as a generic OLE-server blob in
a `.bin` part — the same thing PowerPoint writes for a shell-packaged payload, so
an unrecognized extension never leaks into `[Content_Types].xml`.

Set `progId` explicitly to override the default, e.g. when passing raw base64
with no extension to infer from:

```js
s.addOleObject({ data: rawBase64Xlsx, progId: 'Excel.Sheet.12' })
```

## Sizing

`w`/`h` default to **4 × 3 inches** — the library cannot measure a document it
does not open, so there is no natural size to fall back on. Set them explicitly.

`imgW`/`imgH` are the preview's *native* size in EMU, which PowerPoint uses to
keep the object's aspect ratio when it re-renders the embedded document. They
default to the object's own `w`/`h` converted to EMU; override them only when the
cover image's true pixel dimensions matter.

`showAsIcon: true` sets `p:oleObj@showAsIcon`, PowerPoint's "Display as icon"
checkbox. The `cover` image is still what gets drawn, so supply an icon-looking
preview to match.

## Emitted OOXML

Each object is a `<p:graphicFrame>` whose `<a:graphicData>` carries the
`…/presentationml/2006/ole` URI and an `<mc:AlternateContent>`:

```xml
<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
  <mc:AlternateContent xmlns:mc="…/markup-compatibility/2006">
    <mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">
      <p:oleObj name="Worksheet" r:id="rId1" imgW="5486400" imgH="2743200" progId="Excel.Sheet.12">
        <p:embed/>
      </p:oleObj>
    </mc:Choice>
    <mc:Fallback>
      <p:oleObj name="Worksheet" r:id="rId1" imgW="5486400" imgH="2743200" progId="Excel.Sheet.12">
        <p:embed/>
        <p:pic><!-- the cached cover picture --></p:pic>
      </p:oleObj>
    </mc:Fallback>
  </mc:AlternateContent>
</a:graphicData>
```

This mirrors what PowerPoint itself authors. Two details are worth knowing:

- The `mc:Choice` **declares** the VML namespace (`Requires="v"`) but contains no
  VML. Modern PowerPoint writes no `spid` attribute and no `vmlDrawing` part, so
  neither is emitted here.
- Only the `mc:Fallback` carries the cached `<p:pic>`. That is why the `cover`
  matters most for non-PowerPoint consumers.

Package-wise each object consumes two relationships and two parts:

| Part | Relationship type |
| --- | --- |
| `ppt/embeddings/oleObject-{slide}-{n}.{extn}` | `…/relationships/package` (Office file) or `…/relationships/oleObject` (`.bin`) |
| `ppt/media/image-{slide}-{n}.{extn}` | `…/relationships/image` |

Payload parts are never shared between objects, even when two objects carry
byte-identical bytes: PowerPoint gives each embedded object its own part, and
collapsing them would make editing either one rewrite the other's source. (Cover
images *are* deduplicated, exactly like any other image.)

## Verifying

A `<p:oleObj>` that PowerPoint dislikes is not reported as a corrupt file — the
graphicFrame is silently dropped and the slide simply has no shape where the
object was. Schema validation cannot see that. `pnpm run test:com` (Windows +
PowerPoint) opens a generated OLE deck and reads each shape's `OLEFormat.ProgID`
back out, which is the check that actually catches it.

One trap if you write your own COM check: a *windowless* PowerPoint does not
instantiate embedded objects, so `Shapes` comes back without them — even for a
deck PowerPoint authored itself. Open with a window when reading OLE objects back.
