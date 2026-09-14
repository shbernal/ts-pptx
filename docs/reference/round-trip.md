---
doc-schema-version: 1
title: "Round-trip guarantee"
summary: "What save() writes for each part of a deck opened with pptx-ts/read, the constructs the read model keeps without decoding, and the OPC package layer under Presentation."
read_when:
  - Checking whether loading and saving a deck changes a part
  - Checking whether the read model decodes a construct or only keeps its bytes
  - Working with OPC parts, content types or relationships directly
doc_type: "reference"
---

# Round-trip guarantee

`pptx-ts/read` keeps the bytes of every part it loads. `save()` writes those bytes back for each part nothing changed, so loading and saving a deck leaves its untouched parts byte-identical. An edit reserializes only the parts it touched.

This page states what `save()` writes for each part, which constructs survive without a decoder, and the package layer the object model sits on. [Read object model](read-object-model.md) describes the typed view. [Read and edit a deck](../reading/read-and-edit.md#load-and-save) covers opening and editing one.

## What save writes

| Part state | How a part gets there | What `save()` writes |
| --- | --- | --- |
| Untouched | Loaded and never marked dirty. Reading [`part.dom`](api/read/classes/Part.md#dom) parses the part without marking it. | The loaded bytes, byte-identical. |
| Dirty | [`part.markDirty()`](api/read/classes/Part.md#markdirty), called by you or by a typed setter. | The part's DOM, serialized. The XML is equivalent but not byte-identical: attribute quoting and whitespace can change. The part keeps its XML declaration, or gets a default one when it had none. A part marked dirty with no change is still reserialized. |
| Binary (image, font, media) | It cannot become dirty: `markDirty()` throws `part/not-xml`. | The loaded bytes. [`Picture.setImage()`](api/read/classes/Picture.md#setimage) adds a new part and leaves the old one in place. |
| Added | [`opc.addPart()`](api/read/classes/OpcPackage.md#addpart), called by you or by a method that adds media, notes or copied parts. | The given bytes, after the loaded parts, in the order they were added. `[Content_Types].xml` gains an `Override` unless an existing entry already resolves the part to its type. |
| Removed | [`opc.removePart()`](api/read/classes/OpcPackage.md#removepart), called by you or by [`removeSlide()`](api/read/classes/Presentation.md#removeslide). | Nothing. The part's `.rels` part and its `Override` entry go with it. A `Default` entry stays. |
| The `.rels` part of a changed relationship set | [`add()`](api/read/classes/Relationships.md#add), [`addWithId()`](api/read/classes/Relationships.md#addwithid) or [`remove()`](api/read/classes/Relationships.md#remove) on the set. | The set, serialized. An existing `.rels` part keeps its place in the package. A new one is appended. |
| `[Content_Types].xml`, unchanged | No registration or removal touched it. | The loaded bytes. |
| `[Content_Types].xml`, changed | [`ensureRegistered()`](api/read/classes/ContentTypes.md#ensureregistered) or [`ensureDefault()`](api/read/classes/ContentTypes.md#ensuredefault) added an entry, or `removePart()` dropped an `Override`. | A regenerated part: every `Default` entry, then every `Override` entry. |
| Under `/[trash]/` | PowerPoint leaves deleted parts there, unregistered and unreferenced. | Nothing. `load()` drops them. |

- The guarantee covers part bodies, the set of part names and part order. Loaded parts keep their zip order.
- It does not cover the zip container. Compression and entry metadata can differ, so the file as a whole is not byte-identical.
- [`part.originalBytes`](api/read/classes/Part.md#originalbytes) keeps the loaded bytes for the part's whole life. [`part.serialize()`](api/read/classes/Part.md#serialize) returns the current body, edits included. `save()`, slide copies and imports all read `serialize()`.
- A typed setter marks its part dirty. An edit through `element_` or `part.dom` does not. Call `markDirty()` after such an edit, or `save()` writes the loaded bytes and the edit is lost.
- `test/read/roundtrip.test.js` checks these rules on every deck in the read fixture corpus. The [testing guide](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/testing.md) describes that suite.

`load()` rejects a package it cannot hold to these rules. A malformed `.rels` part throws the first time its relationships are read.

| Code | Thrown when |
| --- | --- |
| `package/not-an-opc-package` | The zip has no `[Content_Types].xml`. |
| `package/part-content-type-missing` | No `Override` or `Default` entry resolves a part's content type. |
| `package/content-types-invalid-root`, `package/content-types-entry-incomplete` | `[Content_Types].xml` has the wrong root, or an entry lacks an attribute. |
| `package/relationships-invalid-root`, `package/relationship-incomplete`, `package/relationship-invalid-target-mode`, `package/duplicate-relationship-id` | A `.rels` part has the wrong root, a relationship lacks `Id`, `Type` or `Target`, has an unknown `TargetMode`, or repeats an id. |

## Kept but not decoded

These constructs load and save byte-identical, and no getter decodes them. Their parts stay reachable through [`presentation.opc`](api/read/classes/Presentation.md#opc), and through the shape that holds them where there is one.

| Construct | Parts | Why it is not decoded | How to reach it |
| --- | --- | --- | --- |
| SmartArt layout, quick style and colours | The `layout`, `quickStyle` and `colors` parts under `ppt/diagrams/` | They are the presets PowerPoint's layout engine runs to draw a diagram. Decoding them means implementing that engine. | [`Diagram.layoutPart`](api/read/classes/Diagram.md#layoutpart), [`quickStylePart`](api/read/classes/Diagram.md#quickstylepart) and [`colorsPart`](api/read/classes/Diagram.md#colorspart). [`layoutTypeId`](api/read/classes/Diagram.md#layouttypeid) names the preset. The data model and drawing cache are decoded: see [SmartArt](read-object-model.md#smartart). |
| OLE objects | `p:oleObj` in a graphic frame, and the payload under `ppt/embeddings/` | The payload is another application's file format. | A [`GraphicFrame`](api/read/classes/GraphicFrame.md) whose [`graphicDataUri`](api/read/classes/GraphicFrame.md#graphicdatauri) is not a table, chart, chartEx or SmartArt namespace. Its parts through [`slide.relationships`](api/read/classes/Slide.md#relationships). See [OLE embedded objects](../ole-objects.md#reading-it-back). |
| 3D models | `am3d:model3d` in a graphic frame, and a `.glb` part under `ppt/media/` | A glTF binary plus camera and lighting markup, with no reader in the library. | A `GraphicFrame` whose `graphicDataUri` is `http://schemas.microsoft.com/office/drawing/2017/model3d`. See [3D models](../3d-models.md#reading-it-back). |
| Ink | `p:contentPart`, and an InkML part | Digitizer strokes, with no renderer and no writer. | `shapes` skips a bare `p:contentPart`. Inside `mc:AlternateContent`, `shapes` reports the shape in the `mc:Fallback` branch. The ink part through `slide.relationships`. |
| Audio and video | `a:audioFile` or `a:videoFile` and `p14:media` on a picture, and the media part under `ppt/media/` | The read model does not decode the media relationships. | The [`Picture`](api/read/classes/Picture.md) reports its poster image through [`imagePartName`](api/read/classes/Picture.md#imagepartname). The media parts through `slide.relationships`. |
| Shape 3D | `a:scene3d` and `a:sp3d` in a shape's `p:spPr` | No getter reads them. | [`shape.element_`](api/read/classes/Shape.md). A table cell's bevel and light rig are decoded, by [`TableCell.cell3D`](api/read/classes/TableCell.md#cell3d). |
| Animation timing | `p:timing` in a slide part | The library keeps the recursive timing tree as DOM, and rewrites only the shape ids it references when ids change. | [`slide.hasAnimations`](api/read/classes/Slide.md#hasanimations) and [`flattenAnimations()`](api/read/classes/Slide.md#flattenanimations). Transitions are decoded, by [`slide.transition`](api/read/classes/Slide.md#transition). See [Animations and transitions](../animations-and-transitions.md#reading-it-back). |
| Custom XML data | The `item` and `itemProps` parts under `customXml/` | Application-defined XML with no schema to decode against. | [`opc.parts`](api/read/classes/OpcPackage.md#parts). Programmatic tags are a different construct, and are decoded: see [Tags](read-object-model.md#tags). |
| Embedded font data | The `.fntdata` parts under `ppt/fonts/` | Glyph data. | [`presentation.embeddedFonts`](api/read/classes/Presentation.md#embeddedfonts) names each face's part, and [`opc.part()`](api/read/classes/OpcPackage.md#part) returns it. See [Embedded fonts](../embedded-fonts.md#reading-it-back). |
| Document statistics | `Slides`, `Words`, `Paragraphs`, `HeadingPairs` and the other counts in `docProps/app.xml` | The producing application computed them for the file it wrote. An edited deck no longer matches them. | `opc.part('/docProps/app.xml')`. [`appProperties`](api/read/classes/Presentation.md#appproperties) reports four other fields: see [Document properties](read-object-model.md#document-properties). |
| Chart workbooks and chartEx sidecars | The workbook embedded behind a chart, and a chartEx chart's style, colours and geography parts | Chart getters read the value caches stored in the chart part. | `presentation.opc.relationshipsFor(chart.partName)`. |

Comments in PowerPoint's 2018 format are decoded, read-only: see [Comments](read-object-model.md#comments).

## Package layer

[`Presentation`](api/read/classes/Presentation.md) wraps an [`OpcPackage`](api/read/classes/OpcPackage.md), reachable as `presentation.opc`. Work at this layer for a part the object model has no class for.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { OpcPackage } from 'pptx-ts/read'

const pkg = await OpcPackage.load(await readFile('deck.pptx'))
const slides = pkg.partsByContentType('application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
console.log(slides.map((part) => part.partName)) // ['/ppt/slides/slide1.xml', ...]
await writeFile('deck-copy.pptx', await pkg.save())
```

| Export | What it is for | Members |
| --- | --- | --- |
| [`OpcPackage`](api/read/classes/OpcPackage.md) | The loaded package: parts by name or content type, each part's relationships, adding and removing parts, saving. | [`load()`](api/read/classes/OpcPackage.md#load), [`parts`](api/read/classes/OpcPackage.md#parts), [`part()`](api/read/classes/OpcPackage.md#part), [`partsByContentType()`](api/read/classes/OpcPackage.md#partsbycontenttype), [`relationshipsFor()`](api/read/classes/OpcPackage.md#relationshipsfor), [`contentTypes`](api/read/classes/OpcPackage.md#contenttypes), [`addPart()`](api/read/classes/OpcPackage.md#addpart), [`removePart()`](api/read/classes/OpcPackage.md#removepart), [`reserveMediaPartName()`](api/read/classes/OpcPackage.md#reservemediapartname), [`reservePartNameLike()`](api/read/classes/OpcPackage.md#reservepartnamelike), [`save()`](api/read/classes/OpcPackage.md#save) |
| [`Part`](api/read/classes/Part.md) | One part: its loaded bytes, its DOM on demand, and whether it is dirty. | [`partName`](api/read/classes/Part.md#partname), [`contentType`](api/read/classes/Part.md#contenttype), [`originalBytes`](api/read/classes/Part.md#originalbytes), [`isXmlPart`](api/read/classes/Part.md#isxmlpart), [`dom`](api/read/classes/Part.md#dom), [`isParsed`](api/read/classes/Part.md#isparsed), [`markDirty()`](api/read/classes/Part.md#markdirty), [`isDirty`](api/read/classes/Part.md#isdirty), [`serialize()`](api/read/classes/Part.md#serialize) |
| [`ContentTypes`](api/read/classes/ContentTypes.md) | The overlay over `[Content_Types].xml` that resolves and registers part types. | [`contentTypeFor()`](api/read/classes/ContentTypes.md#contenttypefor), [`ensureRegistered()`](api/read/classes/ContentTypes.md#ensureregistered), [`ensureDefault()`](api/read/classes/ContentTypes.md#ensuredefault), [`removeOverride()`](api/read/classes/ContentTypes.md#removeoverride), [`isDirty`](api/read/classes/ContentTypes.md#isdirty), [`serialize()`](api/read/classes/ContentTypes.md#serialize) |
| [`Relationships`](api/read/classes/Relationships.md) | The overlay over one `.rels` part. It is iterable. | [`get()`](api/read/classes/Relationships.md#get), [`byType()`](api/read/classes/Relationships.md#bytype), [`resolveTarget()`](api/read/classes/Relationships.md#resolvetarget), [`add()`](api/read/classes/Relationships.md#add), [`addWithId()`](api/read/classes/Relationships.md#addwithid), [`remove()`](api/read/classes/Relationships.md#remove), [`size`](api/read/classes/Relationships.md#size), [`serialize()`](api/read/classes/Relationships.md#serialize) |
| [`resolveRelativePartName()`](api/read/functions/resolveRelativePartName.md), [`relsPartNameFor()`](api/read/functions/relsPartNameFor.md) | Partname arithmetic: a relationship target to an absolute partname, and a part to the partname of its `.rels` part. | |

- `load()` takes a `Uint8Array`, `ArrayBuffer`, `Blob` or `number[]`. Under Node it also takes a file path: a string is always a path, never zip content. See [`OpcInput`](api/read/type-aliases/OpcInput.md).
- `parts` leaves out `[Content_Types].xml`. Read it through `contentTypes`.
- `contentTypeFor()` looks for an `Override` with the exact partname first, then a `Default` for the lowercased extension.
- `dom` throws `part/not-xml` for a binary part. Its `Document` type is `@xmldom/xmldom`'s, not the DOM library TypeScript ships, and the two are not assignable to each other.
- `relationshipsFor()` returns one cached set per part, and an empty set for a part with no `.rels` part. Pass `'/'`, or nothing, for the package relationships in `/_rels/.rels`.
- Relationship ids are opaque and need not be contiguous. `add()` allocates `rId<n>`, where n is one more than the highest numeric id in the set.
- `resolveTarget()` throws `relationship/not-found` for an id the set lacks, and `relationship/external-has-no-partname` for an external target. `resolveRelativePartName()` throws `package/relationship-target-escapes-root` for a target above the package root.
- `partsByContentType()` returns a new array on every call.
- `removePart()` leaves every relationship that points at the part. `removeSlide()` also unlinks the slide from the presentation, and removes each part the slide referenced that no remaining part references. It never removes a layout, master or theme.
- `reserveMediaPartName()` and `reservePartNameLike()` return an unused partname one past the highest index in use, and create nothing.
