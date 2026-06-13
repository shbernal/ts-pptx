---
doc-schema-version: 1
title: "PPTX Read / Round-Trip"
summary: "Open an existing .pptx, inspect its OPC parts, and save it back losslessly (foundation for editing)."
read_when:
  - Opening or editing decks this library did not generate
  - Round-tripping a .pptx with untouched parts byte-identical
  - Reading OPC parts, content types, or relationships
doc_type: "reference"
---

# Reading and round-tripping existing decks (`pptxgenjs/read`)

The `pptxgenjs/read` subpath opens an **existing** `.pptx` file, exposes its
OPC package structure, and saves it back losslessly. It is the foundation for
python-pptx-style editing of decks this library did not generate.

It is a separate subsystem from the generator (`pptxgenjs`) and the inspector
(`pptxgenjs/inspect`): those are one-way and lossy, while `read` keeps the
package's own XML as the source of truth.

Status: **Phase 2 — read object model**. On top of the Phase 1 OPC layer
(load, parts, content types, relationships, lossless save) there is now a
navigable, typed view of the deck: `Presentation → slides → shapes → text
frame → paragraphs → runs`, read from the live DOM. Mutation still happens
directly on a part's DOM plus `markDirty()`; typed *editing* APIs come in
Phase 3.

## Quick start

Read a deck through the typed object model:

```js
import { readFile, writeFile } from 'node:fs/promises'
import { Presentation } from '@shbernal/pptxgenjs/read'

const presentation = await Presentation.load(await readFile('deck.pptx'))

for (const slide of presentation.slides) {
	for (const shape of slide.shapes) {
		console.log(shape.shapeType, shape.name, shape.left, shape.top)
		if (shape.hasTextFrame) console.log(shape.text)
	}
}

// Save it back — untouched parts are byte-identical
await writeFile('deck-roundtrip.pptx', await presentation.save())
```

Or work at the OPC layer directly:

```js
import { OpcPackage } from '@shbernal/pptxgenjs/read'

const pkg = await OpcPackage.load(await readFile('deck.pptx'))
const slides = pkg.partsByContentType(
	'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
)
console.log(slides.map((part) => part.partName)) // ['/ppt/slides/slide1.xml', ...]
await writeFile('deck-roundtrip.pptx', await pkg.save())
```

`Presentation` wraps an `OpcPackage`; reach the lower layer any time via
`presentation.opc`.

The module is isomorphic: bytes in, bytes out, no `node:fs`. File I/O is the
caller's job, so it works in browsers too.

## Fidelity contract

- A `Part` keeps the **original bytes** from the zip for its whole life.
- Accessing `part.dom` parses lazily; parsing alone changes nothing.
- `save()` writes original bytes for every part that was never marked dirty —
  **untouched part bodies are byte-identical** to the input.
- Dirty parts (after `part.markDirty()`) are reserialized from their DOM:
  semantically equivalent and schema-valid, but not byte-identical (attribute
  quoting and whitespace may differ). The XML declaration is preserved.
- Whole-zip byte-identity is **not** promised: zip metadata and compression
  may differ. The contract covers part bodies, the part-name set, and part
  order.

This is verified by `test/read/roundtrip.test.js` against PowerPoint-authored
fixtures (see `test/read/fixtures/README.md`).

## API

### `OpcPackage`

```ts
type OpcInput = string | number[] | Uint8Array | ArrayBuffer | Blob

class OpcPackage {
	static load(input: OpcInput): Promise<OpcPackage>

	/** All parts keyed by partname (e.g. '/ppt/slides/slide1.xml'), in zip order. */
	readonly parts: ReadonlyMap<string, Part>
	/** Content-type resolution overlay over [Content_Types].xml (read-only). */
	readonly contentTypes: ContentTypes

	part(partName: string): Part | undefined
	partsByContentType(contentType: string): Part[]
	/** Relationships owned by a part; '/' (default) = package-level /_rels/.rels. */
	relationshipsFor(sourcePartName?: string): Relationships

	save(): Promise<Uint8Array>
}
```

`load()` rejects when the input is not an OPC package or when a part has no
resolvable content type (no `Override`, no `Default`) — the error names the
offending part.

`[Content_Types].xml` is not enumerated in `parts`; it is managed by the
package and exposed through the `contentTypes` overlay.

### `Part`

```ts
class Part {
	readonly partName: string
	readonly contentType: string

	/** Original bytes from the package. Do not mutate. */
	readonly bytes: Uint8Array
	/** Whether the body is XML (by content type). */
	readonly isXmlPart: boolean
	/** True once the body has been materialized as a DOM. */
	readonly isParsed: boolean
	/** Lazily parsed DOM (throws for binary parts such as images). */
	readonly dom: Document

	/** Call after mutating the DOM so save() reserializes this part. */
	markDirty(): void
	readonly isDirty: boolean

	/** Original bytes when clean; serialized DOM when dirty. */
	serialize(): Uint8Array
}
```

The `Document` type is `@xmldom/xmldom`'s, not lib.dom's — they are not
assignable to each other.

### `ContentTypes`

Read-only overlay over `[Content_Types].xml`.

```ts
class ContentTypes {
	static parse(xml: string): ContentTypes
	/** Exact Override match first, else Default by lowercased extension. */
	contentTypeFor(partName: string): string | undefined
	serialize(): string
}
```

### `Relationships`

Read-only overlay over one `.rels` part. Iterable.

```ts
interface Relationship {
	id: string // 'rId1'
	type: string // relationship type URI
	target: string // as written: relative or absolute
	targetMode?: 'Internal' | 'External'
}

class Relationships {
	static parse(xml: string, sourcePartName: string): Relationships
	readonly sourcePartName: string
	readonly size: number
	get(id: string): Relationship | undefined
	byType(type: string): Relationship[]
	/** Absolute partname for an internal rel; throws for External rels. */
	resolveTarget(id: string): string
	serialize(): string
}
```

Relationship ids are opaque: numbering is not necessarily contiguous or
ordered.

### Partname helpers

```ts
/** OPC pack-URI resolution: relative target + owning part → absolute partname. */
function resolveRelativePartName(sourcePartName: string, target: string): string
/** '.rels' partname for a part; '/' → '/_rels/.rels'. */
function relsPartNameFor(sourcePartName: string): string
```

## Read object model (Phase 2)

A navigable, typed view over the live DOM. Every proxy reads from its DOM
element on each access (no caching) and wraps the very nodes a later phase will
mutate. Geometry is reported in **EMU** (the OOXML unit; 914 400 per inch) and
is `null` when a shape inherits its position from a placeholder.

### `Presentation`

```ts
interface SlideSize {
	widthEmu: number
	heightEmu: number
	widthIn: number
	heightIn: number
}

class Presentation {
	static load(input: OpcInput): Promise<Presentation>
	static fromPackage(opc: OpcPackage): Presentation

	/** The underlying OPC package. */
	readonly opc: OpcPackage
	/** The main presentation part, via the package officeDocument relationship. */
	readonly presentationPart: Part
	/** Slides in presentation order (p:sldIdLst). */
	readonly slides: Slide[]
	/** Slide dimensions, or null if none declared. */
	readonly slideSize: SlideSize | null

	save(): Promise<Uint8Array>
}
```

### `Slide`

```ts
class Slide {
	readonly presentation: Presentation
	readonly part: Part
	readonly slideId: number // from p:sldId/@id
	readonly index: number // zero-based, in presentation order
	readonly partName: string
	readonly relationships: Relationships // this slide part's rels
	readonly name: string | null // p:cSld/@name
	readonly shapes: Shape[] // top-level shapes in the spTree
}
```

### `Shape` and subclasses

`slide.shapes` returns one proxy per shape-tree child, by element:

| Element           | Class          | `shapeType`     |
| ----------------- | -------------- | --------------- |
| `p:sp`            | `AutoShape`    | `autoShape`     |
| `p:pic`           | `Picture`      | `picture`       |
| `p:cxnSp`         | `Connector`    | `connector`     |
| `p:graphicFrame`  | `GraphicFrame` | `graphicFrame`  |
| `p:grpSp`         | `GroupShape`   | `group`         |

```ts
abstract class Shape {
	readonly shapeType: ShapeType
	readonly slide: Slide
	readonly id: number | null // p:cNvPr/@id
	readonly name: string // p:cNvPr/@name ('' if unnamed)
	readonly left: number | null // EMU (a:off/@x)
	readonly top: number | null // EMU (a:off/@y)
	readonly width: number | null // EMU (a:ext/@cx)
	readonly height: number | null // EMU (a:ext/@cy)
	readonly hasTextFrame: boolean
	readonly textFrame: TextFrame | null
	readonly text: string // textFrame?.text ?? ''
	readonly element_: Element // escape hatch to the DOM node
}

class AutoShape extends Shape {
	readonly presetGeometry: string | null // a:prstGeom/@prst, e.g. 'rect'
}

class Picture extends Shape {
	readonly imageRelId: string | null // a:blip/@r:embed
	readonly imagePartName: string | null // resolved via the slide's rels
}

class GraphicFrame extends Shape {
	readonly hasTable: boolean
	readonly hasChart: boolean
}

class GroupShape extends Shape {
	readonly shapes: Shape[] // nested children
}
```

Only `AutoShape` (`p:sp`) reports `hasTextFrame: true` and a non-null
`textFrame` in this read model.

### `TextFrame`, `Paragraph`, `Run`

```ts
class TextFrame {
	readonly paragraphs: Paragraph[]
	readonly text: string // paragraph texts joined by '\n'
}

class Paragraph {
	readonly runs: Run[] // a:r elements only
	readonly level: number // a:pPr/@lvl, 0 if unset
	readonly text: string // runs + fields, with a:br as '\n'
}

class Run {
	readonly text: string // a:t, verbatim
	readonly fontSizePt: number | null // a:rPr/@sz / 100
	readonly bold: boolean | null // null when unset (inherited)
	readonly italic: boolean | null // null when unset (inherited)
	readonly underline: string | null // a:rPr/@u token, e.g. 'sng'
	readonly fontName: string | null // a:latin/@typeface
	readonly color: string | null // a:srgbClr/@val (6-hex)
	readonly schemeColor: string | null // a:schemeClr/@val, e.g. 'accent2'
}
```

Boolean run properties are `null` when the attribute is absent — the value is
inherited from the list/placeholder style, not `false`. Explicit RGB colour and
theme colour are reported separately (`color` vs `schemeColor`); at most one is
non-null for a given run.

## Editing today (low-level)

Until the typed API lands, edits work directly on the DOM:

```js
const slide = pkg.part('/ppt/slides/slide1.xml')
const run = slide.dom.getElementsByTagName('a:t')[0]
run.textContent = 'New title'
slide.markDirty() // without this, save() writes the original bytes

const edited = await pkg.save()
```

Only the touched part is reserialized; everything else stays byte-identical.

## Testing

`pnpm run test:read` runs the round-trip harness
(`test/read/roundtrip.test.js`: part-set stability, byte-identity, laziness,
idempotence, content-type/relationship resolution, dirty-path, schema
validation) and the read-model tests (`test/read/model.test.js`: slide/shape
navigation, geometry, picture image resolution, table detection, run
formatting). Schema cases require the OOXML validator
(`./tools/ooxml-validator/install.sh`) and are skipped with a notice when it
is absent. See [testing](../testing.md).

Beyond the automated suite, `pnpm run test:read:emit` writes each fixture's
`load() → save()` output to `.tmp/roundtrip/` so it can be opened in PowerPoint
to confirm there is no repair prompt. As of 2026-06-13 this manual check has
passed on **PowerPoint for the web** for all four fixtures; the stricter
**desktop PowerPoint** check is still outstanding (tracked in
`test/read/fixtures/README.md`).
