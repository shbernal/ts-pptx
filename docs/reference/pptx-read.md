---
doc-schema-version: 1
title: "PPTX Read / Round-Trip"
summary: "Open an existing .pptx, read its slides/shapes/text, edit text, fonts, and geometry, and save it back losslessly."
read_when:
  - Opening or editing decks this library did not generate
  - Editing run text, fonts, or shape position/size in an existing deck
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

Status: **Phase 4 — rich content & structural edits**. On top of the Phase 1
OPC layer (load, parts, content types, relationships, lossless save), the
Phase 2 navigable read model (`Presentation → slides → shapes → text frame →
paragraphs → runs`), and the Phase 3 edit slice (**run text and character
formatting**, **shape position/size**, and **shape fill/line colour**), the
model now also covers
**tables** (incl. cell borders + style id), **charts** (read-only — classic
`c:chart` with axes/labels/legend/series formatting, plus the `cx:` chartEx family:
waterfall/funnel/treemap/…), **rich run formatting** (strike/caps/baseline/
highlight/hyperlink + paragraph line spacing), **pattern fill and the effect list**
(inner shadow/glow/reflection/soft edge), **slide background/number/autofit**,
**adding and removing shapes**, **adding
pictures**, and **slide cloning**. Setting a property or calling a mutator
mutates the live DOM in place and marks only the affected part(s) dirty, so
`save()` reserializes just those and keeps every other byte for byte.
Lower-level DOM mutation (below) still works for anything the typed setters do
not yet cover. Future directions not yet implemented are tracked outside this
repo in `../PPTX_EDITING_NEXT_STEPS.md`.

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

	/** All parts keyed by partname (e.g. '/ppt/slides/slide1.xml'), in zip/add order. */
	readonly parts: ReadonlyMap<string, Part>
	/** Content-type resolution + registration overlay over [Content_Types].xml. */
	readonly contentTypes: ContentTypes

	part(partName: string): Part | undefined
	partsByContentType(contentType: string): Part[]
	/** Relationships owned by a part; '/' (default) = package-level /_rels/.rels. */
	relationshipsFor(sourcePartName?: string): Relationships

	/** Add a part and register its content type. Throws if the partname is taken. */
	addPart(partName: string, contentType: string, bytes: Uint8Array): Part
	/** Reserve an unused '/ppt/media/<base><n>.<ext>' partname (does not create it). */
	reserveMediaPartName(extension: string, base?: string): string

	save(): Promise<Uint8Array>
}
```

`load()` rejects when the input is not an OPC package or when a part has no
resolvable content type (no `Override`, no `Default`) — the error names the
offending part.

`[Content_Types].xml` is not enumerated in `parts`; it is managed by the
package and exposed through the `contentTypes` overlay.

`save()` flushes any dirty `Relationships` set back into its `.rels` part
(creating it when new) and writes a regenerated `[Content_Types].xml` only when
a registration changed it; everything still untouched stays byte-identical.

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

Overlay over `[Content_Types].xml`: clean → bytes pass through; dirty →
`serialize()` is authoritative on save.

```ts
class ContentTypes {
	static parse(xml: string): ContentTypes
	/** Exact Override match first, else Default by lowercased extension. */
	contentTypeFor(partName: string): string | undefined
	readonly isDirty: boolean
	/** Ensure partName resolves to contentType (no-op if already; else adds an Override). */
	ensureRegistered(partName: string, contentType: string): void
	/** Register a Default content type for an extension if absent. */
	ensureDefault(extension: string, contentType: string): void
	serialize(): string
}
```

### `Relationships`

Overlay over one `.rels` part. Iterable. Clean → bytes pass through; once `add()`
marks it dirty, `OpcPackage.save()` writes `serialize()` into the `.rels` part.

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
	readonly isDirty: boolean
	get(id: string): Relationship | undefined
	byType(type: string): Relationship[]
	/** Absolute partname for an internal rel; throws for External rels. */
	resolveTarget(id: string): string
	/** Add a relationship, allocating 'rId<n>' past the highest existing id. */
	add(type: string, target: string, targetMode?: 'Internal' | 'External'): Relationship
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

## Object model (Phase 2 read, Phase 3 edit)

A navigable, typed view over the live DOM. Every proxy reads from its DOM
element on each access (no caching) and wraps the very nodes the setters mutate
in place. Geometry is reported in **EMU** (the OOXML unit; 914 400 per inch)
and is `null` when a shape inherits its position from a placeholder. Properties
documented below as *settable* write back to the DOM and mark the owning slide
part dirty (see [Editing](#editing-typed-api-phase-3)).

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

	/**
	 * Open a PowerPoint template (.pptx or .potx) as an empty deck shell: its
	 * masters/layouts/theme are kept byte-identical, any sample slides are stripped,
	 * and a .potx main part's template content type is normalized to the editable
	 * presentation type (unless `keepTemplateContentType`). Author onto it with
	 * `appendSlides`. See "Authoring slides onto a template or existing deck".
	 */
	static fromTemplate(input: OpcInput, options?: FromTemplateOptions): Promise<Presentation>

	/** The underlying OPC package. */
	readonly opc: OpcPackage
	/** The main presentation part, via the package officeDocument relationship. */
	readonly presentationPart: Part
	/** Slides in presentation order (p:sldIdLst). */
	readonly slides: Slide[]
	/** Slide dimensions, or null if none declared. */
	readonly slideSize: SlideSize | null

	/**
	 * Phase 4 — duplicate the slide at `index`, insert the copy at `options.at`
	 * (deck order; 0 = first; omitted/out-of-range appends), and return it.
	 */
	cloneSlide(index: number, options?: { at?: number }): Slide

	/**
	 * Phase 4 — copy `source.slides[index]` (from a *different* open package) and
	 * insert it at `options.at` (deck order; 0 = first; omitted/out-of-range
	 * appends), returning the new slide. With `theme: 'copy'` (default) it brings
	 * the slide's layout → master → theme and any media/chart/embedding parts; with
	 * `theme: 'preserve'` it bakes the source theme into the slide and binds it to
	 * this deck's existing master. Source and target slide sizes must match.
	 */
	importSlide(source: Presentation, index: number, options?: ImportSlideOptions): Slide

	/**
	 * Phase 4 — copy one shape from `source.shapes[shapeIndex]` (a slide of any
	 * open package) onto `target` (a slide of *this* presentation), returning the
	 * new Shape. Drags the shape's media/chart/embedding parts across (deduped via
	 * the copy registry), rewrites their relationship references to fresh host-slide
	 * rels, and reassigns the shape's (and any group children's) drawing ids. With
	 * `theme: 'preserve'` (default) it bakes the shape's theme references to literals
	 * against the source theme; `restyle` leaves them symbolic; `copy` is verbatim.
	 * Source and target slide sizes must match.
	 */
	importShape(target: Slide, source: Slide, shapeIndex: number, options?: ImportShapeOptions): Shape

	/** Phase 4 — batch form of `importShape`; media shared by the lifted shapes is copied once. */
	importShapes(target: Slide, source: Slide, shapeIndices: number[], options?: ImportShapeOptions): Shape[]

	/** The deck's layouts, in master-then-layout order — the gallery `appendSlides` binds to. Read-only; copies nothing. */
	layouts(): LayoutHandle[]

	/**
	 * Author the slides of a generator (`source`, e.g. a `PptxGenJS` instance) onto
	 * this deck, bound to one of its existing layouts (by `p:cSld@name` or a
	 * `LayoutHandle`), and return the new Slides. Masters/layouts/theme and every
	 * other untouched part stay byte-identical. Source and deck slide sizes must
	 * match. See "Authoring slides onto a template or existing deck".
	 */
	appendSlides(source: SlideSource, options: AppendSlidesOptions): Promise<Slide[]>

	save(): Promise<Uint8Array>
}

interface FromTemplateOptions {
	keepTemplateContentType?: boolean // keep a .potx main part as ...template.main+xml; default false (normalize to editable)
}

interface LayoutHandle {
	partName: string // the layout part's name, e.g. /ppt/slideLayouts/slideLayout2.xml
	name: string | null // p:cSld/@name, e.g. "Title and Content"
	masterPartName: string
	masterIndex: number
	layoutIndex: number
}

interface AppendSlidesOptions {
	layout: string | LayoutHandle // bind every appended slide to this layout (by name or handle)
	at?: number // zero-based p:sldIdLst position for the first appended slide; default append
	onMediaError?: 'throw' | 'placeholder' // how addImage media errors surface; default 'throw'
}

interface ImportSlideOptions {
	theme?: 'copy' | 'preserve' | 'restyle' // default 'copy'
	carryMasterGraphics?: boolean // preserve/restyle only; default false
	at?: number // insert position in p:sldIdLst (deck order); 0 = first; default append
}

interface ImportShapeOptions {
	theme?: 'preserve' | 'restyle' | 'copy' // default 'preserve'
	left?: number // EMU placement overrides; omitted axes keep the source xfrm
	top?: number
	width?: number
	height?: number
	at?: number // z-order insert position among host shape children; default append (on top)
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
	hidden: boolean // p:sld/@show — read/write; absent attr ⇒ shown
	readonly background: SlideBackground | null // effective bg, walking slide→layout→master
	readonly slideNumberPlaceholder: AutoShape | null // this slide's own p:ph type="sldNum"
	readonly notesText: string | null // flattened speaker-notes body text; null when there is no notes part
	readonly notesTextFrame: TextFrame | null // the notes body as a navigable frame (see below)
	readonly notesSlide: NotesSlide | null // the whole modeled notes slide (its three placeholders)
	addTextBox(options: AddTextBoxOptions): AutoShape // Phase 4 — appends a p:sp
	addPicture(image: Uint8Array, options: AddPictureOptions): Picture // Phase 4 — new media part + rel + p:pic
}

type SlideBackground = {
	type: 'solid' | 'gradient' | 'image' | 'pattern' | 'themeRef' | 'none'
	source: 'slide' | 'layout' | 'master' // which level in the inheritance chain won
	// …type-specific payload (colour, gradient stops, image part name, themeRef idx, …)
}
```

#### Hidden slides (`hidden`)

`slide.hidden` reflects `p:sld/@show`, an `xsd:boolean` that **defaults to
`true`**, so a slide with no `@show` attribute reads as shown (`false`).
PowerPoint writes `show="0"` when you hide a slide (the getter also accepts the
`"false"` lexical form).

This matters whenever you reconcile **render order** with **model order**:
PowerPoint's "present" and LibreOffice both drop hidden slides from a slideshow
and from exported PDFs, so once any earlier slide is hidden the Nth rendered page
is no longer `presentation.slides[N]`. The reconciliation falls out directly —
`slides.length − (visible count) === (hidden count)`:

```ts
const hidden = presentation.slides.filter((s) => s.hidden).length
const visible = presentation.slides.length - hidden // === rendered page count
```

The setter is symmetric and writes the canonical form: assigning `true` writes
`show="0"`; assigning `false` removes the attribute (PowerPoint's shown default),
marking only the owning slide part dirty.

```ts
presentation.slides[1].hidden = true // hide slide 2
presentation.slides[3].hidden = false // un-hide slide 4
await presentation.save()
```

#### Background, slide number, autofit

`slide.background` reports the **effective** background by walking the inheritance
chain — the slide's own `p:bg` wins, else the layout's, else the master's — and
`source` records which level won. That distinction matters: a plainly authored
slide has *no* own `p:bg`, so it falls through to the default layout's
`<p:bgRef idx="1001">` and reads as `{ type: 'themeRef', source: 'layout', idx: 1001 }`,
not as an authored background. Colour tokens resolve through the slide theme; an
image background's `r:embed` resolves against the **owning** part's rels (the
layout's rels for a layout-inherited image). solid/gradient/image are FAITHFUL;
pattern/themeRef are read-only for imported decks.

`slideNumberPlaceholder` is scoped to the slide's **own** shape tree — the
`p:ph type="sldNum"` the per-slide `slide.slideNumber = {…}` setter emits. It
deliberately does **not** resolve a number inherited purely from the master `p:hf`
(a master-level concern); date/footer placeholders are deferred for the same
reason (no writer-authored slide shape to read). Note `pres.setSlideNumber({…})`
puts the placeholder on the master/def-layout only, so a slide added afterward
reads `null` here.

`TextFrame.autofit` ports the mode onto the navigable text frame. The writer's
`fit: 'shrink'` emits a **bare** `<a:normAutofit/>` (PowerPoint computes the scale
on edit, so `autofitFontScale` reads `null`); an explicit baked scale needs the
object form `fit: { type: 'shrink', fontScale, lnSpcReduction }`.

#### Speaker notes

`slide.addNotes(...)` authors a notes slide whose body placeholder (`p:ph
type="body"`) holds the notes runs, serialized through the same text-run generator
as any shape — so bold/italic/underline/colour/size/face and an external-`url`
hyperlink all land in the notes `p:txBody`. The read side exposes that body two
ways, sharing one body-placeholder lookup:

- **`notesText`** — the flattened convenience: the body's text with paragraphs
  joined by `\n`. Character formatting and links are dropped.
- **`notesTextFrame`** — the same body as a navigable `TextFrame`
  (paragraphs → runs), so per-run formatting is recoverable and a notes hyperlink
  resolves its `url`. The frame is threaded with the **notes part's own rels**
  (`notesSlideN.xml.rels`), so `Run.hyperlink.url` resolves for notes links —
  unlike a table-cell run, which reports only the raw `relId`.

Both `notesText` and `notesTextFrame` are thin **delegates over `notesSlide.body`**
(below). They are `null` under the same boundary — no notes-slide part at all — and
`notesTextFrame` is *also* `null` when a notes part exists but carries no body text
frame (there is no frame to hand back), where `notesText` still reports `''`. Note
the writer attaches an **empty notes part to every authored slide** (to keep the
`notesSlide` rel/`_rels` bookkeeping uniform), so an authored slide never hits the
true no-part `null` path — that branch is reachable only from imported decks. A
`\n` in a note starts a new paragraph, so a multi-line note reads back as multiple
`paragraphs`.

The frame *is* threaded with a **notes theme context** — resolved through the notes
part's `notesMaster` rel → `theme2.xml` chain — so a notes run authored with a
*scheme* colour resolves to a literal hex via `Run.resolvedColor` (the `clrMap`
comes from the notesMaster's own `p:clrMap`, the `clrScheme`/`fontScheme` from
`theme2.xml`). What stays inert is *placeholder-inherited* resolution: the frame is
built without a placeholder context because notes inherit character properties
(size/face) from the notesMaster's `p:notesStyle`, not from a slide layout/master
placeholder chain — modeling that `notesStyle` inheritance is a separate, larger
surface.

##### The modeled notes slide (`notesSlide`)

`slide.notesSlide` returns the whole notes slide (`notesSlideN.xml`) as a
`NotesSlide`, or `null` at the same no-part boundary. A notes slide is a small,
fixed surface — a shape tree of exactly three placeholders, never groups/pictures/
connectors/charts — so it is modeled with a dedicated `NotesPlaceholder` rather than
the full `Shape` hierarchy:

```ts
type NotesSlide = {
	readonly part: Part
	readonly placeholders: NotesPlaceholder[] // the three, in document order
	readonly slideImage: NotesPlaceholder | null // p:ph type="sldImg" (the thumbnail)
	readonly body: NotesPlaceholder | null // p:ph type="body" (the notes text)
	readonly slideNumber: NotesPlaceholder | null // p:ph type="sldNum" (the slide-number field)
	readonly textFrame: TextFrame | null // === body?.textFrame
	readonly text: string // === body?.text ?? ''
}
type NotesPlaceholder = {
	readonly type: string | null // sldImg | body | sldNum
	readonly idx: string | null
	readonly name: string
	readonly id: number | null
	readonly left / top / width / height: number | null // own a:xfrm EMU; null when inherited
	readonly textFrame: TextFrame | null // null for the sldImg thumbnail
	readonly text: string
}
```

The measured fidelity is what the writer authors: the three placeholder
`type`/`name`/`idx`, the **body** text (round-tripped through the same frame
`notesTextFrame` hands back), and the **`sldNum`** slide-number `a:fld` — its value
surfaces through `TextFrame.text` (which reads `a:fld` text), so `slideNumber.text`
is the slide's number (`'1'` for the first slide). Geometry is **import-only**: the
writer leaves the `sldImg`/`sldNum` `p:spPr` empty, so on an authored deck every
placeholder's own `left/top/width/height` reads `null` (the geometry is inherited
from the notesMaster); an imported deck carries PowerPoint's stamped values. Each
placeholder's `textFrame` shares the same notes theme context as `notesTextFrame`.

### `Shape` and subclasses

`slide.shapes` returns one proxy per shape-tree child, by element:

| Element           | Class          | `shapeType`     |
| ----------------- | -------------- | --------------- |
| `p:sp`            | `AutoShape`    | `autoShape`     |
| `p:pic`           | `Picture`      | `picture`       |
| `p:cxnSp`         | `Connector`    | `connector`     |
| `p:graphicFrame`  | `GraphicFrame` | `graphicFrame`  |
| `p:grpSp`         | `GroupShape`   | `group`         |

Geometry is read/write. A getter returns `null` when the shape inherits its
position; a setter writes EMU into the shape's transform, creating the
transform (`a:xfrm`/`p:xfrm`) and its container in document order if absent.
Values are rounded to integer EMU; extents (`width`/`height`) reject negatives,
and all four reject `NaN`/`Infinity`.

`rotation`/`flipH`/`flipV` are read-only reads of the shape's own `a:xfrm`.
`rotation` is in **degrees** (the source stores 60000ths) and is faithful to the
XML — a negative angle stored as e.g. `19216344` reads back as `320.27`, not
normalized to a signed range. Like the geometry getters, `rotation` is `null`
when the shape has no own transform and `0` when it has one without a `@rot`.
These report the shape's **own** orientation; they are the per-shape complement
to `absoluteFrame`, which reports the effective position, size, rotation, and
flips after composing enclosing group transforms.

```ts
abstract class Shape {
	readonly shapeType: ShapeType
	readonly slide: Slide
	readonly id: number | null // p:cNvPr/@id
	readonly name: string // p:cNvPr/@name ('' if unnamed)
	left: number | null // EMU (a:off/@x) — settable
	top: number | null // EMU (a:off/@y) — settable
	width: number | null // EMU (a:ext/@cx) — settable
	height: number | null // EMU (a:ext/@cy) — settable
	readonly rotation: number | null // degrees (a:xfrm/@rot ÷ 60000); null when no own xfrm, 0 when present but unrotated
	readonly flipH: boolean // a:xfrm/@flipH; false when unset or no own xfrm
	readonly flipV: boolean // a:xfrm/@flipV; false when unset or no own xfrm
	readonly absoluteFrame: {
		left: number
		top: number
		width: number
		height: number
		rotation: number
		flipH: boolean
		flipV: boolean
	} | null // slide-absolute EMU/degrees after composing enclosing groups
	fillColor: string | null // spPr/a:solidFill/a:srgbClr/@val (6-hex) — settable
	fillSchemeColor: string | null // spPr/a:solidFill/a:schemeClr/@val, e.g. 'accent2' — settable
	lineColor: string | null // spPr/a:ln/a:solidFill/a:srgbClr/@val (6-hex) — settable
	lineSchemeColor: string | null // spPr/a:ln/a:solidFill/a:schemeClr/@val — settable
	noFill(): void // set an explicit <a:noFill/> (transparent surface)
	readonly customGeometry: CustomGeometry | null // spPr/a:custGeom/a:pathLst freeform paths; null for preset/none
	readonly patternFill: PatternFill | null // a:pattFill { preset, foreground, background } (fg/bg theme-resolved)
	readonly shadow: OuterShadow | null // a:effectLst/a:outerShdw
	readonly innerShadow: InnerShadow | null // a:innerShdw (structurally an OuterShadow)
	readonly glow: Glow | null // a:glow { radiusPt, color }
	readonly reflection: Reflection | null // a:reflection (read-only; writer authors none)
	readonly softEdge: SoftEdge | null // a:softEdge { radiusPt } (read-only)
	readonly hasTextFrame: boolean
	readonly textFrame: TextFrame | null
	readonly text: string // textFrame?.text ?? ''
	readonly element_: Element // escape hatch to the DOM node
}

class AutoShape extends Shape {
	readonly presetGeometry: string | null // a:prstGeom/@prst, e.g. 'rect'
}

class Picture extends Shape {
	imageRelId: string | null // a:blip/@r:embed — get, or set to repoint at an existing rel
	readonly imagePartName: string | null // resolved via the slide's rels
	setImage(bytes: Uint8Array, options: { contentType: string; extension?: string }): void // Phase 4 — swap the image
	// Fill setters throw (a picture's surface is out of scope for v1); lineColor
	// (the picture's border) is available.
}

class GraphicFrame extends Shape {
	readonly hasTable: boolean
	readonly hasChart: boolean // classic c:chart
	readonly hasChartEx: boolean // cx: chartEx (waterfall/treemap/…) — see ChartEx below
	readonly table: Table | null // non-null when hasTable
	readonly chart: Chart | null // non-null when hasChart (resolves the chart part)
	readonly chartEx: ChartEx | null // non-null when hasChartEx (resolves the cx:chartSpace part)
	// Fill and line setters throw: a graphicFrame has no p:spPr; its hosted
	// table/chart carries its own fill model.
}

class GroupShape extends Shape {
	readonly shapes: Shape[] // nested children
	// Fill setters write p:grpSpPr/a:solidFill; line setters throw (a group's
	// properties have no a:ln).
}

class Connector extends Shape {
	readonly startConnection: ConnectionSite | null // p:cNvCxnSpPr/a:stCxn — null when the start end is unbound
	readonly endConnection: ConnectionSite | null // p:cNvCxnSpPr/a:endCxn — null when the end end is unbound
	// Supports both fill and line (its outline is the connector itself).
}

interface ConnectionSite {
	shapeId: number // the bound shape's drawing id (p:cNvPr/@id)
	siteIndex: number // connection-site index on that shape (@idx, 0-based, preset-dependent)
	boundShape: AnyShape | null // resolved via slide.shapeByIdDeep (descends into groups); null only when no shape anywhere carries that id
}
```

Only `AutoShape` (`p:sp`) reports `hasTextFrame: true` and a non-null
`textFrame` in this read model.

#### Connector endpoint binding

A connector authored with `slide.addConnector({ startShape, endShape, … })` binds
each end to a shape: the writer resolves each target's `objectName` → drawing id at
serialize time and emits
`<p:cNvCxnSpPr><a:stCxn id idx/><a:endCxn id idx/></p:cNvCxnSpPr>`. The read side
decodes that into `ConnectionSite` per end and resolves the `@id` back to the slide
shape via `slide.shapeByIdDeep` (`boundShape`), which **descends into groups** — so
a connector bound to a shape nested in a group resolves the same as one bound to a
top-level shape (the writer already ids and binds group children). The two-getter
shape mirrors the write API's `startShape`/`endShape` split.

- An **unbound** end (the writer emits a bare `<p:cNvCxnSpPr/>` when a connector
  binds no shapes, or when an `objectName` doesn't resolve — it warns and falls
  back to static endpoint geometry) reports `null`, never a half-populated site.
- A **present** binding whose `@id`/`@idx` is unparseable degrades to `null` rather
  than a partial site; a binding whose shape id isn't found *anywhere* on the slide
  (a genuinely dangling id) keeps `shapeId`/`siteIndex` but leaves `boundShape`
  `null` (faithful degradation, no throw).
- Omitting `startShapeIdx`/`endShapeIdx` writes `idx="0"`, so a single-idx bind
  reads `siteIndex: 0`. `ConnectionSite` is exported from `pptxgenjs/read`.

#### Fill and line colour

`fillColor`/`fillSchemeColor` read and write the shape's solid fill
(`spPr/a:solidFill`); `lineColor`/`lineSchemeColor` do the same for its outline
(`spPr/a:ln/a:solidFill`). They mirror `Run.color`/`Run.schemeColor`:

- The `*Color` accessors take a 6-hex RGB string (optional leading `#`,
  normalized to upper-case; malformed input throws). The `*SchemeColor`
  accessors take a theme token (`accent2`, `bg1`, …). At most one of the RGB /
  scheme pair is non-null, so setting one **clears** the other.
- Setting `fillColor = null` (or `lineColor = null`) removes the `a:solidFill`,
  restoring inheritance from the shape's style/placeholder. This is **distinct**
  from `noFill()`, which writes an explicit `<a:noFill/>` — a deliberately
  transparent surface, not "inherit".
- A setter creates the properties element (`p:spPr`/`p:grpSpPr`), the `a:ln`, and
  the `a:solidFill` in OOXML document order if absent.
- Per-kind support follows the OOXML model (see the class notes above):
  `AutoShape` and `Connector` support both fill and line; `GroupShape` supports
  fill only; `Picture` supports line only; `GraphicFrame` supports neither.
  Setting an unsupported property throws. These are setters for the **token**;
  resolving a scheme colour to RGB is the deck theme's job, not this API's.

#### Custom geometry (freeform paths)

`customGeometry` is the freeform counterpart of `AutoShape.presetGeometry`: it
reads a shape's `spPr/a:custGeom/a:pathLst` and returns `null` when the shape
uses preset geometry or none. It lives on the base `Shape` so it covers both a
freeform `p:sp` and a `p:pic` clipped to a `custGeom`.

```ts
type GeometryCommand =
	| { cmd: 'moveTo'; x: number; y: number }
	| { cmd: 'lnTo'; x: number; y: number }
	| { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { cmd: 'quadBezTo'; x1: number; y1: number; x: number; y: number }
	| { cmd: 'arcTo'; wR: number; hR: number; stAng: number; swAng: number } // angles in degrees
	| { cmd: 'close' }

interface CustomGeometryPath {
	w: number // a:path/@w — path-unit width (the x denominator); default 0
	h: number // a:path/@h — path-unit height (the y denominator); default 0
	fill: string // a:path/@fill (ST_PathFillMode); default 'norm'
	stroke: boolean // a:path/@stroke; default true
	commands: GeometryCommand[] // segments in document order — order is the geometry
}

interface CustomGeometry {
	paths: CustomGeometryPath[] // one entry per <a:path>
}
```

The model is faithful, not flattened: `a:pathLst` is repeatable and each `a:path`
carries its own `fill`/`stroke`, so the read side keeps the array rather than
collapsing to the single-path write DSL. The command verbs deliberately mirror
the write-side `GeometryPoint` DSL, so a consumer maps a `GeometryCommand[]` to
`GeometryPoint[]` one-to-one.

- **Coordinates are raw path-unit integers** in the path's own `0..w` / `0..h`
  space, **not** EMU. To place them in slide space, scale against the path `w`/`h`
  and the shape's box (`width`/`height`). A guide-name (`ST_AdjCoordinate` string)
  reference is not produced by authored freeforms; a non-numeric coordinate
  degrades to `0` rather than throwing.
- **`arcTo` angles are degrees** (the raw 60000ths-of-a-degree values divided by
  60000), matching the write DSL's degree input.
- **Schema defaults are applied** when an attribute is absent: `fill='norm'`,
  `stroke=true`, `w=0`, `h=0`.

> One `a:path` is the rule for PowerPoint-authored freeforms. PowerPoint's own
> Merge Shapes (Union/Combine/Subtract) never emits more than one `a:path` per
> `custGeom`: a shape with a hole is a **single** `a:path` holding two
> `moveTo`…`close` contours in document order (outer ring + inner ring). So
> `paths.length` is 1 for PowerPoint output; a multi-`a:path` `a:pathLst` is
> schema-legal but comes from other producers (e.g. SVG import). The
> `customGeometry` test fixture (`test/read/fixtures/custgeom.pptx`) pins this.

#### Pattern fill and effects

`patternFill` decodes `a:pattFill` — `{ preset, foreground, background }`, with
`fgClr`/`bgClr` resolved through the shape's theme context (note they *wrap* a
colour element, unlike `a:highlight`'s bare child). The effect getters generalize
the shadow read over the shape's `a:effectLst`:

```ts
interface PatternFill {
	preset: string // a:pattFill/@prst
	foreground: ResolvedColor | null // a:fgClr
	background: ResolvedColor | null // a:bgClr
}
interface OuterShadow {
	// type: 'outer'; blurPt, distancePt, directionDeg, color, alignment, … (all optional)
}
interface InnerShadow {
	// type: 'inner'; structurally an OuterShadow minus sx/sy/kx/ky/algn/rotWithShape
}
interface Glow {
	radiusPt: number // a:glow/@rad ÷ 12700
	color: ResolvedColor | null
}
interface Reflection {
	/* stA/stPos/endA/endPos (÷100000 → 0–1), dir/fadeDir (÷60000 → deg), … all optional */
}
interface SoftEdge {
	radiusPt: number // a:softEdge/@rad ÷ 12700
}
```

FAITHFUL (the writer authors them): pattern fill, inner shadow, glow (write-side
text glow emits `a:glow`). READ-ONLY (the writer authors none — carried for
imported decks, never regenerated): reflection, soft edge. Every reflection field
is optional — an absent attribute is **omitted**, not zeroed. Colour alpha rides
the colour transform (`transparency:25` → `a:alpha 75000` → `alpha 0.75`).

### `TextFrame`, `Paragraph`, `Run`

```ts
class TextFrame {
	readonly paragraphs: Paragraph[]
	readonly text: string // paragraph texts joined by '\n'
	readonly autofit: AutofitMode | null // 'none'|'normAutofit'|'spAutoFit'; null when no a:bodyPr
	readonly autofitFontScale: number | null // a:normAutofit/@fontScale ÷ 1000 (percent)
	readonly autofitLineSpaceReduction: number | null // @lnSpcReduction ÷ 1000 (percent)
}

class Paragraph {
	readonly runs: Run[] // a:r elements only
	readonly level: number // a:pPr/@lvl, 0 if unset
	readonly text: string // runs + fields, with a:br as '\n'
	readonly lineSpacing: LineSpacing | null // a:pPr/a:lnSpc
}

class Run {
	text: string // a:t, verbatim — settable
	fontSizePt: number | null // a:rPr/@sz / 100 — settable
	bold: boolean | null // null when unset (inherited) — settable
	italic: boolean | null // null when unset (inherited) — settable
	underline: string | null // a:rPr/@u token, e.g. 'sng' — settable
	fontName: string | null // a:latin/@typeface — settable
	color: string | null // a:srgbClr/@val (6-hex) — settable
	schemeColor: string | null // a:schemeClr/@val, e.g. 'accent2' — settable
	readonly strike: string | null // a:rPr/@strike — raw token 'noStrike'/'sngStrike'/'dblStrike'
	readonly caps: string | null // a:rPr/@cap — raw 'none'/'small'/'all'
	readonly baselinePct: number | null // a:rPr/@baseline ÷ 1000 (superscript +30, subscript -40)
	readonly highlight: ResolvedColor | null // a:rPr/a:highlight — theme-resolved (effectiveHex)
	readonly hyperlink: RunHyperlink | null // a:rPr/a:hlinkClick
}

type LineSpacing =
	| { type: 'points'; valuePt: number } // a:spcPts/@val ÷ 100
	| { type: 'percent'; percent: number } // a:spcPct/@val ÷ 1000 (150000 → 150)

interface RunHyperlink {
	url: string | null // External rel target
	targetPartName: string | null // Internal rel → absolute part name
	action: string | null // a:hlinkClick/@action (e.g. ppaction://…)
	tooltip: string | null // @tooltip
	relId: string | null // raw @r:id
}
```

Boolean run properties are `null` when the attribute is absent — the value is
inherited from the list/placeholder style, not `false`. Explicit RGB colour and
theme colour are reported separately (`color` vs `schemeColor`); at most one is
non-null for a given run.

`strike`/`caps` surface the **raw OOXML token** (not a boolean), because the schema
is tri-state. `baselinePct` divides the `ST_Percentage` (1000ths of a percent) by
1000 to a plain percent — the writer maps `superscript`→30, `subscript`→−40.
`highlight` is theme-resolved (the writer only ever emits a hex, so `effectiveHex`
is that hex). `lineSpacing` is a discriminated union: `a:spcPts/@val` is hundredths
of a point; `a:spcPct/@val` is `multiple × 100000` (÷1000 → percent).

`hyperlink` resolves `@r:id` through the run's **owning-part relationships**: an
External rel fills `url`, an Internal rel fills `targetPartName` (absolute). Only
runs reached via `AutoShape.textFrame` are threaded with those rels; a hyperlink on
a table-cell or notes run reports the raw `relId`/`action`/`tooltip` with
`url`/`targetPartName` left `null` (faithful degradation, not a crash). Empty
`action=""`/`tooltip=""` (what the writer emits on a URL link) coerce to `null`.

Every `Run` property is writable. A setter creates the run's `a:rPr` (and any
needed child, e.g. `a:latin`, `a:solidFill`) in document order:

- `run.text = '...'` rewrites the `a:t`; whitespace-significant text
  automatically gets `xml:space="preserve"`.
- `fontSizePt` takes points (stored as hundredths); it rejects non-positive and
  non-finite values.
- `bold`/`italic` accept `true`/`false`/`null`; setting `null` **removes** the
  attribute (back to inherited) rather than writing `0`.
- `color` accepts a 6-hex RGB string (optional leading `#`, normalized to
  upper-case; malformed input throws); `schemeColor` accepts a theme token.
  A run carries at most one solid fill, so setting one **clears** the other;
  setting `color = null` removes the run's solid fill entirely.

### `Table`, `TableRow`, `TableCell` (Phase 4)

A `GraphicFrame` whose `hasTable` is true exposes its `a:tbl` as a `Table`:

```ts
class Table {
	readonly rows: TableRow[]
	readonly rowCount: number
	readonly columnCount: number // a:tblGrid/a:gridCol count
	readonly columnWidths: (number | null)[] // EMU, per grid column
	readonly firstRowHeader: boolean // a:tblPr/@firstRow
	readonly bandedRows: boolean // a:tblPr/@bandRow
	readonly styleId: string | null // a:tblPr/a:tableStyleId — raw style GUID (see below)
	cell(rowIndex: number, columnIndex: number): TableCell | null
}

class TableRow {
	readonly cells: TableCell[]
	readonly heightEmu: number | null // a:tr/@h
}

interface CellBorder {
	widthPt: number // a:ln/@w ÷ 12700
	dash: string | null // a:prstDash/@val
	color: string | null // effective hex of the stroke's solid fill
	schemeColor: string | null // unresolved theme token, when the stroke is a schemeClr
	noFill: boolean // explicit <a:noFill/> — a deliberately suppressed edge
}

interface CellBorders {
	left: CellBorder | null
	right: CellBorder | null
	top: CellBorder | null
	bottom: CellBorder | null
	tlToBr: CellBorder | null // a:lnTlToBr diagonal
	blToTr: CellBorder | null // a:lnBlToTr diagonal
}

class TableCell {
	text: string // settable convenience (see below)
	readonly textFrame: TextFrame | null // a:txBody — full per-run editing
	readonly gridSpan: number // a:tc/@gridSpan, default 1
	readonly rowSpan: number // a:tc/@rowSpan, default 1
	readonly isMergeContinuation: boolean // @hMerge / @vMerge set
	readonly borders: CellBorders | null // per-edge strokes; null when a:tcPr is bare/absent
}
```

`styleId` surfaces the **raw** `a:tableStyleId` GUID, not a resolved style — the
id is the codegen handoff to the writer's `tableStyle` option (whose built-in
members *are* the GUID string, so it round-trips verbatim); resolving the style
matrix is a separate concern. The stroke colour reuses the same solid-fill decode
as shape strokes.

Borders are faithful, not collapsed. The writer emits a **full four-side set on
every cell**, defaulting an unspecified side to `<a:ln w="0"><a:noFill/></a:ln>`,
so a *written* table's cells never read `borders === null` — that null path is
reachable only from PowerPoint-authored fixtures with a bare/absent `a:tcPr`. A
`noFill` edge (explicit suppression) is reported as such and is **distinct** from
an edge inherited from the table style; don't conflate them. The writer never
emits the diagonals (`tlToBr`/`blToTr`), but the reader decodes them for imported
decks.

`columnIndex` counts `a:tc` elements in the row, so a cell that spans columns
(`gridSpan > 1`) occupies one index; merged-away cells report
`isMergeContinuation: true`.

Two ways to edit cell text, both marking only the slide part dirty:

```js
const table = slide.shapes.find((s) => s.shapeType === 'graphicFrame' && s.table).table

// Convenience: replace the whole cell with one run, keeping the first run's
// character formatting (font, size, colour) when the cell already had a run.
table.cell(0, 0).text = 'Total'

// Precise: edit individual runs, exactly as on a shape's text frame.
const run = table.cell(1, 1).textFrame.paragraphs[0].runs[0]
run.text = '42'
run.bold = true
```

### `Chart`, `ChartSeries` (Phase 4, read-only)

A `GraphicFrame` whose `hasChart` is true resolves its chart part (via the
slide's `chart` relationship) and exposes it as a `Chart`:

```ts
class Chart {
	readonly part: Part
	readonly partName: string
	readonly chartType: string | null // first plot-area group, e.g. 'line' / 'bar' / 'pie'
	readonly chartTypes: string[] // all groups (combo charts have >1)
	readonly title: string | null // c:chart/c:title rich text
	readonly series: ChartSeries[]
	readonly categories: (string | null)[] // from the first series' cache
	readonly axes: ChartAxis[] // plot-area order
	readonly categoryAxis: ChartAxis | null // the c:catAx/c:dateAx
	readonly valueAxis: ChartAxis | null // the c:valAx
	readonly dataLabels: ChartDataLabels | null // group-wide c:dLbls (first plot group)
	readonly legend: ChartLegend | null // c:chart/c:legend
}

class ChartSeries {
	readonly index: number | null // c:ser/c:idx
	readonly name: string | null // cached c:tx
	readonly values: (number | null)[] // cached c:val (c:numCache)
	readonly categories: (string | null)[] // cached c:cat
	readonly fill: ChartFill | null // c:spPr solid / no-fill
	readonly line: ChartLine | null // c:spPr/a:ln — width, dash, colour
}

class ChartAxis {
	readonly kind: 'cat' | 'val' | 'date' | 'ser'
	readonly id: number | null
	readonly orientation: string | null // c:scaling/c:orientation
	readonly min: number | null // c:scaling/c:min
	readonly max: number | null // c:scaling/c:max
	readonly logBase: number | null
	readonly hidden: boolean // c:delete
	readonly position: string | null // c:axPos
	readonly majorGridlines: boolean
	readonly minorGridlines: boolean
	readonly title: string | null // rich text
	readonly numberFormat: AxisNumberFormat | null // c:numFmt { formatCode, sourceLinked }
	readonly majorTickMark: string | null
	readonly minorTickMark: string | null
	readonly tickLabelPosition: string | null
	readonly majorUnit: number | null
	readonly minorUnit: number | null
}

interface ChartLegend {
	position: string | null // c:legendPos/@val
	overlay: boolean
}
interface ChartDataLabels {
	showValue: boolean
	showSeriesName: boolean
	showCategoryName: boolean
	showPercent: boolean
	showLegendKey: boolean
	position: string | null // c:dLblPos/@val
	numberFormat: AxisNumberFormat | null
}
interface ChartFill {
	color: string | null // raw srgbClr hex (NOT theme-resolved)
	schemeColor: string | null // unresolved token
	noFill: boolean
}
interface ChartLine {
	widthPt: number | null // a:ln/@w ÷ 12700
	dash: string | null
	color: string | null
	schemeColor: string | null
	noFill: boolean
}
```

```js
const chart = slide.shapes.find((s) => s.shapeType === 'graphicFrame' && s.chart).chart
chart.chartType // 'line'
chart.series.map((s) => [s.name, s.values]) // [['Costs', [360000, …]], ['Revenue', […]]]
```

Charts are **read-only**: the values exposed are the cache PowerPoint stores
alongside the embedded workbook (`c:numCache` / `c:strCache`). Rewriting chart
data (which means rewriting the embedded `.xlsx`) is not yet supported.

The chart part has **no theme context**, so series/axis colours surface as a raw
`color` (srgbClr hex) plus an unresolved `schemeColor` token — deliberately *not*
flattened to an effective hex the way shape/run colours are. `dataLabels` is the
group-wide block after the series (not the per-series ones). Note that bar/area
series carry no `a:ln` by default, so `ChartSeries.line` is `null` for them; the
stroke path is exercised by line/radar series.

### `ChartEx` — Office-2016 charts (waterfall / funnel / treemap / …)

The `cx:` chart family (waterfall, funnel, treemap, sunburst, histogram, pareto,
box-and-whisker, region map) is a **separate subsystem** from classic `c:chart`:
its own part (`cx:chartSpace`, `application/vnd.ms-office.chartex+xml`), referenced
via the MS chartEx relationship, and the frame is wrapped in `<mc:AlternateContent>`.
A `GraphicFrame` distinguishes the two:

```ts
class GraphicFrame extends Shape {
	readonly hasChart: boolean // classic c:chart
	readonly hasChartEx: boolean // cx: chartEx
	readonly chart: Chart | null
	readonly chartEx: ChartEx | null
}

class ChartEx {
	readonly part: Part
	readonly partName: string
	readonly layoutIds: string[] // raw cx:series/@layoutId tokens, e.g. ['clusteredColumn','paretoLine']
	readonly layoutId: string | null // first layoutId
	readonly chartType: string | null // alias of layoutId (the OOXML truth, not a write-side ChartType)
	readonly title: string | null
	readonly legend: ChartExLegend | null
	readonly series: ChartExSeries[]
	readonly axes: ChartExAxis[]
	readonly categories: string[] // first series' leaf-level categories
}

class ChartExSeries {
	readonly layoutId: string | null
	readonly ownerIndex: number | null // cx:series/@ownerIdx — a derived series' source (e.g. paretoLine → 0)
	readonly name: string | null // cx:tx/cx:txData/cx:v
	readonly dataId: number | null // cx:dataId/@val → the cx:chartData block
	readonly values: (number | null)[] // resolved through dataId → cx:numDim
	readonly categories: string[] // resolved through dataId → cx:strDim (leaf level)
	readonly dataLabels: ChartExDataLabels | null
}

class ChartExAxis {
	readonly id: number | null
	readonly kind: 'cat' | 'val' | null // from the scaling child, NOT an element name
	readonly gapWidth: number | null // cx:catScaling/@gapWidth — a FRACTION (0.5), not a percent
	readonly min: number | null
	readonly max: number | null
	readonly majorGridlines: boolean
	readonly tickLabels: boolean
}
```

Faithful-mapping decisions worth not re-litigating:

- **`layoutId` is not reverse-mapped to a write-side `ChartType`.** A histogram and
  a pareto both render as `clusteredColumn` (pareto adds a second `paretoLine`
  series), so the token doesn't uniquely identify the authoring type. The reader
  exposes the raw `layoutIds`, never a guessed type.
- **Categories read the *first* `cx:lvl`** — the writer emits hierarchy levels
  leaf-first, so the first level is the leaf labels (a treemap reads
  `['US','CA','DE']`, not the parent continents).
- **Axis kind comes from the scaling child** (`cx:catScaling` vs `cx:valScaling`),
  not an element name; **`gapWidth` is a fraction** (0.5 = 50%), read as a float —
  unlike the classic axis's integer percent.
- Data lives in a top-level `cx:chartData/cx:data` block the series point at by
  `cx:dataId`, not the classic `c:cat`/`c:val` caches. The embedded workbook, the
  style/colors sidecars, and a region map's online geo-cache are out of scope.

> **Enumeration prerequisite:** a chartEx frame is wrapped in `mc:AlternateContent`,
> which the shape walker originally skipped — so the frame was invisible.
> `buildShapes` now unwraps `mc:AlternateContent` (preferring the `mc:Choice` shape,
> else `mc:Fallback`). This also surfaces zoom frames and inline-math shapes that
> were likewise hidden.

## Editing (typed API, Phase 3)

Edit through the read model and save. Only the parts you touch are
reserialized; everything else stays byte-identical.

```js
import { readFile, writeFile } from 'node:fs/promises'
import { Presentation } from '@shbernal/pptxgenjs/read'

const presentation = await Presentation.load(await readFile('deck.pptx'))
const shape = presentation.slides[0].shapes.find((s) => s.name === 'Title')

// Geometry (EMU)
shape.left = 914400 // 1"
shape.top = 457200 // 0.5"
shape.width = 8229600

// Fill + line colour
shape.fillColor = '1F4E79' // explicit RGB; clears any scheme fill on the shape
shape.lineColor = 'D4D4D4' // shape outline

// Text + character formatting
const run = shape.textFrame.paragraphs[0].runs[0]
run.text = 'New title'
run.fontSizePt = 32
run.bold = true
run.color = '1F4E79' // explicit RGB; clears any scheme colour on the run

await writeFile('deck-edited.pptx', await presentation.save())
```

Each setter marks only the owning slide part dirty. The scope of the typed
slice is the read-model properties above: run text, `fontSizePt`, `bold`,
`italic`, `underline`, `fontName`, `color`, `schemeColor`; shape
`left`/`top`/`width`/`height`; and shape `fillColor`/`fillSchemeColor`/
`lineColor`/`lineSchemeColor` plus `noFill()`.

### Targeting a shape and replacing all its text

To swap the content of a known shape without walking the `shapes` array, a slide
exposes three finders:

```js
slide.shapeByName('Title') // first top-level shape with that p:cNvPr/@name
slide.shapeById(5) // first top-level shape with that p:cNvPr/@id
slide.placeholder('ctrTitle') // first placeholder of that p:ph/@type
slide.placeholder('subTitle', '1') // …narrowed by idx (defaults to '0' when absent)
```

`placeholder(type, idx?)` returns an `AutoShape` (only `p:sp` shapes can be
placeholders); read a shape's own placeholder identity via `shape.placeholder`
(`{ type, idx } | null`). All three finders scan **top-level** shapes only — a
shape nested in a group is not matched. To find a shape by id *across* groups, use
`slide.shapeByIdDeep(id)`, which walks group subtrees pre-order (this backs the
connector `boundShape` resolution); or walk `groupShape.shapes` yourself.

To replace **all** of a shape's text in one call, set `shape.text` (or
`textFrame.text`). It collapses the body to a single paragraph and run,
preserving the **first** existing run's character formatting (`a:rPr`) — the same
behaviour as `TableCell.text`:

```js
slide.shapeByName('Title').text = 'New title' // keeps the first run's font/size/colour
slide.placeholder('subTitle', '1').text = 'New subtitle'
```

Setting `text` on a shape with no text frame (e.g. a picture) throws. For
multiple runs or per-run formatting, edit `textFrame.paragraphs[].runs[]`
directly instead — that path preserves every run's own formatting, so it is the
right tool when you want to change one run and leave its siblings untouched.

### Adding and removing shapes (Phase 4)

Add a text box to a slide, or remove any shape, mutating only the slide part:

```js
const slide = presentation.slides[0]

// Add — geometry in EMU; width/height must be positive. A slide-unique
// drawing id is allocated automatically. Returns the new AutoShape.
const box = slide.addTextBox({
	text: 'Quarterly review',
	left: 914400, // 1"
	top: 457200, // 0.5"
	width: 4572000, // 5"
	height: 914400, // 1"
	name: 'Caption', // optional; defaults to `TextBox <id>`
})
box.textFrame.paragraphs[0].runs[0].bold = true // edit it like any shape

// Remove — detaches the shape from the slide (or its enclosing group).
slide.shapes.find((s) => s.name === 'Old caption')?.delete()
```

`addTextBox` builds a minimal, schema-valid `p:sp` (`txBox="1"`, a `rect`
preset geometry, and one paragraph). For richer shapes, add the text box and
then mutate it, or use the low-level escape hatch below.

Add a picture from raw image bytes — this creates a `/ppt/media/` part,
registers its content type, and wires an `image` relationship from the slide:

```js
import { readFile } from 'node:fs/promises'

const png = await readFile('logo.png')
slide.addPicture(png, { left: 914400, top: 457200, width: 1828800, height: 1828800 })
// The PNG/JPEG/GIF/BMP/TIFF/WebP format is sniffed from the bytes; pass
// { extension, contentType } to override or for an unrecognized format.
```

On save, the new media part is appended, the slide's `.rels` is rewritten with
the added relationship, and `[Content_Types].xml` is regenerated only if the
image's type was not already registered — every other part stays byte-identical.

### Replacing a picture's image (Phase 4)

`Picture.setImage` swaps the bytes behind an existing picture — the primitive a
stitching workflow needs when it lifts a slide from a reference deck and drops in
its own logo or photo. Like `addPicture`, it mints a `/ppt/media/` part,
registers its content type, and wires an `image` relationship from the slide;
then it repoints the picture's blip (`a:blip/@r:embed`) at the new part:

```js
import { readFile } from 'node:fs/promises'

const logo = await readFile('our-logo.png')
const picture = slide.shapes.find((s) => s.shapeType === 'picture')
picture.setImage(logo, { contentType: 'image/png' })
// `contentType` is required (the bytes are not sniffed); `extension` defaults
// from it (image/png → png, image/svg+xml → svg, …) and can be passed to override.
```

`setImage` is **copy-on-write**: it always adds a new media part and never
mutates or removes the old one. After `importSlide` (and in PowerPoint's own
dedup) a single media part is frequently shared by several pictures, so
overwriting bytes in place would silently change every picture pointing at it.
Minting a fresh part means the swap affects exactly this one picture; the
now-orphaned old part is left in place (harmless, just not pruned).

Geometry and crop are left untouched — `setImage` repoints the blip and leaves
`a:xfrm` and any `a:srcRect` as-is, so the caller owns sizing. To point a picture
at an image **already** present in the slide's relationships without adding a
part, assign the rel id directly: `picture.imageRelId = otherPicture.imageRelId`.

### Cloning a slide (Phase 4)

Duplicate an existing slide and append the copy to the deck:

```js
const clone = presentation.cloneSlide(0) // returns the new (last) Slide
clone.shapes.find((s) => s.hasTextFrame).textFrame.paragraphs[0].runs[0].text = 'Copy'
```

The clone gets its own slide part (a verbatim byte copy of the source) and its
own `.rels`, so it shares the source's layout, theme, and images by reference
while staying independent for edits. A presentation→slide relationship and a
`p:sldId` (with a fresh slide id) are wired up; only `presentation.xml`, its
`.rels`, and `[Content_Types].xml` change, plus the two new slide parts.

Relationships are copied as-is. If the source slide owns a one-to-one part such
as a notes slide, the copy would reference the same part; clone slides without
per-slide notes, or detach them afterward via the low-level API.

Pass `{ at }` to place the duplicate at a specific deck position instead of
appending: `presentation.cloneSlide(0, { at: 0 })` makes the copy the new first
slide. `at` is a zero-based index into `p:sldIdLst` (deck order); an `at` past the
current slide count — or omitting it — appends. The returned slide's `.index`
reflects where it landed.

### Importing a slide from another deck (Phase 4)

Copy a slide from one open package into a different one. Unlike `cloneSlide`
(same-deck duplicate), `importSlide` copies the connected sub-graph the slide
depends on — its `slideLayout → slideMaster → theme`, plus any media, charts, and
embeddings — into the target under fresh partnames:

```js
const target = await Presentation.load(await readFile('deck.pptx'))
const source = await Presentation.load(await readFile('library.pptx'))
const imported = target.importSlide(source, 0) // returns the new (last) Slide
const bytes = await target.save()
```

Only the layout(s) actually used by imported slides are copied, and the imported
master's `p:sldLayoutIdLst` is pruned to exactly those — mirroring PowerPoint's
"Reuse Slides". Parts shared by repeated imports from the same source deck are
copied once and reused. Untouched parts of the target stay byte-identical.

Source and target slide sizes must match (`importSlide` throws otherwise; v1 does
no geometry rescaling). Source notes are dropped, and fonts embedded via
`presentation.xml` are not carried across.

#### Slide position: `at`

By default the imported slide appends. Pass `{ at }` to insert it at a specific
deck position — the same zero-based `p:sldIdLst` index as `cloneSlide`'s `at`,
where `0` makes it first and an out-of-range/omitted `at` appends. This places
brand **bookends** around generator-authored interior slides regardless of import
order — a cover first, a closer last:

```js
deck.importSlide(source, COVER_INDEX, { theme: 'copy', at: 0 }) // cover first
deck.importSlide(source, CLOSER_INDEX, { theme: 'copy' })       // closer appended last
```

`importSlide` and `cloneSlide` are the read/import API; interior slides are
authored with the generate API (`new PptxGenJS()`). The two compose: emit the
generated deck to bytes (`await pptx.stream()`), `Presentation.load` those bytes,
`importSlide` the bookends, then `await deck.save()`.

#### Themes: `copy` (default) vs `preserve`

Each imported slide is structurally bound to its own `slideLayout → slideMaster →
theme`. The default `theme: 'copy'` brings that whole subgraph across, so a deck
stitched from N source decks carries **N themes / N masters**. That renders
faithfully in PowerPoint, but it is untidy for handoff and trips renderers
(notably LibreOffice) that resolve a slide's per-element `schemeClr` / style-matrix
references against the *wrong* (first) theme — branded backgrounds turn white and
scheme-coloured fills turn black, while literal `srgbClr` content is unaffected.

`theme: 'preserve'` fixes both by **flattening then attaching**:

```js
const imported = target.importSlide(source, 0, { theme: 'preserve' })
```

- **Flatten** — bake what the *source* theme would have produced into the slide
  XML: every `a:schemeClr` is resolved through the source `clrMap`/`clrScheme` to a
  literal `a:srgbClr` (colour transforms like `lumMod`/`shade` carried through
  unchanged, so tints render identically); each shape's `p:style`
  `lnRef`/`fillRef`/`effectRef` is resolved from the theme `fmtScheme` into an
  explicit `spPr` fill/line/effect and neutralized; and the slide's *effective
  background* (its own `p:bg`, else the one it inherited from the source
  layout/master, including a theme-indexed `p:bgRef`) is resolved to a literal
  `p:bgPr` and written onto the slide so it survives rebinding.
- **Carry inherited placeholder values** — a placeholder draws position, size,
  colour, and font from the source `slideLayout`/`slideMaster` it no longer
  points at after the rebind, so anything it does not set explicitly would snap
  to the destination master's defaults. `preserve` resolves and bakes that
  inheritance onto the slide: a placeholder with no own `a:xfrm` gets the
  effective `a:xfrm` (off/ext) from the matching source layout (else master)
  placeholder, so titles cannot shift or clip; and each placeholder run that
  sets none of its own gets the inherited colour and size/weight (`sz`/`b`/`i`),
  resolved per paragraph list level through the source placeholder `a:lstStyle`
  → master `a:lstStyle` → master `p:txStyles` chain. Typeface (`a:latin`) is
  deliberately left unbaked — it re-binds to the destination theme along with
  `fontRef` (see below).
- **Attach** — bind the now theme-independent slide to *this* deck's existing
  master/layout instead of importing the source theme. The result is a
  single-theme file whose imported slides keep their original colours.

Because the colours are frozen to literals, `preserve` does not re-colour to the
destination brand — its thesis is "same pixels, one theme". The `fontRef` and
typeface are deliberately left to re-bind to the destination theme (a font
normalization bonus on attach). Deliberate re-branding (a `restyle` mode) is not
yet implemented.

Decorative graphics on the source `slideMaster`/`slideLayout` shape trees (logos,
accent shapes, drawn footers — everything there *except* placeholders) belong to
the master that `preserve` drops, so by default they do not travel with the slide.
Pass `carryMasterGraphics: true` to bake them onto the imported slide behind its
own content (their media copied across and theme references flattened the same
way), for cover/divider slides whose branding must survive the rebind:

```js
const imported = target.importSlide(source, 0, { theme: 'preserve', carryMasterGraphics: true })
```

### Composing a slide from shapes of several decks (Phase 4)

Where `importSlide` brings a **whole** slide across, `importShape` lifts an
**individual** shape — an autoshape, picture, table, chart, connector, or group —
from any open deck onto a slide of *this* presentation. It is the primitive behind
a "stitching" workflow: build one target slide from, say, the comparison table of
deck A's slide 38 and the icon row of deck B's slide 34.

```js
const target = await Presentation.load(await readFile('deck.pptx'))
const libraryA = await Presentation.load(await readFile('library-a.pptx'))
const libraryB = await Presentation.load(await readFile('library-b.pptx'))

const slide = target.slides[0]
// Lift the table at index 2 of libraryA's slide 38…
const table = target.importShape(slide, libraryA.slides[38], 2)
// …and three icons from libraryB's slide 34, repositioned, in one call.
const icons = target.importShapes(slide, libraryB.slides[34], [4, 5, 6], { left: 1_000_000, top: 4_000_000 })

const bytes = await target.save()
```

`importShape(target, source, shapeIndex)` resolves `source.shapes[shapeIndex]` and
copies that subtree self-consistently:

- **Dependencies travel.** Every media / chart (and its embedded workbook) /
  embedding the shape references is copied into this package under a fresh
  partname — deduped against earlier imports from the same source deck via the
  copy registry — and its `r:embed` / `r:id` / `r:link` are rewritten to fresh
  host-slide relationships. So pictures, styled tables, and charts come across
  intact, not as re-synthesized plain shapes.
- **Ids cannot collide.** The lifted shape's `p:cNvPr/@id` (and every group
  child's) is reassigned to ids unused on the host slide.
- **Placement.** `left` / `top` / `width` / `height` (EMU) override the shape's
  source `a:xfrm`; omitted axes keep it verbatim (no rescale). `at` sets the
  z-order insert position among the host's shape children (default: append, on
  top). A batch inserts in the given order starting at `at`.

#### Themes: `preserve` (default), `restyle`, `copy`

Same three semantics as `importSlide`, scoped to the one shape subtree:

- **`preserve`** (default) — bake the shape's `a:schemeClr` and `p:style`
  `lnRef`/`fillRef`/`effectRef` to literals against the *source* theme, so it keeps
  its look on a host slide whose theme differs. A lifted *placeholder* also gets
  its inherited geometry/colour/size baked (best-effort — prefer lifting concrete
  content shapes over placeholders). Unlike a slide import this never runs the
  slide-scoped background passes; a background belongs to a slide, not a shape.
- **`restyle`** — leave the shape's theme references symbolic so it re-brands to
  the host theme. Only *symbolic* colours re-brand; a literal `a:srgbClr` the
  source baked in stays put.
- **`copy`** — bring the XML across untouched; only sane when the host already
  shares the source theme.

v1 limitations match `importSlide`: source and target slide sizes must match
(no geometry rescale), and the source slide's build animation/timing for the
lifted shape is dropped (the result is an editable static layout).

### Authoring slides onto a template or existing deck

`importSlide` / `importShape` move authored content *between loaded decks*. The
complementary path is to **generate new slides and graft them onto a loaded deck**,
reusing its masters/layouts/theme verbatim — the hybrid "generate-onto-existing"
workflow. Two methods cover it:

- **`presentation.layouts()`** enumerates the deck's layout gallery as
  `LayoutHandle[]` (master-then-layout order). It is a read-only discovery call —
  it copies nothing and leaves the package byte-identical. The `name` is the
  layout's `p:cSld@name` ("Title and Content", "Blank", …), which is what you bind
  to.
- **`presentation.appendSlides(source, { layout })`** authors the slides of a
  *generator* (`source` — any object exposing `extractSlides()`, which a `PptxGenJS`
  instance does) and splices them into this deck, each slide bound to the named
  existing layout. Only `presentation.xml`, its `.rels`, `[Content_Types].xml`, and
  the new slide/media/chart parts change; masters, layouts, theme, and every other
  untouched part stay byte-identical. Source and deck slide sizes must match
  (`appendSlides` throws otherwise — size the generator to the deck).

```js
import PptxGenJS from '@shbernal/pptxgenjs'
import { Presentation } from '@shbernal/pptxgenjs/read'

const deck = await Presentation.load(await readFile('deck.pptx'))

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE' // must match deck.slideSize
pptx.addSlide().addText('Generated', { x: 1, y: 1, w: 6, h: 1 })

const added = await deck.appendSlides(pptx, { layout: 'Title and Content' })
const bytes = await deck.save()
```

Each appended slide's `slideLayout` relationship is repointed at the **existing**
layout part (no new chrome is created); relationship ids inside the slide body are
preserved and only their targets are rewritten. Text, images, charts (chart XML +
`.rels` + embedded workbook), embedded audio/video, and internal slide-to-slide
hyperlinks (`slide:N`, repointed at the Nth appended slide) all carry across. Pass
`{ at }` to insert at a specific deck position (zero-based `p:sldIdLst` index, same
convention as `cloneSlide`/`importSlide`), and `{ onMediaError: 'placeholder' }` to
substitute a placeholder instead of throwing when an `addImage` source can't be
read.

#### Starting from a PowerPoint template — `fromTemplate`

To author a fresh deck on a **corporate template** instead of an existing deck,
open it with `Presentation.fromTemplate(input)`. It returns the template as an
empty shell ready for `appendSlides`:

```js
const deck = await Presentation.fromTemplate(await readFile('brand.potx')) // .pptx or .potx
deck.layouts().map((l) => l.name) // discover the template's layouts

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE' // size to deck.slideSize
pptx.addSlide().addText('Hello', { x: 1, y: 1, w: 6, h: 1 })

await deck.appendSlides(pptx, { layout: 'Title and Content' })
const out = await deck.save() // editable .pptx using the template's masters/layouts/theme
```

`fromTemplate` does two things on top of `load`:

- **Strips sample slides to a shell.** Most templates ship with sample slides you
  don't want; they are removed via the same pruning `removeSlide` uses, which never
  touches shared chrome, so masters/layouts/theme stay byte-identical. A template
  that already has zero slides makes this a no-op.
- **Normalizes a `.potx` to an editable `.pptx`.** A `.potx` package declares its
  main part with content type `…presentationml.template.main+xml`; by default that
  `[Content_Types].xml` override is flipped to `…presentationml.presentation.main+xml`
  so the saved file opens as a normal editable deck rather than spawning a new one
  from a template. Pass `{ keepTemplateContentType: true }` to keep the template
  type. (A `.pptx` input is already editable and needs no flip.)

This is higher fidelity than rebuilding the masters in code with
`defineSlideMaster()`: the template's authored master/layout/theme parts are kept
verbatim rather than round-tripped through the generator's lossy model. The only
requirement is that the generator's slide size matches the template's
(`deck.slideSize`).

### Editing anything else (low-level escape hatch)

For structure the typed setters do not yet cover, mutate the DOM directly and
mark the part dirty yourself — `element_` gives you the live node:

```js
const slide = presentation.opc.part('/ppt/slides/slide1.xml')
const run = slide.dom.getElementsByTagName('a:t')[0]
run.textContent = 'New title'
slide.markDirty() // without this, save() writes the original bytes

const edited = await presentation.save()
```

Only the touched part is reserialized; everything else stays byte-identical.

## Testing

`pnpm run test:read` runs the round-trip harness
(`test/read/roundtrip.test.js`: part-set stability, byte-identity, laziness,
idempotence, content-type/relationship resolution, dirty-path, schema
validation), the read-model tests (`test/read/model.test.js`: slide/shape
navigation, geometry, picture image resolution, table detection, run
formatting), and the edit tests (`test/read/edit.test.js`: text/font/geometry
setters survive a save → reopen round-trip, untouched parts stay
byte-identical, edited packages stay schema-valid, and invalid input is
rejected), and the table tests (`test/read/table.test.js`: table/row/cell
navigation, merge metadata, and cell-text edits surviving a round-trip), and
the structural-edit tests (`test/read/shapes-edit.test.js`: `addTextBox` /
`delete` surviving a round-trip with untouched parts byte-identical), and the
shape fill/line tests (`test/read/shape-fill-edit.test.js`: `fillColor` /
`lineColor` / `noFill()` round-tripping, document-order insertion, per-kind
support, and edited packages staying schema-valid), and the
picture tests (`test/read/picture-edit.test.js`: `addPicture` creating a media
part + content-type + relationship, format sniffing, and `setImage` swapping a
picture's bytes copy-on-write — minting a fresh part, repointing the blip, and
leaving the old part and any sibling sharing it untouched), and the clone tests (`test/read/clone-slide.test.js`:
`cloneSlide` appending an independent duplicate with correct presentation/rels
wiring), and the import tests (`test/read/import-slide.test.js`: `importSlide`
copying a slide's layout/master/theme/media sub-graph across a package boundary,
deduping a shared master and pruning its layout list, dropping notes, rejecting a
size mismatch, and staying schema-valid), and the theme-preserve import tests
(`test/read/import-slide-preserve.test.js`: `importSlide({ theme: 'preserve' })`
flattening scheme colours and `p:style` refs to literals, carrying the slide's
effective background, baking each placeholder's inherited run colour, geometry
(`a:xfrm`), and run size onto the slide, optionally carrying source
master/layout decorations via `carryMasterGraphics`, attaching to the
destination master without a new theme, and staying schema-valid), and the
shape-import tests (`test/read/import-shape.test.js`: `importShape`/`importShapes`
lifting a picture/table/chart/group onto a foreign host, deduping shared media,
reassigning ids off every host id, baking scheme colours to literals under
`preserve` vs leaving them symbolic under `restyle`, honouring placement + z-order
overrides, batching in order, rejecting size/index/ownership errors, and staying
schema-valid), and the chart tests
(`test/read/chart.test.js`: chart part resolution, type/title/series/values
reads, and a read-only open staying byte-identical), and the append tests
(`test/read/append-onto-existing.test.js`: `appendSlides` authoring generator
slides onto a loaded deck bound to an existing layout, keeping chrome
byte-identical, carrying text/image/chart/internal-link/audio/video, and staying
schema-valid), and the template tests (`test/read/template-masters.test.js`:
`fromTemplate` stripping sample slides to a shell while preserving the layout
gallery and chrome byte-for-byte, flipping a `.potx` main part to the editable
presentation content type — verified against the PowerPoint-authored
`template.potx` oracle — honouring `keepTemplateContentType`, and authoring onto a
zero-slide template shell to a schema-valid result).
The read-model expansions above are each proven by a **write→read fidelity**
suite built on the shared harness `test/read/authored.js`: author the feature with
the write API (which already emits it), load the bytes back through the deep read
model, and assert the extracted model — keeping the write and read paths
independent so a bug in one can't mask a bug in the other. Those suites are
`table-borders.test.js`, `chart-format.test.js`, `run-props.test.js`,
`chartex-read.test.js`, `connector-read.test.js` (endpoint binding),
`notes-read.test.js` (speaker-notes rich text), and the fidelity legs added to
`shape-effect-reads.test.js` and `slide-read-edges.test.js`. Schema cases require the OOXML validator
(`./tools/ooxml-validator/install.sh`) and are skipped with a notice when it
is absent. See [testing](../testing.md).

Beyond the automated suite, two scripts emit decks for a manual PowerPoint open
(schema validity is necessary but does not prove PowerPoint won't show a repair
prompt):

- `pnpm run test:read:emit` writes each fixture's unmodified `load() → save()`
  output to `.tmp/roundtrip/` — confirms the round-trip envelope opens clean.
- `pnpm run test:read:emit:edits` writes one *edited* deck per editing
  capability (added text box, added picture, deleted shape, cloned slide, edited
  table cells, imported image/table slide) to `.tmp/read-edits/` — confirms the
  reserialized/added parts open
  clean and render as intended. This is the check that matters for the editing
  API, since desktop PowerPoint validates the reserialized XML more strictly than
  the web.

Both checklists (web + desktop, with current status) live in
`test/read/fixtures/README.md`.
