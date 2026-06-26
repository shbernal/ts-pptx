/**
 * Read-model proxy for one slide (`p:sld`), backed by its live part DOM.
 */
import type { Part } from '../opc/part.js'
import { relativePartName } from '../opc/partnames.js'
import type { Relationships } from '../opc/relationships.js'
import { OOXML_NS, attr, createElement, firstChild, insertInOrder, intValue, removeAttr, setAttr, type Document, type Element } from '../oxml/dom.js'
import type { FlattenContext } from '../oxml/theme.js'
import { resolveSlideColorContext } from './theme-context.js'
import type { Presentation } from './presentation.js'
import { AutoShape, Picture, buildShapes, type Shape } from './shapes.js'
import { buildTransition, parseTransition, removeTransition, type TransitionInfo, type TransitionInput } from './transition.js'
import { enumerateSpids, hasAnimations, pruneSpids, remapSpids } from './animation.js'

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/** Options for {@link Slide.addTextBox}. Geometry is in EMU. */
export interface AddTextBoxOptions {
	/** Left edge in EMU (`a:off/@x`). */
	left: number
	/** Top edge in EMU (`a:off/@y`). */
	top: number
	/** Width in EMU (`a:ext/@cx`); must be positive. */
	width: number
	/** Height in EMU (`a:ext/@cy`); must be positive. */
	height: number
	/** Initial text; omitted/empty yields an empty paragraph. */
	text?: string
	/** Shape name (`p:cNvPr/@name`); defaults to `TextBox <id>`. */
	name?: string
}

/** Options for {@link Slide.addPicture}. Geometry is in EMU. */
export interface AddPictureOptions {
	/** Left edge in EMU (`a:off/@x`). */
	left: number
	/** Top edge in EMU (`a:off/@y`). */
	top: number
	/** Width in EMU (`a:ext/@cx`); must be positive. */
	width: number
	/** Height in EMU (`a:ext/@cy`); must be positive. */
	height: number
	/** Shape name (`p:cNvPr/@name`); defaults to `Picture <id>`. */
	name?: string
	/**
	 * Image file extension (e.g. `png`). When omitted it is sniffed from the
	 * image's magic bytes; supply it (with `contentType`) for formats the
	 * sniffer does not recognize.
	 */
	extension?: string
	/** MIME content type (e.g. `image/png`); sniffed alongside `extension` when omitted. */
	contentType?: string
}

interface ImageType {
	extension: string
	contentType: string
}

/** Recognize a handful of common raster formats from their leading bytes. */
function sniffImageType(bytes: Uint8Array): ImageType | null {
	const b = bytes
	if (b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { extension: 'png', contentType: 'image/png' }
	if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return { extension: 'jpeg', contentType: 'image/jpeg' }
	if (b.length >= 4 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return { extension: 'gif', contentType: 'image/gif' }
	if (b.length >= 2 && b[0] === 0x42 && b[1] === 0x4d) return { extension: 'bmp', contentType: 'image/bmp' }
	if (b.length >= 4 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0x00) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0x00 && b[3] === 0x2a)))
		return { extension: 'tiff', contentType: 'image/tiff' }
	if (b.length >= 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50)
		return { extension: 'webp', contentType: 'image/webp' }
	return null
}

export class Slide {
	constructor(
		readonly presentation: Presentation,
		/** The slide's OPC part (`/ppt/slides/slideN.xml`). */
		readonly part: Part,
		/** The slide id from `p:sldIdLst` (`p:sldId/@id`). */
		readonly slideId: number,
		/** Zero-based position in presentation order. */
		readonly index: number
	) {}

	/** Partname of this slide's part. */
	get partName(): string {
		return this.part.partName
	}

	/** This slide part's relationships (image embeds, layout, hyperlinks, …). */
	get relationships(): Relationships {
		return this.presentation.opc.relationshipsFor(this.partName)
	}

	#themeColors?: FlattenContext

	/**
	 * The slide's resolved theme context (`clrMap` + `clrScheme` + the theme
	 * `fmtScheme` + `fontScheme`, plus the layout/master roots), walked once from
	 * slide → layout → master → theme and cached on this proxy. Backs the read-model
	 * resolved getters ({@link Shape.resolvedFill}, `Run.resolvedColor`,
	 * `Run.resolvedSizePt`, `Run.resolvedFontFace`) so a `schemeClr` token — including
	 * one delivered through a shape's `p:style` `fillRef`/`lnRef` — resolves to a
	 * literal hex, and a placeholder run's inherited size/typeface (a `+mj-*`/`+mn-*`
	 * font token included) resolves to a literal value. The maps/roots are empty when
	 * the theme chain is incomplete, in which case tokens simply stay unresolved.
	 */
	themeContext(): FlattenContext {
		return (this.#themeColors ??= resolveSlideColorContext(this.presentation.opc, this.partName))
	}

	/** Authoring name of the slide (`p:cSld/@name`), or `null` if unnamed. */
	get name(): string | null {
		const cSld = this.#cSld()
		return cSld ? attr(cSld, 'name') : null
	}

	/**
	 * Whether this slide is hidden (`p:sld/@show="0"`). The attribute is
	 * `xsd:boolean` defaulting to `true`, so an absent attribute means shown.
	 * Hidden slides are dropped from PowerPoint/LibreOffice presentations and
	 * exported PDFs, so render order diverges from model order when any earlier
	 * slide is hidden.
	 */
	get hidden(): boolean {
		const root = this.part.dom.documentElement
		const show = root ? attr(root, 'show') : null
		return show === '0' || show === 'false'
	}

	/**
	 * Hide or show this slide. Hiding writes `p:sld/@show="0"`; showing removes
	 * the attribute, restoring PowerPoint's canonical shown form (absent ⇒ shown).
	 * Marks the slide part dirty.
	 */
	set hidden(value: boolean) {
		const root = this.part.dom.documentElement
		if (!root) throw new Error(`Slide ${this.partName} has no root <p:sld> element`)
		if (value) setAttr(root, 'show', '0')
		else removeAttr(root, 'show')
		this.part.markDirty()
	}

	/**
	 * The slide's show transition (`p:transition`), decoded into a typed model, or
	 * `null` when the slide has none. Handles both PowerPoint forms: the bare
	 * `<p:transition>` and the `mc:AlternateContent` wrapper that carries the exact
	 * `p14:dur` duration (the `p14` Choice is preferred so `durationMs` is recovered).
	 */
	get transition(): TransitionInfo | null {
		const root = this.part.dom.documentElement
		return root ? parseTransition(root) : null
	}

	/**
	 * Set or clear the slide's show transition. Assigning `null` removes it.
	 * Writing a transition with `durationMs` emits the `mc:AlternateContent` form
	 * (a `p14` Choice carrying `p14:dur` plus a base `mc:Fallback`); otherwise the
	 * bare `<p:transition>` is written. The node is inserted at its schema slot —
	 * after `p:clrMapOvr`, before `p:timing`/`p:extLst`. Marks the slide part dirty.
	 */
	set transition(value: TransitionInput | null) {
		const root = this.part.dom.documentElement
		if (!root) throw new Error(`Slide ${this.partName} has no root <p:sld> element`)
		removeTransition(root)
		if (value) {
			const doc = this.part.dom
			insertInOrder(root, buildTransition(doc, value), ['p:timing', 'p:extLst'])
		}
		this.part.markDirty()
	}

	/**
	 * Whether the slide carries build animations (`p:timing` with a `<p:bldP>` or a
	 * `presetID`-bearing time node). The animation tree itself is preserved opaquely;
	 * see {@link animationSpids}.
	 */
	get hasAnimations(): boolean {
		const root = this.part.dom.documentElement
		return root ? hasAnimations(root) : false
	}

	/**
	 * @internal The sorted, de-duplicated set of shape ids (`spid`) referenced by
	 * the slide's animations (`<p:spTgt spid>` and `<p:bldP spid>`). Exposed for the
	 * import paths' spid-coherence checks.
	 */
	animationSpids(): number[] {
		const root = this.part.dom.documentElement
		return root ? enumerateSpids(root) : []
	}

	/**
	 * @internal Rewrite every animation `spid` per `mapping` (old → new), keeping
	 * the opaque timing tree coherent when shape ids are reassigned. Marks the slide
	 * part dirty only when a reference changed.
	 */
	remapAnimationSpids(mapping: Map<number, number>): void {
		const root = this.part.dom.documentElement
		if (root && remapSpids(root, mapping)) this.part.markDirty()
	}

	/**
	 * @internal Remove the build animations targeting the given shape ids (their
	 * `<p:bldP>` and effect nodes), so deleting a shape never leaves a dangling
	 * `spid` reference. Marks the slide part dirty only when something was removed.
	 */
	pruneAnimationSpids(spids: Iterable<number>): void {
		const root = this.part.dom.documentElement
		if (root && pruneSpids(root, spids)) this.part.markDirty()
	}

	/** Top-level shapes in the slide's shape tree, in document order. */
	get shapes(): Shape[] {
		const spTree = this.#spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	/** The first top-level shape with the given drawing id (`p:cNvPr/@id`), or `undefined`. */
	shapeById(id: number): Shape | undefined {
		return this.shapes.find((shape) => shape.id === id)
	}

	/** The first top-level shape with the given name (`p:cNvPr/@name`), or `undefined`. */
	shapeByName(name: string): Shape | undefined {
		return this.shapes.find((shape) => shape.name === name)
	}

	/**
	 * The first placeholder of the given type (`p:ph/@type`, e.g. `title`,
	 * `ctrTitle`, `subTitle`, `body`), optionally narrowed by `idx`. Returns
	 * `undefined` when none match. Only `p:sp` shapes can be placeholders, so the
	 * result is an {@link AutoShape}.
	 */
	placeholder(type: string, idx?: string): AutoShape | undefined {
		return this.shapes.find((shape): shape is AutoShape => {
			const ph = shape instanceof AutoShape ? shape.placeholder : null
			return ph !== null && ph.type === type && (idx === undefined || ph.idx === idx)
		})
	}

	/**
	 * Append a text box (`p:sp` with `txBox="1"`) to the slide's shape tree and
	 * return it. Geometry is required (EMU); width and height must be positive.
	 * Allocates a drawing id unique within the slide. Marks the slide part dirty.
	 */
	addTextBox(options: AddTextBoxOptions): AutoShape {
		const { left, top, width, height } = options
		requireFinite(left, 'left')
		requireFinite(top, 'top')
		requirePositive(width, 'width')
		requirePositive(height, 'height')

		const spTree = this.#spTree()
		if (!spTree) throw new Error(`Slide ${this.partName} has no spTree to add a shape to`)
		const doc = spTree.ownerDocument
		if (!doc) throw new Error('Slide DOM has no owner document')

		const id = this.#nextShapeId()
		const sp = buildTextBox(doc, {
			id,
			name: options.name ?? `TextBox ${id}`,
			text: options.text ?? '',
			left: Math.round(left),
			top: Math.round(top),
			width: Math.round(width),
			height: Math.round(height),
		})
		this.#appendShape(spTree, sp)
		return new AutoShape(sp, this)
	}

	/**
	 * Add a picture (`p:pic`) from raw image bytes and return it. Creates a media
	 * part under `/ppt/media/`, registers its content type, wires an `image`
	 * relationship from this slide, and appends the picture to the shape tree.
	 * Geometry is required (EMU); width and height must be positive. The image
	 * format is sniffed from the bytes unless `extension`/`contentType` are given.
	 */
	addPicture(image: Uint8Array, options: AddPictureOptions): Picture {
		const { left, top, width, height } = options
		requireFinite(left, 'left')
		requireFinite(top, 'top')
		requirePositive(width, 'width')
		requirePositive(height, 'height')

		const sniffed = sniffImageType(image)
		const extension = (options.extension ?? sniffed?.extension)?.toLowerCase().replace(/^\./, '')
		const contentType = options.contentType ?? sniffed?.contentType
		if (!extension || !contentType) {
			throw new Error('Could not determine image type; pass { extension, contentType } to addPicture')
		}

		const spTree = this.#spTree()
		if (!spTree) throw new Error(`Slide ${this.partName} has no spTree to add a picture to`)
		const doc = spTree.ownerDocument
		if (!doc) throw new Error('Slide DOM has no owner document')

		const opc = this.presentation.opc
		const mediaPartName = opc.reserveMediaPartName(extension)
		opc.addPart(mediaPartName, contentType, image)
		const relId = this.relationships.add(IMAGE_REL_TYPE, relativePartName(this.partName, mediaPartName)).id

		const id = this.#nextShapeId()
		const pic = buildPicture(doc, {
			id,
			name: options.name ?? `Picture ${id}`,
			relId,
			left: Math.round(left),
			top: Math.round(top),
			width: Math.round(width),
			height: Math.round(height),
		})
		this.#appendShape(spTree, pic)
		return new Picture(pic, this)
	}

	/**
	 * @internal The slide's shape tree (`p:cSld/p:spTree`), or `null` if absent.
	 * Exposed for cross-slide composition (`Presentation.importShape`).
	 */
	shapeTree(): Element | null {
		return this.#spTree()
	}

	/**
	 * @internal The smallest drawing id (`p:cNvPr/@id`) not already used on this
	 * slide. Exposed so `Presentation.importShape` can give a lifted shape (and its
	 * group children) collision-free ids on the host.
	 */
	nextShapeId(): number {
		return this.#nextShapeId()
	}

	/** Insert a shape after grpSpPr and before any trailing p:extLst on the tree; mark dirty. */
	#appendShape(spTree: Element, shape: Element): void {
		spTree.insertBefore(shape, firstChild(spTree, 'p:extLst'))
		this.part.markDirty()
	}

	/** The smallest drawing id (`p:cNvPr/@id`) not already used on the slide. */
	#nextShapeId(): number {
		const root = this.part.dom.documentElement
		let max = 1
		if (root) {
			const cNvPrs = root.getElementsByTagNameNS(OOXML_NS.p, 'cNvPr')
			for (let i = 0; i < cNvPrs.length; i++) {
				const id = intValue(attr(cNvPrs[i], 'id'))
				if (id !== null && id > max) max = id
			}
		}
		return max + 1
	}

	#cSld(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'p:cSld') : null
	}

	#spTree(): Element | null {
		const cSld = this.#cSld()
		return cSld ? firstChild(cSld, 'p:spTree') : null
	}
}

function requireFinite(value: number, name: string): void {
	if (!Number.isFinite(value)) throw new Error(`${name} must be a finite number of EMU, got ${value}`)
}

function requirePositive(value: number, name: string): void {
	requireFinite(value, name)
	if (value <= 0) throw new Error(`${name} must be positive, got ${value}`)
}

interface TextBoxSpec {
	id: number
	name: string
	text: string
	left: number
	top: number
	width: number
	height: number
}

/** Build a minimal, schema-valid text-box `p:sp` element (not yet attached). */
function buildTextBox(doc: Document, spec: TextBoxSpec): Element {
	const make = (qname: string): Element => createElement(doc, qname)
	const append = (parent: Element, qname: string): Element => {
		const child = make(qname)
		parent.appendChild(child)
		return child
	}

	const sp = make('p:sp')

	const nvSpPr = append(sp, 'p:nvSpPr')
	const cNvPr = append(nvSpPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvSpPr = append(nvSpPr, 'p:cNvSpPr')
	setAttr(cNvSpPr, 'txBox', '1')
	append(nvSpPr, 'p:nvPr')

	const spPr = append(sp, 'p:spPr')
	const xfrm = append(spPr, 'a:xfrm')
	const off = append(xfrm, 'a:off')
	setAttr(off, 'x', String(spec.left))
	setAttr(off, 'y', String(spec.top))
	const ext = append(xfrm, 'a:ext')
	setAttr(ext, 'cx', String(spec.width))
	setAttr(ext, 'cy', String(spec.height))
	const prstGeom = append(spPr, 'a:prstGeom')
	setAttr(prstGeom, 'prst', 'rect')
	append(prstGeom, 'a:avLst')

	const txBody = append(sp, 'p:txBody')
	append(txBody, 'a:bodyPr')
	append(txBody, 'a:lstStyle')
	const p = append(txBody, 'a:p')
	if (spec.text !== '') {
		const r = append(p, 'a:r')
		const t = append(r, 'a:t')
		t.textContent = spec.text
		if (spec.text !== spec.text.trim()) setAttr(t, 'xml:space', 'preserve')
	}

	return sp
}

interface PictureSpec {
	id: number
	name: string
	relId: string
	left: number
	top: number
	width: number
	height: number
}

/** Build a minimal, schema-valid `p:pic` element (not yet attached). */
function buildPicture(doc: Document, spec: PictureSpec): Element {
	const make = (qname: string): Element => createElement(doc, qname)
	const append = (parent: Element, qname: string): Element => {
		const child = make(qname)
		parent.appendChild(child)
		return child
	}

	const pic = make('p:pic')

	const nvPicPr = append(pic, 'p:nvPicPr')
	const cNvPr = append(nvPicPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvPicPr = append(nvPicPr, 'p:cNvPicPr')
	const picLocks = append(cNvPicPr, 'a:picLocks')
	setAttr(picLocks, 'noChangeAspect', '1')
	append(nvPicPr, 'p:nvPr')

	const blipFill = append(pic, 'p:blipFill')
	const blip = append(blipFill, 'a:blip')
	setAttr(blip, 'r:embed', spec.relId)
	const stretch = append(blipFill, 'a:stretch')
	append(stretch, 'a:fillRect')

	const spPr = append(pic, 'p:spPr')
	const xfrm = append(spPr, 'a:xfrm')
	const off = append(xfrm, 'a:off')
	setAttr(off, 'x', String(spec.left))
	setAttr(off, 'y', String(spec.top))
	const ext = append(xfrm, 'a:ext')
	setAttr(ext, 'cx', String(spec.width))
	setAttr(ext, 'cy', String(spec.height))
	const prstGeom = append(spPr, 'a:prstGeom')
	setAttr(prstGeom, 'prst', 'rect')
	append(prstGeom, 'a:avLst')

	return pic
}
