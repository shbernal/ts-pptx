/**
 * Read-model proxies for the shapes on a slide.
 *
 * The `p:spTree` holds five shape kinds; each wraps its element and exposes
 * non-visual identity (id/name), geometry (left/top/width/height in EMU), and
 * kind-specific reads. Proxies hold a back-reference to the owning `Slide` so
 * pictures can resolve their image relationship and so future edits can mark
 * the slide part dirty.
 */
import { ELEMENT_NODE, OOXML_NS, attr, firstChild, getOrAddChild, intValue, setAttr, type Element } from '../oxml/dom.js'
import { Table } from './table.js'
import { TextFrame } from './text.js'
import type { Slide } from './slide.js'

const A_TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'
const A_CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

// Schema successors used to keep elements in document order when a geometry
// setter has to create one.
const SPPR_AFTER_XFRM = [
	'a:custGeom',
	'a:prstGeom',
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
	'a:ln',
	'a:effectLst',
	'a:effectDag',
	'a:scene3d',
	'a:sp3d',
	'a:extLst',
]
const GRPSPPR_AFTER_XFRM = [
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
	'a:effectLst',
	'a:effectDag',
	'a:scene3d',
	'a:extLst',
]
// spPr itself sits before p:style / p:txBody within p:sp (and before p:style
// within p:pic / p:cxnSp); blipFill / nv*Pr precede it and are excluded.
const SHAPE_AFTER_SPPR = ['p:style', 'p:txBody']

/** Discriminator for the concrete `Shape` subclass. */
export type ShapeType = 'autoShape' | 'picture' | 'connector' | 'graphicFrame' | 'group'

/** First `<p:cNvPr>` reached through the shape's non-visual properties wrapper (`p:nv*Pr`). */
function nonVisualCNvPr(element: Element): Element | null {
	for (let node = element.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const child = node as Element
		if (child.namespaceURI === OOXML_NS.p && child.localName.startsWith('nv')) {
			return firstChild(child, 'p:cNvPr')
		}
	}
	return null
}

function emuFrom(parent: Element | null, qname: string, attribute: string): number | null {
	const element = parent && firstChild(parent, qname)
	return element ? intValue(attr(element, attribute)) : null
}

/** Validate and round an EMU geometry value; extents (`cx`/`cy`) must be non-negative. */
function toEmu(value: number, attribute: string, allowNegative: boolean): number {
	if (!Number.isFinite(value)) throw new Error(`${attribute} must be a finite number of EMU, got ${value}`)
	if (!allowNegative && value < 0) throw new Error(`${attribute} must be non-negative, got ${value}`)
	return Math.round(value)
}

/** Common base for every shape in a slide's shape tree. */
export abstract class Shape {
	constructor(
		protected readonly element: Element,
		readonly slide: Slide
	) {}

	/** Which concrete shape kind this is. */
	abstract readonly shapeType: ShapeType

	/** The transform element (`a:xfrm` or `p:xfrm`) carrying this shape's geometry, or `null` if inherited. */
	protected abstract xfrm(): Element | null

	/** The transform element, creating it (and its container) in document order if absent. */
	protected abstract getOrAddXfrm(): Element

	/** Mark the owning slide part dirty so `save()` reserializes it. */
	protected markDirty(): void {
		this.slide.part.markDirty()
	}

	/** Drawing id (`p:cNvPr/@id`), or `null` if absent. */
	get id(): number | null {
		const cNvPr = nonVisualCNvPr(this.element)
		return cNvPr ? intValue(attr(cNvPr, 'id')) : null
	}

	/** Shape name (`p:cNvPr/@name`), or `''` if unnamed. */
	get name(): string {
		const cNvPr = nonVisualCNvPr(this.element)
		return (cNvPr && attr(cNvPr, 'name')) ?? ''
	}

	/** Left edge in EMU (`a:off/@x`), or `null` when the shape has no own transform. */
	get left(): number | null {
		return emuFrom(this.xfrm(), 'a:off', 'x')
	}

	set left(value: number) {
		this.#setOffset('x', value, true)
	}

	/** Top edge in EMU (`a:off/@y`), or `null` when the shape has no own transform. */
	get top(): number | null {
		return emuFrom(this.xfrm(), 'a:off', 'y')
	}

	set top(value: number) {
		this.#setOffset('y', value, true)
	}

	/** Width in EMU (`a:ext/@cx`), or `null` when the shape has no own transform. */
	get width(): number | null {
		return emuFrom(this.xfrm(), 'a:ext', 'cx')
	}

	set width(value: number) {
		this.#setExtent('cx', value)
	}

	/** Height in EMU (`a:ext/@cy`), or `null` when the shape has no own transform. */
	get height(): number | null {
		return emuFrom(this.xfrm(), 'a:ext', 'cy')
	}

	set height(value: number) {
		this.#setExtent('cy', value)
	}

	#setOffset(axis: 'x' | 'y', value: number, allowNegative: boolean): void {
		const emu = toEmu(value, axis, allowNegative)
		const off = getOrAddChild(this.getOrAddXfrm(), 'a:off', ['a:ext'])
		setAttr(off, axis, String(emu))
		this.markDirty()
	}

	#setExtent(axis: 'cx' | 'cy', value: number): void {
		const emu = toEmu(value, axis, false)
		const ext = getOrAddChild(this.getOrAddXfrm(), 'a:ext')
		setAttr(ext, axis, String(emu))
		this.markDirty()
	}

	/** Whether this shape can hold text (only `p:sp` does in this read model). */
	get hasTextFrame(): boolean {
		return false
	}

	/** The shape's text frame, or `null` when it cannot hold text. */
	get textFrame(): TextFrame | null {
		return null
	}

	/** Convenience: the shape's full text, or `''` if it has none. */
	get text(): string {
		return this.textFrame?.text ?? ''
	}

	/**
	 * Remove this shape from its parent (the slide's shape tree, or an enclosing
	 * group) and mark the owning slide part dirty. The proxy is dead afterwards.
	 */
	delete(): void {
		const parent = this.element.parentNode
		if (!parent) throw new Error('Shape is not attached to a parent and cannot be deleted')
		parent.removeChild(this.element)
		this.markDirty()
	}

	/** The underlying shape element, for advanced reads and future mutation. */
	get element_(): Element {
		return this.element
	}
}

/** An auto shape, text box, or placeholder (`p:sp`). The only kind that holds text. */
export class AutoShape extends Shape {
	readonly shapeType = 'autoShape' as const

	protected xfrm(): Element | null {
		const spPr = firstChild(this.element, 'p:spPr')
		return spPr ? firstChild(spPr, 'a:xfrm') : null
	}

	protected getOrAddXfrm(): Element {
		return getOrAddSpPrXfrm(this.element)
	}

	override get hasTextFrame(): boolean {
		return firstChild(this.element, 'p:txBody') !== null
	}

	override get textFrame(): TextFrame | null {
		const txBody = firstChild(this.element, 'p:txBody')
		return txBody ? new TextFrame(txBody, this.slide.part) : null
	}

	/** Preset geometry name (`a:prstGeom/@prst`, e.g. `rect`), or `null` for custom/none. */
	get presetGeometry(): string | null {
		const spPr = firstChild(this.element, 'p:spPr')
		const prstGeom = spPr && firstChild(spPr, 'a:prstGeom')
		return prstGeom ? attr(prstGeom, 'prst') : null
	}
}

/** A picture (`p:pic`). */
export class Picture extends Shape {
	readonly shapeType = 'picture' as const

	protected xfrm(): Element | null {
		const spPr = firstChild(this.element, 'p:spPr')
		return spPr ? firstChild(spPr, 'a:xfrm') : null
	}

	protected getOrAddXfrm(): Element {
		return getOrAddSpPrXfrm(this.element)
	}

	/** Relationship id of the embedded image (`p:blipFill/a:blip/@r:embed`), or `null`. */
	get imageRelId(): string | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const blip = blipFill && firstChild(blipFill, 'a:blip')
		return blip ? attr(blip, 'r:embed') : null
	}

	/** Absolute partname of the embedded image, resolved via the slide's relationships, or `null`. */
	get imagePartName(): string | null {
		const relId = this.imageRelId
		return relId ? this.slide.relationships.resolveTarget(relId) : null
	}
}

/** A connector / line (`p:cxnSp`). */
export class Connector extends Shape {
	readonly shapeType = 'connector' as const

	protected xfrm(): Element | null {
		const spPr = firstChild(this.element, 'p:spPr')
		return spPr ? firstChild(spPr, 'a:xfrm') : null
	}

	protected getOrAddXfrm(): Element {
		return getOrAddSpPrXfrm(this.element)
	}
}

/** A graphic frame (`p:graphicFrame`) — host for tables and charts. */
export class GraphicFrame extends Shape {
	readonly shapeType = 'graphicFrame' as const

	protected xfrm(): Element | null {
		// graphicFrame carries its own `p:xfrm` directly, not inside `p:spPr`.
		return firstChild(this.element, 'p:xfrm')
	}

	protected getOrAddXfrm(): Element {
		// p:xfrm sits between p:nvGraphicFramePr and a:graphic.
		return getOrAddChild(this.element, 'p:xfrm', ['a:graphic', 'p:extLst'])
	}

	/** Whether this frame hosts a table (`a:graphicData/@uri` is the table URI). */
	get hasTable(): boolean {
		return this.#graphicDataUri() === A_TABLE_URI
	}

	/** Whether this frame hosts a chart (`a:graphicData/@uri` is the chart URI). */
	get hasChart(): boolean {
		return this.#graphicDataUri() === A_CHART_URI
	}

	/** The hosted table, or `null` when this frame is not a table. */
	get table(): Table | null {
		if (!this.hasTable) return null
		const graphicData = this.#graphicData()
		const tbl = graphicData && firstChild(graphicData, 'a:tbl')
		return tbl ? new Table(tbl, this.slide.part) : null
	}

	#graphicData(): Element | null {
		const graphic = firstChild(this.element, 'a:graphic')
		return graphic ? firstChild(graphic, 'a:graphicData') : null
	}

	#graphicDataUri(): string | null {
		const graphicData = this.#graphicData()
		return graphicData ? attr(graphicData, 'uri') : null
	}
}

/** A group shape (`p:grpSp`) — contains nested shapes. */
export class GroupShape extends Shape {
	readonly shapeType = 'group' as const

	protected xfrm(): Element | null {
		const grpSpPr = firstChild(this.element, 'p:grpSpPr')
		return grpSpPr ? firstChild(grpSpPr, 'a:xfrm') : null
	}

	protected getOrAddXfrm(): Element {
		const grpSpPr = getOrAddChild(this.element, 'p:grpSpPr', ['p:sp', 'p:grpSp', 'p:pic', 'p:cxnSp', 'p:graphicFrame'])
		return getOrAddChild(grpSpPr, 'a:xfrm', GRPSPPR_AFTER_XFRM)
	}

	/** The shapes nested directly inside this group, in document order. */
	get shapes(): Shape[] {
		return buildShapes(this.element, this.slide)
	}
}

/** Get-or-add `p:spPr/a:xfrm` for shapes whose transform lives in `p:spPr` (`p:sp`, `p:pic`, `p:cxnSp`). */
function getOrAddSpPrXfrm(shapeElement: Element): Element {
	const spPr = getOrAddChild(shapeElement, 'p:spPr', SHAPE_AFTER_SPPR)
	return getOrAddChild(spPr, 'a:xfrm', SPPR_AFTER_XFRM)
}

/**
 * Build shape proxies for the shape-tree children of `parent` (a `p:spTree` or
 * `p:grpSp`), skipping non-shape children (`p:nvGrpSpPr`, `p:grpSpPr`, …).
 */
export function buildShapes(parent: Element, slide: Slide): Shape[] {
	const shapes: Shape[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const element = node as Element
		if (element.namespaceURI !== OOXML_NS.p) continue
		switch (element.localName) {
			case 'sp':
				shapes.push(new AutoShape(element, slide))
				break
			case 'pic':
				shapes.push(new Picture(element, slide))
				break
			case 'cxnSp':
				shapes.push(new Connector(element, slide))
				break
			case 'graphicFrame':
				shapes.push(new GraphicFrame(element, slide))
				break
			case 'grpSp':
				shapes.push(new GroupShape(element, slide))
				break
			default:
				break
		}
	}
	return shapes
}
