/**
 * Read-model proxies for the shapes on a slide.
 *
 * The `p:spTree` holds five shape kinds; each wraps its element and exposes
 * non-visual identity (id/name), geometry (left/top/width/height in EMU), and
 * kind-specific reads. Proxies hold a back-reference to the owning `Slide` so
 * pictures can resolve their image relationship and so future edits can mark
 * the slide part dirty.
 */
import {
	ELEMENT_NODE,
	OOXML_NS,
	attr,
	firstChild,
	getOrAddChild,
	intValue,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { FILL_CHOICES, normalizeHex, setSolidFill, solidFillColor } from '../oxml/fill.js'
import { Chart } from './chart.js'
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

// Successor arrays for inserting a fill / line *into* a properties element.
// Distinct from the *_AFTER_XFRM arrays above, which sequence a:xfrm (the first
// child): a:solidFill and a:ln sit mid-sequence, so their `before` lists must
// contain only the children that legally follow them (CT_ShapeProperties /
// CT_GroupShapeProperties / CT_LineProperties).
const SPPR_FILL_AFTER = ['a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const SPPR_LN_AFTER = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const GRPSPPR_FILL_AFTER = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:extLst']
const LN_FILL_AFTER = ['a:prstDash', 'a:custDash', 'a:round', 'a:bevel', 'a:miter', 'a:headEnd', 'a:tailEnd', 'a:extLst']

/**
 * A shape's properties element (`p:spPr` / `p:grpSpPr`) paired with the schema
 * successors for ordered insertion of its `a:solidFill` and `a:ln` children.
 * `lnAfter` is `null` for kinds with no `a:ln` (group shapes).
 */
interface ShapeProperties {
	props: Element
	fillAfter: string[]
	lnAfter: string[] | null
}

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

	/** The shape's properties element (`p:spPr` / `p:grpSpPr`), or `null` when absent. */
	protected properties(): Element | null {
		return firstChild(this.element, 'p:spPr')
	}

	/** Get-or-add the properties element in document order, with the successor
	 *  arrays for inserting its fill / line children. Subclasses override this
	 *  to point at `p:grpSpPr`, or to reject kinds with no properties element. */
	protected getOrAddProperties(): ShapeProperties {
		const props = getOrAddChild(this.element, 'p:spPr', SHAPE_AFTER_SPPR)
		return { props, fillAfter: SPPR_FILL_AFTER, lnAfter: SPPR_LN_AFTER }
	}

	/** Whether a solid fill can be set on this shape kind. Pictures and graphic
	 *  frames opt out (they carry their own image / table-cell fill model). */
	protected get supportsFill(): boolean {
		return true
	}

	/** Explicit RGB fill colour as a 6-hex string (`spPr/a:solidFill/a:srgbClr/@val`), or `null`. */
	get fillColor(): string | null {
		return solidFillColor(this.properties(), 'a:srgbClr')
	}

	set fillColor(value: string | null) {
		this.#setFill(value === null ? null : { qname: 'a:srgbClr', val: normalizeHex(value) })
	}

	/** Theme colour token when the fill is a scheme colour (`a:solidFill/a:schemeClr/@val`, e.g. `accent2`), or `null`. */
	get fillSchemeColor(): string | null {
		return solidFillColor(this.properties(), 'a:schemeClr')
	}

	set fillSchemeColor(value: string | null) {
		this.#setFill(value === null ? null : { qname: 'a:schemeClr', val: value })
	}

	/**
	 * Set an explicit `<a:noFill/>` on the shape — a transparent surface. This is
	 * distinct from clearing the fill (`fillColor = null`), which removes the
	 * `a:solidFill` and lets the fill inherit from the shape's style/placeholder.
	 */
	noFill(): void {
		if (!this.supportsFill) throw new Error(`${this.shapeType} shapes do not support a solid fill`)
		const { props, fillAfter } = this.getOrAddProperties()
		removeChildrenByQName(props, FILL_CHOICES)
		getOrAddChild(props, 'a:noFill', fillAfter)
		this.markDirty()
	}

	/** Explicit RGB line/border colour (`spPr/a:ln/a:solidFill/a:srgbClr/@val`), or `null`. */
	get lineColor(): string | null {
		return solidFillColor(this.#line(), 'a:srgbClr')
	}

	set lineColor(value: string | null) {
		this.#setLine(value === null ? null : { qname: 'a:srgbClr', val: normalizeHex(value) })
	}

	/** Theme colour token when the line is a scheme colour (`a:ln/a:solidFill/a:schemeClr/@val`), or `null`. */
	get lineSchemeColor(): string | null {
		return solidFillColor(this.#line(), 'a:schemeClr')
	}

	set lineSchemeColor(value: string | null) {
		this.#setLine(value === null ? null : { qname: 'a:schemeClr', val: value })
	}

	/** The line element (`spPr/a:ln`), or `null` when absent. */
	#line(): Element | null {
		const props = this.properties()
		return props ? firstChild(props, 'a:ln') : null
	}

	#setFill(color: { qname: string; val: string } | null): void {
		if (color === null) {
			const props = this.properties()
			if (!props || !firstChild(props, 'a:solidFill')) return
			removeChildrenByQName(props, ['a:solidFill'])
			this.markDirty()
			return
		}
		if (!this.supportsFill) throw new Error(`${this.shapeType} shapes do not support a solid fill`)
		const { props, fillAfter } = this.getOrAddProperties()
		setSolidFill(props, fillAfter, color)
		this.markDirty()
	}

	#setLine(color: { qname: string; val: string } | null): void {
		if (color === null) {
			const ln = this.#line()
			if (!ln || !firstChild(ln, 'a:solidFill')) return
			removeChildrenByQName(ln, ['a:solidFill'])
			this.markDirty()
			return
		}
		const { props, lnAfter } = this.getOrAddProperties()
		if (lnAfter === null) throw new Error(`${this.shapeType} shapes do not support a line colour`)
		const ln = getOrAddChild(props, 'a:ln', lnAfter)
		setSolidFill(ln, LN_FILL_AFTER, color)
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

	// A picture's image is its sibling `p:blipFill`, not a fill of `p:spPr`, so a
	// solid `spPr` fill would not clobber the image. v1 still omits fill setters
	// here — recolouring a picture surface is rarely what a caller means — and
	// exposes only the border via `lineColor`. Reads of `fillColor` stay valid.
	protected override get supportsFill(): boolean {
		return false
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

	// A graphicFrame has no p:spPr; its hosted table/chart carries its own fill
	// model. There is nothing to get-or-add, so fill and line setters reject it.
	protected override getOrAddProperties(): ShapeProperties {
		throw new Error('graphicFrame shapes have no shape properties; fill and line colours are not supported')
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

	/** The hosted chart, or `null` when this frame is not a chart or its part is missing. */
	get chart(): Chart | null {
		if (!this.hasChart) return null
		const graphicData = this.#graphicData()
		const chartRef = graphicData && firstChild(graphicData, 'c:chart')
		const relId = chartRef && attr(chartRef, 'r:id')
		if (!relId) return null
		const partName = this.slide.relationships.resolveTarget(relId)
		const part = this.slide.presentation.opc.part(partName)
		return part ? new Chart(part) : null
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
		return getOrAddChild(this.#getOrAddGrpSpPr(), 'a:xfrm', GRPSPPR_AFTER_XFRM)
	}

	protected override properties(): Element | null {
		return firstChild(this.element, 'p:grpSpPr')
	}

	// A group's fill lives in p:grpSpPr, which has no a:ln (no line colour).
	protected override getOrAddProperties(): ShapeProperties {
		return { props: this.#getOrAddGrpSpPr(), fillAfter: GRPSPPR_FILL_AFTER, lnAfter: null }
	}

	#getOrAddGrpSpPr(): Element {
		return getOrAddChild(this.element, 'p:grpSpPr', ['p:sp', 'p:grpSp', 'p:pic', 'p:cxnSp', 'p:graphicFrame'])
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
