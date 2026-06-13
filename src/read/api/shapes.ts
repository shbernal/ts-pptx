/**
 * Read-model proxies for the shapes on a slide.
 *
 * The `p:spTree` holds five shape kinds; each wraps its element and exposes
 * non-visual identity (id/name), geometry (left/top/width/height in EMU), and
 * kind-specific reads. Proxies hold a back-reference to the owning `Slide` so
 * pictures can resolve their image relationship and so future edits can mark
 * the slide part dirty.
 */
import { ELEMENT_NODE, OOXML_NS, attr, firstChild, intValue, type Element } from '../oxml/dom.js'
import { TextFrame } from './text.js'
import type { Slide } from './slide.js'

const A_TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'
const A_CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

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

	/** Top edge in EMU (`a:off/@y`), or `null` when the shape has no own transform. */
	get top(): number | null {
		return emuFrom(this.xfrm(), 'a:off', 'y')
	}

	/** Width in EMU (`a:ext/@cx`), or `null` when the shape has no own transform. */
	get width(): number | null {
		return emuFrom(this.xfrm(), 'a:ext', 'cx')
	}

	/** Height in EMU (`a:ext/@cy`), or `null` when the shape has no own transform. */
	get height(): number | null {
		return emuFrom(this.xfrm(), 'a:ext', 'cy')
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

	override get hasTextFrame(): boolean {
		return firstChild(this.element, 'p:txBody') !== null
	}

	override get textFrame(): TextFrame | null {
		const txBody = firstChild(this.element, 'p:txBody')
		return txBody ? new TextFrame(txBody) : null
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
}

/** A graphic frame (`p:graphicFrame`) — host for tables and charts. */
export class GraphicFrame extends Shape {
	readonly shapeType = 'graphicFrame' as const

	protected xfrm(): Element | null {
		// graphicFrame carries its own `p:xfrm` directly, not inside `p:spPr`.
		return firstChild(this.element, 'p:xfrm')
	}

	/** Whether this frame hosts a table (`a:graphicData/@uri` is the table URI). */
	get hasTable(): boolean {
		return this.#graphicDataUri() === A_TABLE_URI
	}

	/** Whether this frame hosts a chart (`a:graphicData/@uri` is the chart URI). */
	get hasChart(): boolean {
		return this.#graphicDataUri() === A_CHART_URI
	}

	#graphicDataUri(): string | null {
		const graphic = firstChild(this.element, 'a:graphic')
		const graphicData = graphic && firstChild(graphic, 'a:graphicData')
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

	/** The shapes nested directly inside this group, in document order. */
	get shapes(): Shape[] {
		return buildShapes(this.element, this.slide)
	}
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
