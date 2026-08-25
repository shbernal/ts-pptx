/**
 * Read-model proxies for the shapes in a `p:spTree`.
 *
 * The `p:spTree` holds five shape kinds; each wraps its element and exposes
 * non-visual identity (id/name), geometry (left/top/width/height in EMU), and
 * kind-specific reads. Proxies hold a back-reference to the owning
 * {@link ShapeHost} — the slide, layout, or master whose part carries the tree —
 * so pictures can resolve their image relationship and so edits can mark that
 * part dirty.
 *
 * This module is the entry point for that model: the `Shape` base and the four leaf kinds live in
 * `./shapes/`, and only the dispatch that turns an element into a proxy — plus `GroupShape`, which
 * recurses back through it — is here. Everything the read API surfaces is re-exported below, so
 * `ts-pptx/read` and its consumers keep importing shapes from one place.
 */
import { ELEMENT_NODE, OOXML_NS, firstChild, getElements, getOrAddChild, type Element } from '../oxml/dom.js'
import { GRPSPPR_AFTER_XFRM, GRPSPPR_FILL_AFTER, type ShapeProperties } from './shapes/oxml.js'
import { readBox } from './shapes/geometry.js'
import type { ChildFrame } from './shapes/types.js'
import { Shape } from './shapes/base.js'
import { AutoShape } from './shapes/autoshape.js'
import { Picture } from './shapes/picture.js'
import { Connector } from './shapes/connector.js'
import { GraphicFrame } from './shapes/graphic-frame.js'
import type { ShapeHost } from './shapes/host.js'

export { Shape } from './shapes/base.js'
export type { ShapeHost } from './shapes/host.js'
export { AutoShape } from './shapes/autoshape.js'
export { Picture } from './shapes/picture.js'
export { Connector } from './shapes/connector.js'
export { GraphicFrame } from './shapes/graphic-frame.js'

// Re-exported so `ts-pptx/read` keeps surfacing the gradient types from here even
// though their definitions moved to ./gradient.js (shared with the slide-background reader).
export type { GradientStop, GradientFill } from './gradient.js'
export type { PictureFill, PictureFillTile, FillRect } from './picture-fill.js'
// Likewise for the shape value types, whose definitions live in ./shapes/types.js.
export type {
	AbsoluteFrame,
	ChildFrame,
	ConnectionSite,
	CustomGeometry,
	CustomGeometryPath,
	GeometryCommand,
	Glow,
	InnerShadow,
	LineEnd,
	LineEnds,
	OuterShadow,
	PatternFill,
	Recolor,
	RecolorColor,
	Reflection,
	ShapeType,
	SoftEdge,
} from './shapes/types.js'

/**
 * A group shape (`p:grpSp`) — contains nested shapes.
 *
 * It is the one kind defined here rather than under `./shapes/`: its `shapes` getter recurses
 * through `buildShapes` below, so putting it in a sibling module would make the dispatch and the
 * group import each other.
 */
export class GroupShape extends Shape {
	readonly shapeType = 'group' as const

	protected override xfrm(): Element | null {
		const grpSpPr = firstChild(this.element, 'p:grpSpPr')
		return grpSpPr ? firstChild(grpSpPr, 'a:xfrm') : null
	}

	protected override getOrAddXfrm(): Element {
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

	/**
	 * The group's own child coordinate space (`p:grpSpPr/a:xfrm/a:chOff` and
	 * `a:chExt`) in EMU, or `null` when the group has no transform or either
	 * element is incomplete.
	 *
	 * Most consumers should not need this: {@link Shape.absoluteFrame} already
	 * composes the whole ancestor chain — child space, group flips and rotations —
	 * and doing that arithmetic per consumer is how a shape in a nested group ends
	 * up subtly displaced. It is here for a *replica* consumer rather than a
	 * *paint* one: rebuilding this group and its children as OOXML needs the
	 * source child space to reproduce the scaling, and without it only groups whose
	 * child space is the identity can be rebuilt.
	 */
	get childFrame(): ChildFrame | null {
		const xfrm = this.xfrm()
		const box = xfrm && readBox(xfrm, 'a:chOff', 'a:chExt')
		if (!box) return null
		return { offsetX: box.x, offsetY: box.y, extentX: box.cx, extentY: box.cy }
	}

	/** The shapes nested directly inside this group, in document order. */
	get shapes(): AnyShape[] {
		return buildShapes(this.element, this.host)
	}
}

/**
 * The concrete shape kinds as a discriminated union keyed on the
 * {@link Shape.shapeType} literal. The read API (`Slide.shapes`,
 * `Slide.shapeById`, `GroupShape.shapes`, `Presentation.importShape`, …) returns
 * this rather than the abstract {@link Shape} base, so a consumer can narrow to a
 * subtype by its discriminant and reach the subtype-only members — e.g.
 * `if (shape.shapeType === 'graphicFrame') shape.chart` or, equivalently, the
 * {@link isGraphicFrame} guard. The abstract {@link Shape} base is still exported
 * for the "common members only" case and for `instanceof` checks.
 */
export type AnyShape = AutoShape | Picture | Connector | GraphicFrame | GroupShape

/** Narrow a shape to an {@link AutoShape} (`p:sp` — auto shape, text box, or placeholder). */
export function isAutoShape(shape: Shape): shape is AutoShape {
	return shape.shapeType === 'autoShape'
}

/** Narrow a shape to a {@link Picture} (`p:pic`). */
export function isPicture(shape: Shape): shape is Picture {
	return shape.shapeType === 'picture'
}

/** Narrow a shape to a {@link Connector} (`p:cxnSp`). */
export function isConnector(shape: Shape): shape is Connector {
	return shape.shapeType === 'connector'
}

/** Narrow a shape to a {@link GraphicFrame} (`p:graphicFrame` — table/chart host). */
export function isGraphicFrame(shape: Shape): shape is GraphicFrame {
	return shape.shapeType === 'graphicFrame'
}

/** Narrow a shape to a {@link GroupShape} (`p:grpSp`). */
export function isGroupShape(shape: Shape): shape is GroupShape {
	return shape.shapeType === 'group'
}

/**
 * Wrap a single shape-tree element (`p:sp`/`p:pic`/`p:cxnSp`/`p:graphicFrame`/
 * `p:grpSp`) in its concrete `Shape` proxy, or `null` if it is not a shape kind
 * (e.g. `p:nvGrpSpPr`, `p:grpSpPr`, `p:extLst`).
 */
export function wrapShapeElement(element: Element, host: ShapeHost): AnyShape | null {
	if (element.namespaceURI !== OOXML_NS.p) return null
	switch (element.localName) {
		case 'sp':
			return new AutoShape(element, host)
		case 'pic':
			return new Picture(element, host)
		case 'cxnSp':
			return new Connector(element, host)
		case 'graphicFrame':
			return new GraphicFrame(element, host)
		case 'grpSp':
			return new GroupShape(element, host)
		default:
			return null
	}
}

/**
 * The richer of an `mc:AlternateContent`'s two branches: the first `mc:Choice`
 * that carries a shape element (what PowerPoint renders when it understands the
 * required feature), falling back to `mc:Fallback`. PowerPoint wraps a shape this
 * way when its full form needs a feature namespace a plain consumer lacks — a
 * chartEx chart's `p:graphicFrame`, a zoom frame, an inline-math shape. Returns
 * the wrapped inner shape, or `null` when neither branch holds one.
 */
function unwrapAlternateContent(altContent: Element, host: ShapeHost): AnyShape | null {
	const branches = [...getElements(altContent, 'mc:Choice'), ...getElements(altContent, 'mc:Fallback')]
	for (const branch of branches) {
		for (let node = branch.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const shape = wrapShapeElement(node as Element, host)
			if (shape) return shape
		}
	}
	return null
}

/**
 * Build shape proxies for the shape-tree children of `parent` (a `p:spTree` or
 * `p:grpSp`), skipping non-shape children (`p:nvGrpSpPr`, `p:grpSpPr`, …). An
 * `mc:AlternateContent` wrapper (chartEx charts, zoom frames, inline math) is
 * unwrapped to the shape inside its preferred branch.
 */
export function buildShapes(parent: Element, host: ShapeHost): AnyShape[] {
	const shapes: AnyShape[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const element = node as Element
		if (element.namespaceURI === OOXML_NS.mc && element.localName === 'AlternateContent') {
			const unwrapped = unwrapAlternateContent(element, host)
			if (unwrapped) shapes.push(unwrapped)
			continue
		}
		const shape = wrapShapeElement(element, host)
		if (shape) shapes.push(shape)
	}
	return shapes
}

/**
 * The first shape in `shapes` with the given drawing id (`p:cNvPr/@id`), walked
 * pre-order so a group is visited before its children. Backs
 * {@link ShapeHost.shapeByIdDeep} on all three hosts — drawing ids are unique
 * within a part, so the first match is the only match and the walk order just
 * fixes a deterministic one.
 */
export function findShapeByIdDeep(shapes: AnyShape[], id: number): AnyShape | undefined {
	for (const shape of shapes) {
		if (shape.id === id) return shape
		if (shape instanceof GroupShape) {
			const found = findShapeByIdDeep(shape.shapes, id)
			if (found) return found
		}
	}
	return undefined
}
