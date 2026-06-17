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
	boolValue,
	firstChild,
	getElements,
	getOrAddChild,
	intValue,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { fitSrcRectPercents, getImageSizeFromBytes } from '../../gen-utils.js'
import { relativePartName } from '../opc/partnames.js'
import { FILL_CHOICES, normalizeHex, setSolidFill, solidFillColor } from '../oxml/fill.js'
import { resolveSolidFillColor, type ResolvedColor } from './theme-context.js'
import { Chart } from './chart.js'
import { Table } from './table.js'
import { TextFrame } from './text.js'
import type { Slide } from './slide.js'

const A_TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'
const A_CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
// Microsoft's SVG blip extension namespace (a:blip/a:extLst/a:ext/asvg:svgBlip).
const ASVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'

// Schema successors within p:pic (CT_Picture: nvPicPr, blipFill, spPr, style?)
// and within a:blipFill (blip?, srcRect?, (tile|stretch)?), used to keep a
// get-or-added p:blipFill / a:blip in document order.
const PIC_AFTER_BLIPFILL = ['p:spPr', 'p:style']
const BLIPFILL_AFTER_BLIP = ['a:srcRect', 'a:tile', 'a:stretch']

/** Known content-type → file-extension map for image media parts. */
const IMAGE_EXTENSION_BY_CONTENT_TYPE: Readonly<Record<string, string>> = Object.freeze({
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/tiff': 'tiff',
	'image/webp': 'webp',
	'image/svg+xml': 'svg',
	'image/x-emf': 'emf',
	'image/x-wmf': 'wmf',
})

/**
 * Default a media-part file extension from a content type. Known image types use
 * an explicit map; otherwise fall back to the content-type subtype (before any
 * `+suffix`, with a leading `x-` stripped), e.g. `image/x-foo` → `foo`.
 */
function extFromContentType(contentType: string): string {
	const known = IMAGE_EXTENSION_BY_CONTENT_TYPE[contentType.toLowerCase()]
	if (known) return known
	const subtype = contentType.toLowerCase().split('/')[1] ?? ''
	const ext = subtype.split('+')[0].replace(/^x-/, '')
	if (!ext) throw new Error(`Cannot derive a file extension from content type "${contentType}"; pass { extension }`)
	return ext
}

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
const LN_FILL_AFTER = [
	'a:prstDash',
	'a:custDash',
	'a:round',
	'a:bevel',
	'a:miter',
	'a:headEnd',
	'a:tailEnd',
	'a:extLst',
]

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

/** One stop of a gradient fill (`a:gsLst/a:gs`), as read from a shape. */
export interface GradientStop {
	/** Stop offset along the gradient, 0–1 (from `@pos`, thousandths of a percent), or `null` if unset. */
	position: number | null
	/** Explicit RGB colour as 6-hex (`a:srgbClr/@val`), or `null` when the stop uses a scheme colour. */
	color: string | null
	/** Theme colour token (`a:schemeClr/@val`, e.g. `accent1`), or `null` when the stop uses an explicit colour. */
	schemeColor: string | null
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

	/**
	 * Whether the shape is explicitly hidden (`p:cNvPr/@hidden="1"`); `false` when
	 * the attribute is unset. A hidden shape stays in the slide XML but is not
	 * rendered — decks use it as a fallback layer (e.g. a duotone-recolour source
	 * sitting behind the visible icon), so a faithful reader must distinguish it
	 * from the drawn shapes.
	 */
	get hidden(): boolean {
		const cNvPr = nonVisualCNvPr(this.element)
		return boolValue(cNvPr && attr(cNvPr, 'hidden')) === true
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

	/** Line/border width in points (`spPr/a:ln/@w` is EMU; 12700 EMU = 1pt), or `null` when unset. */
	get lineWidthPt(): number | null {
		const ln = this.#line()
		const w = ln ? intValue(attr(ln, 'w')) : null
		return w === null ? null : w / 12700
	}

	/**
	 * Preset-geometry adjustment values (`spPr/a:prstGeom/a:avLst/a:gd`) as a
	 * name → formula map, e.g. `{ adj: 'val 16667' }`. Empty when the shape has no
	 * adjust handles (or uses custom geometry). Pair with {@link presetGeometry}.
	 */
	get adjustValues(): Record<string, string> {
		const props = this.properties()
		const prstGeom = props && firstChild(props, 'a:prstGeom')
		const avLst = prstGeom && firstChild(prstGeom, 'a:avLst')
		const out: Record<string, string> = {}
		if (avLst) {
			for (const gd of getElements(avLst, 'a:gd')) {
				const name = attr(gd, 'name')
				if (name) out[name] = attr(gd, 'fmla') ?? ''
			}
		}
		return out
	}

	/**
	 * Gradient fill stops (`spPr/a:gradFill/a:gsLst/a:gs`) in document order, or
	 * `null` when the shape's fill is not a gradient. Each stop carries its
	 * position (0–1, from `@pos` in thousandths of a percent) and either an
	 * explicit `color` (hex) or a `schemeColor` token, mirroring the
	 * {@link fillColor} / {@link fillSchemeColor} split for solid fills.
	 */
	get gradientStops(): GradientStop[] | null {
		const props = this.properties()
		const grad = props && firstChild(props, 'a:gradFill')
		if (!grad) return null
		const gsLst = firstChild(grad, 'a:gsLst')
		if (!gsLst) return []
		return getElements(gsLst, 'a:gs').map((gs) => {
			const pos = intValue(attr(gs, 'pos'))
			const srgb = firstChild(gs, 'a:srgbClr')
			const scheme = firstChild(gs, 'a:schemeClr')
			return {
				position: pos === null ? null : pos / 100000,
				color: srgb ? attr(srgb, 'val') : null,
				schemeColor: scheme ? attr(scheme, 'val') : null,
			}
		})
	}

	/**
	 * The shape's solid fill resolved against the slide's theme
	 * ({@link Slide.themeContext}) to a literal hex — the resolved counterpart of
	 * {@link fillColor}/{@link fillSchemeColor}, which report the raw reference.
	 * `null` when the shape has no `a:solidFill` (a gradient/none/inherited fill)
	 * or the colour cannot be made literal. The returned {@link ResolvedColor}
	 * reports child colour transforms (`lumMod`/`shade`/…) but does not apply them.
	 */
	get resolvedFill(): ResolvedColor | null {
		return resolveSolidFillColor(this.properties(), this.slide.themeContext())
	}

	/**
	 * The shape's line/border solid fill resolved against the slide's theme to a
	 * literal hex — the resolved counterpart of {@link lineColor}/{@link lineSchemeColor}.
	 * `null` when the shape has no `a:ln/a:solidFill` or it cannot be made literal.
	 */
	get resolvedLine(): ResolvedColor | null {
		return resolveSolidFillColor(this.#line(), this.slide.themeContext())
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
		return txBody ? new TextFrame(txBody, this.slide.part, this.slide.themeContext()) : null
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

	/**
	 * Repoint the blip at a relationship id already present in the slide's
	 * relationships, without minting a new media part. The caller owns ensuring
	 * the id exists and targets an image; use {@link setImage} to add fresh bytes.
	 */
	set imageRelId(value: string) {
		setAttr(this.#getOrAddBlip(), 'r:embed', value)
		this.markDirty()
	}

	/** Absolute partname of the embedded image, resolved via the slide's relationships, or `null`. */
	get imagePartName(): string | null {
		const relId = this.imageRelId
		return relId ? this.slide.relationships.resolveTarget(relId) : null
	}

	/**
	 * Relationship id of the embedded **vector (SVG)** image, read from the
	 * Microsoft SVG blip extension (`a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed`),
	 * or `null` when the picture has no SVG. PowerPoint usually pairs this with a
	 * raster fallback in `a:blip/@r:embed` ({@link imageRelId}), but some exporters
	 * emit an SVG-only blip where `imageRelId` is absent and only this resolves —
	 * so a reader that wants the real drawn art must consult both.
	 */
	get svgRelId(): string | null {
		const svg = this.#svgBlip()
		return svg ? attr(svg, 'r:embed') : null
	}

	/** Absolute partname of the embedded SVG image, resolved via the slide's relationships, or `null`. */
	get svgPartName(): string | null {
		const relId = this.svgRelId
		return relId ? this.slide.relationships.resolveTarget(relId) : null
	}

	/** The `<asvg:svgBlip>` element inside the blip's extLst, or `null` when the picture carries no SVG. */
	#svgBlip(): Element | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const blip = blipFill && firstChild(blipFill, 'a:blip')
		const extLst = blip && firstChild(blip, 'a:extLst')
		if (!extLst) return null
		for (const ext of getElements(extLst, 'a:ext')) {
			for (let node = ext.firstChild; node; node = node.nextSibling) {
				if (node.nodeType !== ELEMENT_NODE) continue
				const el = node as Element
				if (el.localName === 'svgBlip' && el.namespaceURI === ASVG_NS) return el
			}
		}
		return null
	}

	/**
	 * Replace this picture's image with new bytes. Mints a fresh media part under
	 * `/ppt/media/`, registers its content type, wires an `image` relationship
	 * from the owning slide, and repoints the blip's `@r:embed` at it.
	 *
	 * Copy-on-write: the previous media part is never mutated or removed, so any
	 * other picture sharing it (common after `importSlide`/dedup) is unaffected;
	 * an orphaned old part is left in place for a later GC pass to prune.
	 *
	 * `contentType` is required (e.g. `image/png`); the bytes are not sniffed.
	 * `extension` defaults from the content type.
	 *
	 * `fit` controls the picture's `a:srcRect` crop against its current frame
	 * extent (`a:xfrm/a:ext`):
	 * - omitted (default): geometry and crop are left untouched — the caller owns
	 *   sizing. Note an inherited `a:srcRect` was tuned to the *previous* image's
	 *   aspect ratio, so swapping in an image of a different ratio reuses a crop
	 *   that no longer fits and the result looks stretched; pass `fit` to refit.
	 * - `'cover'`: fill the frame, cropping the overflowing axis (no distortion).
	 * - `'contain'`: fit the whole image inside the frame, letterboxing the short
	 *   axis (no distortion).
	 * - `'stretch'`: drop any crop so the full image is stretched to the frame.
	 *
	 * `'cover'`/`'contain'` measure the new bytes' natural size; if unmeasurable
	 * (e.g. an unknown format) the crop is left as-is and a warning is emitted.
	 */
	setImage(
		bytes: Uint8Array,
		options: { contentType: string; extension?: string; fit?: 'cover' | 'contain' | 'stretch' }
	): void {
		const { contentType } = options
		if (!contentType) throw new Error('setImage requires a contentType (e.g. "image/png")')
		const extension = (options.extension ?? extFromContentType(contentType)).toLowerCase().replace(/^\./, '')

		const opc = this.slide.presentation.opc
		const mediaPartName = opc.reserveMediaPartName(extension)
		opc.addPart(mediaPartName, contentType, bytes)
		const relId = this.slide.relationships.add(IMAGE_REL_TYPE, relativePartName(this.slide.partName, mediaPartName)).id

		setAttr(this.#getOrAddBlip(), 'r:embed', relId)
		if (options.fit) this.#applyFit(options.fit, bytes)
		this.markDirty()
	}

	/**
	 * Refit the blip crop after a {@link setImage} swap. `stretch` removes any
	 * `a:srcRect`; `cover`/`contain` recompute it from the new image's natural
	 * size against the frame extent so the swap is aspect-correct.
	 */
	#applyFit(fit: 'cover' | 'contain' | 'stretch', bytes: Uint8Array): void {
		const blipFill = getOrAddChild(this.element, 'p:blipFill', PIC_AFTER_BLIPFILL)
		if (fit === 'stretch') {
			removeChildrenByQName(blipFill, ['a:srcRect'])
			return
		}
		const natural = getImageSizeFromBytes(bytes)
		if (!natural) {
			console.warn(
				`setImage fit '${fit}': could not measure the new image's natural size; leaving the crop unchanged (it may look stretched). Provide a raster (PNG/JPEG/GIF/BMP/WebP) or an SVG with width/height or a viewBox.`
			)
			return
		}
		const cx = this.width
		const cy = this.height
		if (cx == null || cy == null) {
			throw new Error(`setImage fit '${fit}' needs a frame extent (a:xfrm/a:ext); this picture has no transform`)
		}
		const { l, r, t, b } = fitSrcRectPercents(fit, { w: natural.w, h: natural.h }, { w: cx, h: cy })
		const srcRect = getOrAddChild(blipFill, 'a:srcRect', ['a:tile', 'a:stretch'])
		setAttr(srcRect, 'l', String(l))
		setAttr(srcRect, 'r', String(r))
		setAttr(srcRect, 't', String(t))
		setAttr(srcRect, 'b', String(b))
	}

	/** Get-or-add `p:blipFill/a:blip`, keeping both in document order. */
	#getOrAddBlip(): Element {
		const blipFill = getOrAddChild(this.element, 'p:blipFill', PIC_AFTER_BLIPFILL)
		return getOrAddChild(blipFill, 'a:blip', BLIPFILL_AFTER_BLIP)
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
		return tbl ? new Table(tbl, this.slide.part, this.slide.themeContext()) : null
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
 * Wrap a single shape-tree element (`p:sp`/`p:pic`/`p:cxnSp`/`p:graphicFrame`/
 * `p:grpSp`) in its concrete `Shape` proxy, or `null` if it is not a shape kind
 * (e.g. `p:nvGrpSpPr`, `p:grpSpPr`, `p:extLst`).
 */
export function wrapShapeElement(element: Element, slide: Slide): Shape | null {
	if (element.namespaceURI !== OOXML_NS.p) return null
	switch (element.localName) {
		case 'sp':
			return new AutoShape(element, slide)
		case 'pic':
			return new Picture(element, slide)
		case 'cxnSp':
			return new Connector(element, slide)
		case 'graphicFrame':
			return new GraphicFrame(element, slide)
		case 'grpSp':
			return new GroupShape(element, slide)
		default:
			return null
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
		const shape = wrapShapeElement(node as Element, slide)
		if (shape) shapes.push(shape)
	}
	return shapes
}
