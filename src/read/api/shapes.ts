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
	firstChildElement,
	getElements,
	getOrAddChild,
	intValue,
	removeAttr,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { fitSrcRectPercents, getImageSizeFromBytes } from '../../media/image-size.js'
import { composeGroupFrame, type GroupTransform } from '../../group-transform.js'
import { warn } from '../../log.js'
import { relativePartName } from '../opc/partnames.js'
import { FILL_CHOICES, normalizeHex, setSolidFill, solidFillColor } from '../oxml/fill.js'
import {
	resolveColorElement,
	resolveInheritedFrame,
	resolveSolidFillColor,
	resolveStyleFillColor,
	resolveStyleFontRef,
	resolveStyleLineColor,
	type PlaceholderRef,
	type ResolvedColor,
	type ResolvedFrame,
} from './theme-context.js'
import { readGradientFill, readGradientStops, type GradientFill, type GradientStop } from './gradient.js'
import { Chart } from './chart.js'
import { ChartEx } from './chartex.js'
import { Table } from './table.js'
import { TextFrame } from './text.js'
import type { Slide } from './slide.js'

// Re-exported so `pptxgenjs/read` keeps surfacing the gradient types from here even
// though their definitions moved to ./gradient.js (shared with the slide-background reader).
export type { GradientStop, GradientFill } from './gradient.js'

const A_TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'
const A_CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
// chartEx (Office-2016 chart family) graphicData URI + `cx:chart` reference child namespace.
const A_CHARTEX_URI = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
// Microsoft's SVG blip extension namespace (a:blip/a:extLst/a:ext/asvg:svgBlip).
const ASVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
// Microsoft's "decorative" accessibility extension: p:cNvPr/a:extLst/a:ext
// (uri {C183D7F6-B498-43B3-948B-1728B52AA6E4}) / adec:decorative. Confirmed
// against real PowerPoint output ("Mark as decorative").
const ADEC_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative'

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
	const ext = (subtype.split('+')[0] ?? '').replace(/^x-/, '')
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
		if (child.namespaceURI === OOXML_NS.p && child.localName?.startsWith('nv')) {
			return firstChild(child, 'p:cNvPr')
		}
	}
	return null
}

function emuFrom(parent: Element | null, qname: string, attribute: string): number | null {
	const element = parent && firstChild(parent, qname)
	return element ? intValue(attr(element, attribute)) : null
}

/**
 * One `a:pt` coordinate as a raw path-unit integer. A guide-name reference
 * (the `ST_AdjCoordinate` string form) is not produced by authored freeforms;
 * a non-numeric value degrades to `0` rather than crashing (documented edge).
 */
function ptAxis(pt: Element | undefined, axis: 'x' | 'y'): number {
	return (pt ? intValue(attr(pt, axis)) : null) ?? 0
}

/** Parse one `<a:path>` into its viewport attrs (with schema defaults) and ordered segments. */
function readGeometryPath(path: Element): CustomGeometryPath {
	const commands: GeometryCommand[] = []
	for (let node = path.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const seg = node as Element
		if (seg.namespaceURI !== OOXML_NS.a) continue
		const pts = getElements(seg, 'a:pt')
		switch (seg.localName) {
			case 'moveTo':
				commands.push({ cmd: 'moveTo', x: ptAxis(pts[0], 'x'), y: ptAxis(pts[0], 'y') })
				break
			case 'lnTo':
				commands.push({ cmd: 'lnTo', x: ptAxis(pts[0], 'x'), y: ptAxis(pts[0], 'y') })
				break
			case 'cubicBezTo':
				commands.push({
					cmd: 'cubicBezTo',
					x1: ptAxis(pts[0], 'x'),
					y1: ptAxis(pts[0], 'y'),
					x2: ptAxis(pts[1], 'x'),
					y2: ptAxis(pts[1], 'y'),
					x: ptAxis(pts[2], 'x'),
					y: ptAxis(pts[2], 'y'),
				})
				break
			case 'quadBezTo':
				commands.push({
					cmd: 'quadBezTo',
					x1: ptAxis(pts[0], 'x'),
					y1: ptAxis(pts[0], 'y'),
					x: ptAxis(pts[1], 'x'),
					y: ptAxis(pts[1], 'y'),
				})
				break
			case 'arcTo':
				commands.push({
					cmd: 'arcTo',
					wR: intValue(attr(seg, 'wR')) ?? 0,
					hR: intValue(attr(seg, 'hR')) ?? 0,
					stAng: (intValue(attr(seg, 'stAng')) ?? 0) / 60000,
					swAng: (intValue(attr(seg, 'swAng')) ?? 0) / 60000,
				})
				break
			case 'close':
				commands.push({ cmd: 'close' })
				break
		}
	}
	return {
		w: intValue(attr(path, 'w')) ?? 0,
		h: intValue(attr(path, 'h')) ?? 0,
		fill: attr(path, 'fill') ?? 'norm',
		stroke: boolValue(attr(path, 'stroke')) ?? true,
		commands,
	}
}

/** Validate and round an EMU geometry value; extents (`cx`/`cy`) must be non-negative. */
function toEmu(value: number, attribute: string, allowNegative: boolean): number {
	if (!Number.isFinite(value)) throw new Error(`${attribute} must be a finite number of EMU, got ${value}`)
	if (!allowNegative && value < 0) throw new Error(`${attribute} must be non-negative, got ${value}`)
	return Math.round(value)
}

/** Direct child *elements* of `parent`, in document order. */
function childElements(parent: Element): Element[] {
	const out: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) out.push(node as Element)
	}
	return out
}

/** A point + extent pair (`a:off`/`a:ext` or `a:chOff`/`a:chExt`) from a transform, or `null` if either is incomplete. */
function readBox(
	xfrm: Element,
	offName: string,
	extName: string
): { x: number; y: number; cx: number; cy: number } | null {
	const off = firstChild(xfrm, offName)
	const ext = firstChild(xfrm, extName)
	const x = off && intValue(attr(off, 'x'))
	const y = off && intValue(attr(off, 'y'))
	const cx = ext && intValue(attr(ext, 'cx'))
	const cy = ext && intValue(attr(ext, 'cy'))
	if (x === null || y === null || cx === null || cy === null) return null
	return { x, y, cx, cy }
}

/** Convert a DrawingML colour element to a {@link RecolorColor}, or `null` when it is not an `a:` colour element. */
function recolorColorOf(color: Element | null): RecolorColor | null {
	if (!color || color.namespaceURI !== OOXML_NS.a) return null
	return {
		color: color.localName === 'srgbClr' ? attr(color, 'val') : null,
		schemeColor: color.localName === 'schemeClr' ? attr(color, 'val') : null,
		presetColor: color.localName === 'prstClr' ? attr(color, 'val') : null,
	}
}

function rotationDegrees(xfrm: Element): number {
	const rot = intValue(attr(xfrm, 'rot'))
	return rot === null ? 0 : rot / 60000
}

function transformFlipH(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipH')) === true
}

function transformFlipV(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipV')) === true
}

/** One end of a connector/line (`a:ln/a:headEnd` or `a:tailEnd`), as read from a shape. */
export interface LineEnd {
	/** Arrowhead type (`@type`: `none`/`triangle`/`stealth`/`diamond`/`oval`/`arrow`). */
	type: string
	/** Width class (`@w`: `sm`/`med`/`lg`), or `null` when unset. */
	width: string | null
	/** Length class (`@len`: `sm`/`med`/`lg`), or `null` when unset. */
	length: string | null
}

/** Both ends of a shape's line/border, when either carries an arrowhead. */
export interface LineEnds {
	head: LineEnd | null
	tail: LineEnd | null
}

/**
 * One bound end of a connector (`p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn` or `a:endCxn`),
 * as read. A connector authored with `startShape`/`endShape` attaches each end to
 * a shape by that shape's `p:cNvPr/@id` plus a connection-site index; an unbound
 * end (the writer emits a bare `p:cNvCxnSpPr`) reports no site at all — see
 * {@link Connector.startConnection}. This is the read counterpart of the write
 * API's `startShape`/`startShapeIdx` split, which binds by `objectName` and
 * resolves the name → id at serialize time.
 */
export interface ConnectionSite {
	/** The bound shape's drawing id (`@id`, i.e. its `p:cNvPr/@id`). */
	shapeId: number
	/** Connection-site index on the bound shape (`@idx`; 0-based, preset-dependent). */
	siteIndex: number
	/**
	 * The bound shape resolved to a read-model shape via {@link Slide.shapeByIdDeep},
	 * which descends into groups — so a connector bound to a shape nested in a group
	 * resolves the same as one bound to a top-level shape. `null` only when no shape
	 * anywhere on the slide carries that id (a genuinely dangling binding), which is
	 * faithful degradation and does not throw.
	 */
	boundShape: AnyShape | null
}

/**
 * A shape's outer drop shadow (`spPr/a:effectLst/a:outerShdw`), as read from a
 * shape and resolved against the slide theme. Distances are in points (the EMU
 * source ÷ 12700) and the direction in degrees (the `60000`ths source ÷ 60000),
 * matching the write-side {@link ShadowProps} convention so it round-trips.
 */
export interface OuterShadow {
	/** Effective shadow colour as 6-hex (theme-resolved, transforms applied), or `null`. */
	color: string | null
	/** Theme colour token when the shadow colour was a scheme colour (e.g. `accent1`), else `undefined`. */
	colorToken?: string
	/** Shadow opacity 0–1 (from the colour's `a:alpha`), or `undefined` when fully opaque. */
	alpha?: number
	/** Blur radius in points (`@blurRad` ÷ 12700), or `undefined` when unset. */
	blurPt?: number
	/** Offset distance in points (`@dist` ÷ 12700), or `undefined` when unset. */
	offsetPt?: number
	/** Offset direction in degrees, clockwise from 3 o'clock (`@dir` ÷ 60000), or `undefined` when unset. */
	angleDeg?: number
}

/**
 * A shape's inner shadow (`spPr/a:effectLst/a:innerShdw`), resolved against the
 * slide theme. Identical fields to {@link OuterShadow} — CT_InnerShadowEffect
 * carries the same `blurRad`/`dist`/`dir` + colour — but a distinct type so a
 * consumer can tell an inset shadow from a drop shadow. Distances in points,
 * direction in degrees, matching the write-side `shadow: { type: 'inner' }`.
 */
export interface InnerShadow {
	/** Effective shadow colour as 6-hex (theme-resolved, transforms applied), or `null`. */
	color: string | null
	/** Theme colour token when the shadow colour was a scheme colour (e.g. `accent1`), else `undefined`. */
	colorToken?: string
	/** Shadow opacity 0–1 (from the colour's `a:alpha`), or `undefined` when fully opaque. */
	alpha?: number
	/** Blur radius in points (`@blurRad` ÷ 12700), or `undefined` when unset. */
	blurPt?: number
	/** Offset distance in points (`@dist` ÷ 12700), or `undefined` when unset. */
	offsetPt?: number
	/** Offset direction in degrees, clockwise from 3 o'clock (`@dir` ÷ 60000), or `undefined` when unset. */
	angleDeg?: number
}

/**
 * A shape's glow effect (`spPr/a:effectLst/a:glow`) — a coloured halo — resolved
 * against the slide theme. The write-side text glow (`glow: { size, color,
 * opacity }`) emits the same element, so {@link radiusPt} (`@rad` ÷ 12700) and the
 * colour round-trip.
 */
export interface Glow {
	/** Effective glow colour as 6-hex (theme-resolved, transforms applied), or `null`. */
	color: string | null
	/** Theme colour token when the glow colour was a scheme colour (e.g. `accent1`), else `undefined`. */
	colorToken?: string
	/** Glow opacity 0–1 (from the colour's `a:alpha`), or `undefined` when fully opaque. */
	alpha?: number
	/** Glow radius in points (`@rad` ÷ 12700), or `undefined` when unset. */
	radiusPt?: number
}

/**
 * A shape's reflection effect (`spPr/a:effectLst/a:reflection`) — a mirrored fade
 * beneath the shape. This library's writer authors none, so this is a **read-only**
 * surface: a consumer that finds one should carry the part verbatim rather than
 * regenerate it. Only the attributes a faithful replica needs are decoded —
 * distances in points (÷ 12700), directions in degrees (÷ 60000), and the start/end
 * alpha and position pairs as 0–1 fractions (the `1000`ths-of-a-percent source
 * ÷ 100000). All fields are optional; an attribute absent from the source is omitted.
 */
export interface Reflection {
	/** Blur radius in points (`@blurRad` ÷ 12700). */
	blurPt?: number
	/** Start opacity 0–1 (`@stA` ÷ 100000). */
	startAlpha?: number
	/** Start position along the reflection 0–1 (`@stPos` ÷ 100000). */
	startPos?: number
	/** End opacity 0–1 (`@endA` ÷ 100000). */
	endAlpha?: number
	/** End position along the reflection 0–1 (`@endPos` ÷ 100000). */
	endPos?: number
	/** Offset distance in points (`@dist` ÷ 12700). */
	distPt?: number
	/** Offset direction in degrees (`@dir` ÷ 60000). */
	angleDeg?: number
	/** Fade direction in degrees (`@fadeDir` ÷ 60000). */
	fadeAngleDeg?: number
}

/**
 * A shape's soft-edge effect (`spPr/a:effectLst/a:softEdge`) — a feathered border.
 * Like {@link Reflection}, the writer authors none, so carry it verbatim. `@rad`
 * (the feather radius) ÷ 12700 → points.
 */
export interface SoftEdge {
	/** Feather radius in points (`@rad` ÷ 12700). */
	radiusPt: number
}

/**
 * A shape's pattern fill (`spPr/a:pattFill`) — a two-colour preset hatch. The
 * write-side `fill: { type: 'pattern', pattern: { preset, fgColor, bgColor } }`
 * emits the same element, so the {@link preset} name and both colours round-trip.
 * Colours resolve against the slide theme (a scheme token → literal hex) the same
 * way {@link Shape.resolvedFill} resolves a solid fill.
 */
export interface PatternFill {
	/** Preset pattern name (`@prst`, e.g. `pct50`/`diagCross`/`ltUpDiag`), or `null` when unset. */
	preset: string | null
	/** Foreground colour (`a:fgClr`) resolved against the theme, or `null`. */
	foreground: ResolvedColor | null
	/** Background colour (`a:bgClr`) resolved against the theme, or `null`. */
	background: ResolvedColor | null
}

/**
 * One segment of a custom-geometry path (`a:path`), as read from a shape. The
 * command verbs mirror the write-side `GeometryPoint` DSL (`src/core-interfaces.ts`)
 * so a consumer can map a `GeometryCommand[]` to `GeometryPoint[]` one-to-one.
 *
 * Coordinates (`x`/`y`/`x1`…) are raw path-unit integers in the path's own
 * `0..w` / `0..h` space (see {@link CustomGeometryPath.w}); they are not EMU and
 * must be scaled by the consumer against the shape's box. `arcTo` angles are in
 * **degrees** (the raw `60000`ths-of-a-degree values divided by 60000), matching
 * the write DSL's degree input.
 */
export type GeometryCommand =
	| { cmd: 'moveTo'; x: number; y: number }
	| { cmd: 'lnTo'; x: number; y: number }
	| { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { cmd: 'quadBezTo'; x1: number; y1: number; x: number; y: number }
	| { cmd: 'arcTo'; wR: number; hR: number; stAng: number; swAng: number }
	| { cmd: 'close' }

/** One `<a:path>` of a custom geometry, with its path-unit viewport and render attrs. */
export interface CustomGeometryPath {
	/** Path-unit width (`a:path/@w`); the `x` denominator for this path's coords. Default `0`. */
	w: number
	/** Path-unit height (`a:path/@h`); the `y` denominator for this path's coords. Default `0`. */
	h: number
	/** Fill mode (`a:path/@fill`, `ST_PathFillMode`, e.g. `norm`/`none`/`lighten`). Default `norm`. */
	fill: string
	/** Whether the path is stroked (`a:path/@stroke`). Default `true`. */
	stroke: boolean
	/** The path's segments in document order — order *is* the geometry. */
	commands: GeometryCommand[]
}

/**
 * Custom freeform geometry (`spPr/a:custGeom/a:pathLst`), as read from a shape.
 * Faithfully exposes every `a:path` rather than flattening to a single command
 * list: the schema allows repeatable `a:path`, each with independent
 * `fill`/`stroke`. Desktop PowerPoint's own freeforms only ever emit one
 * `a:path` (a hole is one path with two `moveTo`…`close` contours), but
 * multi-`a:path` input is schema-legal (e.g. SVG import) and preserved here.
 */
export interface CustomGeometry {
	paths: CustomGeometryPath[]
}

/** A shape's resolved position and size in slide-absolute EMU. */
export interface AbsoluteFrame {
	left: number
	top: number
	width: number
	height: number
	/** Effective clockwise rotation in degrees after composing enclosing group rotations. */
	rotation: number
	/** Effective horizontal flip after XOR-composing enclosing group flips. */
	flipH: boolean
	/** Effective vertical flip after XOR-composing enclosing group flips. */
	flipV: boolean
}

/**
 * A colour reference inside a picture recolour effect, split by colour model
 * (mirrors {@link GradientStop}). At most one field is non-`null`.
 */
export interface RecolorColor {
	/** Explicit RGB as 6-hex (`a:srgbClr/@val`), or `null`. */
	color: string | null
	/** Theme colour token (`a:schemeClr/@val`, e.g. `accent1`), or `null`. */
	schemeColor: string | null
	/** Preset colour name (`a:prstClr/@val`, e.g. `black`/`white` — the duotone icon-tint stops), or `null`. */
	presetColor: string | null
}

/**
 * A picture's blip recolour effect (`p:blipFill/a:blip` recolour child), as read.
 * A small discriminated union over the effects a faithful reader needs to
 * reproduce a recoloured image. Colours use the same `color`/`schemeColor`/
 * `presetColor` split as {@link GradientStop}, so theme tokens can resolve through
 * {@link Slide.themeContext}. `threshold`/`amount` are 0–1 fractions.
 */
export type Recolor =
	| { kind: 'duotone'; stops: RecolorColor[] }
	| { kind: 'clrChange'; from: RecolorColor | null; to: RecolorColor | null }
	| { kind: 'grayscale' }
	| { kind: 'biLevel'; threshold: number | null }
	| { kind: 'alphaModFix'; amount: number }

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

	/**
	 * The shape's alt-text description (`p:cNvPr/@descr`), or `null` when unset.
	 * This is the accessibility primitive a screen reader announces; an audit or
	 * accessible-export tool reads it here. Distinct from {@link name}, which is
	 * the authoring-time shape name and is never surfaced to assistive tech.
	 */
	get description(): string | null {
		const cNvPr = nonVisualCNvPr(this.element)
		const v = cNvPr && attr(cNvPr, 'descr')
		return v ?? null
	}

	/**
	 * Set (or clear, with `''`) the alt-text description. Requires the shape's
	 * `p:cNvPr` to exist, which every well-formed shape carries.
	 */
	set description(value: string) {
		const cNvPr = nonVisualCNvPr(this.element)
		if (!cNvPr) throw new Error('cannot set description: shape has no p:cNvPr')
		if (value === '') removeAttr(cNvPr, 'descr')
		else setAttr(cNvPr, 'descr', value)
		this.markDirty()
	}

	/**
	 * The shape's alt-text title (`p:cNvPr/@title`), or `null` when unset. Modern
	 * PowerPoint no longer exposes a separate title field (only description +
	 * "mark as decorative"), so this is usually `null`; it survives on decks
	 * authored by older PowerPoint or other producers that still write it.
	 */
	get title(): string | null {
		const cNvPr = nonVisualCNvPr(this.element)
		const v = cNvPr && attr(cNvPr, 'title')
		return v ?? null
	}

	/**
	 * Whether the shape is flagged **decorative** (`p:cNvPr/a:extLst/a:ext`
	 * uri `{C183D7F6-…}` / `adec:decorative@val`), PowerPoint's "Mark as
	 * decorative" — a purely visual element assistive tech should skip. `false`
	 * when the extension is absent. A decorative shape typically has no
	 * {@link description}; the two are alternatives, not companions.
	 */
	get isDecorative(): boolean {
		const cNvPr = nonVisualCNvPr(this.element)
		const extLst = cNvPr && firstChild(cNvPr, 'a:extLst')
		if (!extLst) return false
		for (const ext of getElements(extLst, 'a:ext')) {
			for (const child of childElements(ext)) {
				if (child.localName === 'decorative' && child.namespaceURI === ADEC_NS) {
					return boolValue(attr(child, 'val')) === true
				}
			}
		}
		return false
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

	/**
	 * This shape's effective position and size in EMU: its own `a:xfrm` when it has
	 * one ({@link left}/{@link top}/{@link width}/{@link height}, tagged
	 * `source: 'own'`); otherwise, for a placeholder, the geometry it inherits from
	 * the matching layout placeholder, else the master's (tagged accordingly). A
	 * non-placeholder shape with no own transform has nothing to inherit from and
	 * reads `null`, as does a placeholder whose layout/master chain defines no
	 * matching geometry either.
	 *
	 * The writer always emits an explicit `a:xfrm` on every placeholder it authors,
	 * so `source` reads `'own'` for every authored deck; `'layout'`/`'master'` is
	 * the case an *imported* deck exercises, when PowerPoint itself leaves a
	 * placeholder's geometry to inherit.
	 */
	get resolvedFrame(): ResolvedFrame | null {
		const left = this.left
		const top = this.top
		const width = this.width
		const height = this.height
		if (left !== null && top !== null && width !== null && height !== null) {
			return { left, top, width, height, source: 'own' }
		}
		const ph = this.placeholder
		if (!ph) return null
		return resolveInheritedFrame(ph, this.slide.themeContext())
	}

	/**
	 * Clockwise rotation in degrees (`a:xfrm/@rot` ÷ 60000), or `null` when the
	 * shape has no own transform. A present xfrm with no `@rot` reads as `0`, so
	 * — mirroring {@link left}/{@link top}/{@link width}/{@link height} — `null`
	 * ("inherits layout geometry") stays distinct from `0` ("has a transform, not
	 * rotated"). The value is faithful to the XML and not normalised to a signed
	 * range, so a `@rot` past 360° (e.g. a negative angle stored as `19216344`)
	 * reads back greater than 360. This is the shape's own orientation; use
	 * {@link absoluteFrame} when you need the effective orientation after
	 * enclosing group transforms are composed.
	 */
	get rotation(): number | null {
		const xfrm = this.xfrm()
		if (!xfrm) return null
		return rotationDegrees(xfrm)
	}

	/**
	 * Whether the shape is flipped horizontally (`a:xfrm/@flipH`); `false` when
	 * unset or when the shape has no own transform. This is the shape's own
	 * horizontal flip; use {@link absoluteFrame} for the effective value after
	 * enclosing group transforms are composed.
	 */
	get flipH(): boolean {
		const xfrm = this.xfrm()
		return xfrm !== null && transformFlipH(xfrm)
	}

	/**
	 * Whether the shape is flipped vertically (`a:xfrm/@flipV`); `false` when
	 * unset or when the shape has no own transform. This is the shape's own
	 * vertical flip; use {@link absoluteFrame} for the effective value after
	 * enclosing group transforms are composed.
	 */
	get flipV(): boolean {
		const xfrm = this.xfrm()
		return xfrm !== null && transformFlipV(xfrm)
	}

	/**
	 * This shape's position, size, and effective orientation in **slide-absolute**
	 * EMU/degrees, composing every enclosing group transform.
	 *
	 * {@link left}/{@link top}/{@link width}/{@link height} report a group child's
	 * geometry in its group's child coordinate space (`a:chOff`/`a:chExt`), which is
	 * not directly placeable on the slide. This getter walks the `p:grpSp` ancestor
	 * chain outward, mapping the box through each group's
	 * `off + (p - chOff) * (ext / chExt)` transform, then composing group flips and
	 * rotations about the group centre. For a shape already at slide level the box
	 * equals its own `{ left, top, width, height }`, while `rotation`/`flipH`/`flipV`
	 * equal the shape's own transform values.
	 *
	 * `null` when the shape (or any enclosing group) has no own transform, or a
	 * group's `a:chExt` is degenerate (zero) — there is then no resolvable frame.
	 *
	 * A rotated shape's returned `left`/`top` remain PowerPoint's unrotated
	 * placement box (the same box PowerPoint writes after Ungroup), with the
	 * effective rotation exposed separately.
	 */
	get absoluteFrame(): AbsoluteFrame | null {
		const xfrm = this.xfrm()
		if (!xfrm) return null
		const box = readBox(xfrm, 'a:off', 'a:ext')
		if (!box) return null

		const groups: GroupTransform[] = []
		for (let node = this.element.parentNode; node && node.nodeType === ELEMENT_NODE; node = node.parentNode) {
			const parent = node as Element
			if (parent.namespaceURI !== OOXML_NS.p || parent.localName !== 'grpSp') break // reached the shape tree (or a non-group)
			const grpSpPr = firstChild(parent, 'p:grpSpPr')
			const groupXfrm = grpSpPr && firstChild(grpSpPr, 'a:xfrm')
			const outer = groupXfrm && readBox(groupXfrm, 'a:off', 'a:ext')
			const child = groupXfrm && readBox(groupXfrm, 'a:chOff', 'a:chExt')
			if (!groupXfrm || !outer || !child) return null
			groups.push({
				outer,
				child,
				rotation: rotationDegrees(groupXfrm),
				flipH: transformFlipH(groupXfrm),
				flipV: transformFlipV(groupXfrm),
			})
		}

		const frame = composeGroupFrame(
			{ box, rotation: rotationDegrees(xfrm), flipH: transformFlipH(xfrm), flipV: transformFlipV(xfrm) },
			groups
		)
		if (!frame) return null
		return {
			left: Math.round(frame.box.x),
			top: Math.round(frame.box.y),
			width: Math.round(frame.box.cx),
			height: Math.round(frame.box.cy),
			rotation: frame.rotation,
			flipH: frame.flipH,
			flipV: frame.flipV,
		}
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
	 * The line/border dash style (`spPr/a:ln/a:prstDash/@val`), e.g. `'dash'`,
	 * `'lgDashDot'`, `'sysDot'`, or `null` when the line is solid/unset. A faithful
	 * replica of dashed dividers and dashed card borders needs this — it is
	 * otherwise invisible in {@link lineColor}/{@link lineWidthPt} alone.
	 */
	get lineDash(): string | null {
		const ln = this.#line()
		const dash = ln && firstChild(ln, 'a:prstDash')
		return dash ? (attr(dash, 'val') ?? null) : null
	}

	/**
	 * `true` when the shape sets an explicit no-line (`spPr/a:ln/a:noFill`) — a
	 * deliberately border-less shape. Distinct from simply having no `a:ln`
	 * (an inherited line), which {@link resolvedLine} cannot tell apart: both
	 * report `null`. A replica that relies on a shadow instead of a border needs
	 * to know the border was explicitly suppressed.
	 */
	get lineNoFill(): boolean {
		const ln = this.#line()
		return !!(ln && firstChild(ln, 'a:noFill'))
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
		return props ? this.#gradientStopsIn(props) : null
	}

	/**
	 * Read `a:gradFill/a:gsLst` stops from a container (either `spPr` for a fill or
	 * `a:ln` for a line stroke). `null` when the container has no gradient; `[]`
	 * when the gradient carries no stop list.
	 */
	#gradientStopsIn(container: Element): GradientStop[] | null {
		return readGradientStops(container, this.slide.themeContext())
	}

	/**
	 * Read the full `a:gradFill` (stops + linear angle / path shape) from a
	 * container (`spPr` for a fill or `a:ln` for a line stroke). `null` when the
	 * container has no gradient.
	 */
	#gradientFillIn(container: Element): GradientFill | null {
		return readGradientFill(container, this.slide.themeContext())
	}

	/**
	 * The shape's gradient fill with its geometry (`spPr/a:gradFill`), or `null`
	 * when the fill is not a gradient. Unlike {@link gradientStops} (stops only),
	 * this also carries the linear {@link GradientFill.angleDeg} or the
	 * {@link GradientFill.path} shape — the geometry a faithful replica needs and
	 * which the bare stop list omits.
	 */
	get gradientFill(): GradientFill | null {
		const props = this.properties()
		return props ? this.#gradientFillIn(props) : null
	}

	/**
	 * The shape's line/border **gradient** stroke (`spPr/a:ln/a:gradFill`), or
	 * `null` when the line is a solid, absent, or inherited border (see
	 * {@link resolvedLine}). The line counterpart of {@link gradientFill}: a
	 * gradient-stroked connector — common for faded process arrows in styled decks
	 * — otherwise surfaces only its {@link lineWidthPt}, dropping the colour
	 * entirely, so a replica cannot reproduce the stroke.
	 */
	get lineGradient(): GradientFill | null {
		const ln = this.#line()
		return ln ? this.#gradientFillIn(ln) : null
	}

	/**
	 * The shape's line/connector arrowheads (`a:ln/a:headEnd` + `a:tailEnd`), or
	 * `null` when neither end carries one. Essential for replicating connectors,
	 * whose dot/arrow ends are otherwise invisible in the geometry.
	 */
	get lineEnds(): LineEnds | null {
		const ln = this.#line()
		if (!ln) return null
		const read = (qn: string): LineEnd | null => {
			const el = firstChild(ln, qn)
			if (!el) return null
			return { type: attr(el, 'type') ?? 'none', width: attr(el, 'w') ?? null, length: attr(el, 'len') ?? null }
		}
		const head = read('a:headEnd')
		const tail = read('a:tailEnd')
		return head || tail ? { head, tail } : null
	}

	/**
	 * The shape's outer drop shadow (`spPr/a:effectLst/a:outerShdw`), resolved
	 * against the slide theme, or `null` when the shape has no outer shadow. The
	 * soft brand shadows the eye reads as "floating" panels live here and are
	 * invisible in geometry/fill alone.
	 */
	get shadow(): OuterShadow | null {
		const shdw = this.#effect('a:outerShdw')
		return shdw ? this.#readShadow(shdw) : null
	}

	/**
	 * The shape's **inner** shadow (`spPr/a:effectLst/a:innerShdw`), resolved against
	 * the slide theme, or `null` when the shape has no inner shadow. The inset
	 * counterpart of {@link shadow}: the write-side `shadow: { type: 'inner' }`
	 * emits it, and it is invisible in geometry/fill alone.
	 */
	get innerShadow(): InnerShadow | null {
		const shdw = this.#effect('a:innerShdw')
		return shdw ? this.#readShadow(shdw) : null
	}

	/**
	 * The shape's glow halo (`spPr/a:effectLst/a:glow`), resolved against the slide
	 * theme, or `null` when the shape has no glow. Same element the write-side text
	 * glow emits, so its {@link Glow.radiusPt} and colour round-trip.
	 */
	get glow(): Glow | null {
		const glow = this.#effect('a:glow')
		if (!glow) return null
		const out: Glow = { color: null }
		this.#applyEffectColor(out, firstChildElement(glow))
		const rad = intValue(attr(glow, 'rad'))
		if (rad !== null) out.radiusPt = rad / 12700
		return out
	}

	/**
	 * The shape's reflection (`spPr/a:effectLst/a:reflection`), or `null` when it has
	 * none. Read-only: this library authors no reflection, so a replica should carry
	 * the part rather than regenerate it — see {@link Reflection}.
	 */
	get reflection(): Reflection | null {
		const refl = this.#effect('a:reflection')
		if (!refl) return null
		const out: Reflection = {}
		const put = (target: keyof Reflection, name: string, div: number): void => {
			const v = intValue(attr(refl, name))
			if (v !== null) out[target] = v / div
		}
		put('blurPt', 'blurRad', 12700)
		put('distPt', 'dist', 12700)
		put('angleDeg', 'dir', 60000)
		put('fadeAngleDeg', 'fadeDir', 60000)
		put('startAlpha', 'stA', 100000)
		put('startPos', 'stPos', 100000)
		put('endAlpha', 'endA', 100000)
		put('endPos', 'endPos', 100000)
		return out
	}

	/**
	 * The shape's soft (feathered) edge (`spPr/a:effectLst/a:softEdge`), or `null`
	 * when it has none. Read-only like {@link reflection}: carry, don't regenerate.
	 */
	get softEdge(): SoftEdge | null {
		const soft = this.#effect('a:softEdge')
		if (!soft) return null
		const rad = intValue(attr(soft, 'rad'))
		return { radiusPt: rad === null ? 0 : rad / 12700 }
	}

	/**
	 * The shape's pattern (hatch) fill (`spPr/a:pattFill`), or `null` when the fill
	 * is not a pattern. Surfaces the {@link PatternFill.preset} name and both
	 * colours resolved against the slide theme — the pattern counterpart of
	 * {@link resolvedFill}, which reports `null` for a non-solid fill and so drops a
	 * hatched surface entirely.
	 */
	get patternFill(): PatternFill | null {
		const props = this.properties()
		const patt = props && firstChild(props, 'a:pattFill')
		if (!patt) return null
		const ctx = this.slide.themeContext()
		const resolveWrap = (qname: string): ResolvedColor | null => {
			const wrap = firstChild(patt, qname)
			const colorEl = wrap && firstChildElement(wrap)
			return colorEl ? resolveColorElement(colorEl, ctx) : null
		}
		return {
			preset: attr(patt, 'prst') ?? null,
			foreground: resolveWrap('a:fgClr'),
			background: resolveWrap('a:bgClr'),
		}
	}

	/** A named child of the shape's effect list (`spPr/a:effectLst/<qname>`), or `null`. */
	#effect(qname: string): Element | null {
		const props = this.properties()
		const effectLst = props && firstChild(props, 'a:effectLst')
		return effectLst ? firstChild(effectLst, qname) : null
	}

	/** Resolve `colorEl` against the theme and stamp `color`/`colorToken`/`alpha` onto an effect result. */
	#applyEffectColor(out: { color: string | null; colorToken?: string; alpha?: number }, colorEl: Element | null): void {
		const resolved = resolveColorElement(colorEl, this.slide.themeContext())
		if (resolved) {
			out.color = resolved.effectiveHex
			if (resolved.alpha !== undefined) out.alpha = resolved.alpha
		}
		if (colorEl && colorEl.localName === 'schemeClr') out.colorToken = attr(colorEl, 'val') ?? undefined
	}

	/** Decode a shadow element (`a:outerShdw`/`a:innerShdw` share the fields), resolving its colour. */
	#readShadow(shdw: Element): OuterShadow {
		const out: OuterShadow = { color: null }
		this.#applyEffectColor(out, firstChild(shdw, 'a:srgbClr') ?? firstChild(shdw, 'a:schemeClr'))
		const blur = intValue(attr(shdw, 'blurRad'))
		const dist = intValue(attr(shdw, 'dist'))
		const dir = intValue(attr(shdw, 'dir'))
		if (blur !== null) out.blurPt = blur / 12700
		if (dist !== null) out.offsetPt = dist / 12700
		if (dir !== null) out.angleDeg = dir / 60000
		return out
	}

	/**
	 * The shape's solid fill resolved against the slide's theme
	 * ({@link Slide.themeContext}) to a literal hex — the resolved counterpart of
	 * {@link fillColor}/{@link fillSchemeColor}, which report the raw reference.
	 * `null` when the shape has no `a:solidFill` (a gradient/none/inherited fill)
	 * or the colour cannot be made literal. The returned {@link ResolvedColor}
	 * carries the base `hex` and raw transforms, and `effectiveHex` — the base with
	 * its colour transforms (`lumMod`/`shade`/…) applied (read that for the final
	 * rendered colour).
	 *
	 * When the shape carries no explicit `spPr` fill choice, this falls back to the
	 * fill the shape inherits from its `p:style` `a:fillRef` (the theme style
	 * matrix), resolved the same way the `theme: 'preserve'` flatten path bakes it.
	 */
	get resolvedFill(): ResolvedColor | null {
		const ctx = this.slide.themeContext()
		const props = this.properties()
		if (props && FILL_CHOICES.some((q) => firstChild(props, q))) return resolveSolidFillColor(props, ctx)
		return resolveStyleFillColor(this.element, ctx)
	}

	/**
	 * The shape's line/border solid fill resolved against the slide's theme to a
	 * literal hex — the resolved counterpart of {@link lineColor}/{@link lineSchemeColor}.
	 * `null` when the shape has no `a:ln/a:solidFill` or it cannot be made literal.
	 * Like {@link resolvedFill}, the result carries `effectiveHex` (the base colour
	 * with its transforms applied) for the final rendered colour.
	 *
	 * When the shape carries no explicit `spPr/a:ln`, this falls back to the line
	 * the shape inherits from its `p:style` `a:lnRef` (the theme style matrix).
	 */
	get resolvedLine(): ResolvedColor | null {
		const ctx = this.slide.themeContext()
		const ln = this.#line()
		return ln ? resolveSolidFillColor(ln, ctx) : resolveStyleLineColor(this.element, ctx)
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
	 * Convenience: replace the shape's text with a single run, preserving the
	 * first existing run's formatting (see {@link TextFrame.text}). Throws when the
	 * shape has no text frame. For multiple runs or per-run formatting, edit
	 * `textFrame.paragraphs[].runs[]` directly.
	 */
	set text(value: string) {
		const frame = this.textFrame
		if (!frame) throw new Error('Shape has no text frame to set text on')
		frame.text = value
	}

	/**
	 * This shape's placeholder identity (`p:ph` `type`/`idx`), or `null` when it is
	 * not a placeholder. Only `p:sp` shapes can be placeholders, so the base
	 * implementation always returns `null`; {@link AutoShape} overrides it.
	 */
	get placeholder(): PlaceholderRef | null {
		return null
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
		if (!txBody) return null
		const flatten = this.slide.themeContext()
		// Every slide shape resolves its `p:style/a:fontRef` and the presentation's
		// `p:defaultTextStyle`, so a context is always supplied; `ph` is null for a
		// non-placeholder shape, which then skips only the layout/master placeholder tiers.
		const ph = this.placeholder
		const fontRef = resolveStyleFontRef(this.element, flatten)
		return new TextFrame(txBody, this.slide.part, flatten, { ph, flatten, fontRef }, this.slide.relationships)
	}

	/**
	 * This shape's placeholder identity (`p:ph` `type`/`idx`), or `null` when it is
	 * not a placeholder. `idx` defaults to `'0'` when the attribute is absent, as
	 * PowerPoint does. Use {@link Slide.placeholder} to find a placeholder by type.
	 */
	override get placeholder(): PlaceholderRef | null {
		const nvSpPr = firstChild(this.element, 'p:nvSpPr')
		const nvPr = nvSpPr && firstChild(nvSpPr, 'p:nvPr')
		const ph = nvPr && firstChild(nvPr, 'p:ph')
		return ph ? { type: attr(ph, 'type'), idx: attr(ph, 'idx') ?? '0' } : null
	}

	/** Preset geometry name (`a:prstGeom/@prst`, e.g. `rect`), or `null` for custom/none. */
	get presetGeometry(): string | null {
		const spPr = firstChild(this.element, 'p:spPr')
		const prstGeom = spPr && firstChild(spPr, 'a:prstGeom')
		return prstGeom ? attr(prstGeom, 'prst') : null
	}

	/**
	 * Custom freeform geometry (`spPr/a:custGeom/a:pathLst`), or `null` when the
	 * shape uses preset geometry / none. The faithful, multi-path counterpart of
	 * {@link presetGeometry}: each `a:path` keeps its own path-unit viewport
	 * (`w`/`h`) and ordered {@link GeometryCommand}s. Coordinates are raw path-unit
	 * integers, not EMU — pair the path `w`/`h` with the shape's box size to map
	 * them into slide space.
	 */
	get customGeometry(): CustomGeometry | null {
		const props = this.properties()
		const custGeom = props && firstChild(props, 'a:custGeom')
		if (!custGeom) return null
		const pathLst = firstChild(custGeom, 'a:pathLst')
		const paths = pathLst ? getElements(pathLst, 'a:path').map((p) => readGeometryPath(p)) : []
		return { paths }
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

	/**
	 * Which drawable media this picture carries:
	 * - `'raster'` — only a raster blip (`a:blip/@r:embed`);
	 * - `'svg'` — only a vector blip (`asvg:svgBlip/@r:embed`, no raster). This
	 *   is what PowerPoint's *Insert → Icons* and a plain SVG insert produce;
	 * - `'both'` — a raster fallback *and* an SVG (PowerPoint's usual pairing);
	 * - `'none'` — a `p:pic` with no embedded blip at all (e.g. a linked image).
	 *
	 * Lets a caller distinguish an SVG-only picture — where {@link imagePartName}
	 * is legitimately `null` — from a genuinely empty one, without two null checks.
	 */
	get mediaKind(): 'raster' | 'svg' | 'both' | 'none' {
		const hasRaster = this.imageRelId != null
		const hasSvg = this.svgRelId != null
		if (hasRaster && hasSvg) return 'both'
		if (hasRaster) return 'raster'
		if (hasSvg) return 'svg'
		return 'none'
	}

	/**
	 * Absolute partname of whichever part actually carries this picture's drawn
	 * data — the raster part when present, otherwise the SVG part — or `null`
	 * when the picture embeds neither. Use this when you just want "the bytes
	 * this picture shows"; prefer {@link imagePartName} / {@link svgPartName}
	 * (and {@link mediaKind}) when you need to know *which* kind it is. An
	 * SVG-only picture returns its SVG part here even though `imagePartName` is
	 * `null`.
	 */
	get mediaPartName(): string | null {
		return this.imagePartName ?? this.svgPartName
	}

	/**
	 * The picture's crop as fractions of the *source image*, read from
	 * `p:blipFill/a:srcRect` — `{ left, top, right, bottom }`, each the amount
	 * trimmed off that edge (so `0.1` = 10 % cropped away, an uncropped edge is
	 * `0`). `null` when there is no `a:srcRect` at all; an explicit
	 * `{0,0,0,0}` crop still reports zeros, since its presence is meaningful.
	 * The raw attributes are thousandths of a percent; this divides by 100000 to
	 * match the fraction convention used elsewhere in the read API (see
	 * {@link recolor}).
	 */
	get crop(): { left: number; top: number; right: number; bottom: number } | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const srcRect = blipFill && firstChild(blipFill, 'a:srcRect')
		if (!srcRect) return null
		const edge = (name: string): number => {
			const v = intValue(attr(srcRect, name))
			return v === null ? 0 : v / 100000
		}
		return { left: edge('l'), top: edge('t'), right: edge('r'), bottom: edge('b') }
	}

	/**
	 * The picture's blip recolour effect (`p:blipFill/a:blip` recolour child), or
	 * `null` when the blip carries none. Recognises the effects a faithful reader
	 * needs to reproduce a recoloured image: `a:duotone` (the two-stop icon-tint
	 * trick), `a:clrChange`, `a:grayscl`, `a:biLevel`, and `a:alphaModFix`; the
	 * first such effect in document order wins. Colours mirror the
	 * {@link GradientStop} split (`color`/`schemeColor`/`presetColor`) so theme
	 * tokens resolve through {@link Slide.themeContext}. `threshold`/`amount` are
	 * 0–1 fractions. {@link hidden} (the duotone fallback-layer trick) reports the
	 * *visibility* of a recolour source; this reports the *tint* itself.
	 */
	get recolor(): Recolor | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const blip = blipFill && firstChild(blipFill, 'a:blip')
		if (!blip) return null
		for (const child of childElements(blip)) {
			if (child.namespaceURI !== OOXML_NS.a) continue
			switch (child.localName) {
				case 'duotone':
					return {
						kind: 'duotone',
						stops: childElements(child)
							.map(recolorColorOf)
							.filter((c): c is RecolorColor => c !== null),
					}
				case 'clrChange': {
					const from = firstChild(child, 'a:clrFrom')
					const to = firstChild(child, 'a:clrTo')
					return {
						kind: 'clrChange',
						from: from ? recolorColorOf(firstChildElement(from)) : null,
						to: to ? recolorColorOf(firstChildElement(to)) : null,
					}
				}
				case 'grayscl':
					return { kind: 'grayscale' }
				case 'biLevel': {
					const thresh = intValue(attr(child, 'thresh'))
					return { kind: 'biLevel', threshold: thresh === null ? null : thresh / 100000 }
				}
				case 'alphaModFix': {
					const amt = intValue(attr(child, 'amt'))
					return { kind: 'alphaModFix', amount: amt === null ? 1 : amt / 100000 }
				}
			}
		}
		return null
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
			warn(
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

	/**
	 * The connector's **start**-point shape binding (`p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn`),
	 * or `null` when the start point is unbound (a bare `p:cNvCxnSpPr`, i.e. a
	 * connector placed by static endpoint geometry). Mirrors the write API's
	 * `startShape`/`startShapeIdx` split; see {@link endConnection} for the other end.
	 */
	get startConnection(): ConnectionSite | null {
		return this.#connection('a:stCxn')
	}

	/**
	 * The connector's **end**-point shape binding (`p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn`),
	 * or `null` when the end point is unbound. See {@link startConnection}.
	 */
	get endConnection(): ConnectionSite | null {
		return this.#connection('a:endCxn')
	}

	/** Decode one `a:stCxn` / `a:endCxn` binding, resolving its `@id` to a slide shape. */
	#connection(qname: string): ConnectionSite | null {
		const nvCxnSpPr = firstChild(this.element, 'p:nvCxnSpPr')
		const cNvCxnSpPr = nvCxnSpPr && firstChild(nvCxnSpPr, 'p:cNvCxnSpPr')
		const cxn = cNvCxnSpPr && firstChild(cNvCxnSpPr, qname)
		if (!cxn) return null
		const shapeId = intValue(attr(cxn, 'id'))
		const siteIndex = intValue(attr(cxn, 'idx'))
		// CT_Connection requires both @id and @idx; an unparseable pair degrades to null
		// rather than a half-populated site.
		if (shapeId === null || siteIndex === null) return null
		return { shapeId, siteIndex, boundShape: this.slide.shapeByIdDeep(shapeId) ?? null }
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

	/** Whether this frame hosts a classic chart (`a:graphicData/@uri` is the chart URI). */
	get hasChart(): boolean {
		return this.#graphicDataUri() === A_CHART_URI
	}

	/** Whether this frame hosts a chartEx chart (`a:graphicData/@uri` is the chartEx URI). */
	get hasChartEx(): boolean {
		return this.#graphicDataUri() === A_CHARTEX_URI
	}

	/** The hosted table, or `null` when this frame is not a table. */
	get table(): Table | null {
		if (!this.hasTable) return null
		const graphicData = this.#graphicData()
		const tbl = graphicData && firstChild(graphicData, 'a:tbl')
		return tbl ? new Table(tbl, this.slide.part, this.slide.themeContext(), this.slide.presentation.opc) : null
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

	/**
	 * The hosted chartEx chart (waterfall/funnel/treemap/…), or `null` when this
	 * frame is not a chartEx chart or its part is missing. The reference child is
	 * `cx:chart` (not the classic `c:chart`), carrying the MS `chartEx` rel id.
	 */
	get chartEx(): ChartEx | null {
		if (!this.hasChartEx) return null
		const graphicData = this.#graphicData()
		const chartRef = graphicData && firstChild(graphicData, 'cx:chart')
		const relId = chartRef && attr(chartRef, 'r:id')
		if (!relId) return null
		const partName = this.slide.relationships.resolveTarget(relId)
		const part = this.slide.presentation.opc.part(partName)
		return part ? new ChartEx(part) : null
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
	get shapes(): AnyShape[] {
		return buildShapes(this.element, this.slide)
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
export function wrapShapeElement(element: Element, slide: Slide): AnyShape | null {
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
 * The richer of an `mc:AlternateContent`'s two branches: the first `mc:Choice`
 * that carries a shape element (what PowerPoint renders when it understands the
 * required feature), falling back to `mc:Fallback`. PowerPoint wraps a shape this
 * way when its full form needs a feature namespace a plain consumer lacks — a
 * chartEx chart's `p:graphicFrame`, a zoom frame, an inline-math shape. Returns
 * the wrapped inner shape, or `null` when neither branch holds one.
 */
function unwrapAlternateContent(altContent: Element, slide: Slide): AnyShape | null {
	const branches = [...getElements(altContent, 'mc:Choice'), ...getElements(altContent, 'mc:Fallback')]
	for (const branch of branches) {
		for (let node = branch.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const shape = wrapShapeElement(node as Element, slide)
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
export function buildShapes(parent: Element, slide: Slide): AnyShape[] {
	const shapes: AnyShape[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const element = node as Element
		if (element.namespaceURI === OOXML_NS.mc && element.localName === 'AlternateContent') {
			const unwrapped = unwrapAlternateContent(element, slide)
			if (unwrapped) shapes.push(unwrapped)
			continue
		}
		const shape = wrapShapeElement(element, slide)
		if (shape) shapes.push(shape)
	}
	return shapes
}
