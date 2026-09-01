/**
 * The abstract base every shape kind extends.
 *
 * It owns everything a shape has regardless of kind: non-visual identity (id / name / hidden /
 * description / decorative), geometry through the subclass-supplied `a:xfrm`, and the `p:spPr`
 * surface — fill, line, gradient, effects — resolved against the host's theme. Kind-specific
 * reads live on the subclasses in the sibling modules.
 *
 * A proxy holds a back-reference to its owning {@link ShapeHost} — a `Slide`, a `SlideLayout` or a
 * `SlideMaster` — so it can resolve relationships and theme colours, and so a mutation can mark
 * that host's part dirty.
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
} from '../../oxml/dom.js'
import { composeGroupFrame, type GroupTransform } from './group-transform.js'
import { FILL_CHOICES, normalizeHex, setSolidFill, solidFillColor } from '../../oxml/fill.js'
import {
	resolveColorElement,
	resolveInheritedFrame,
	resolveSolidFillColor,
	resolveStyleFillColor,
	resolveStyleLineColor,
	type PlaceholderRef,
	type ResolvedColor,
	type ResolvedFrame,
} from '../theme-context.js'
import { readGradientFill, readGradientStops, type GradientFill, type GradientStop } from '../gradient.js'
import { readPictureFill, type PictureFill } from '../picture-fill.js'
import { readPatternFill } from '../pattern-fill.js'
import { TextFrame } from '../text.js'
import type { ShapeHost } from './host.js'
import {
	childElements,
	emuFrom,
	getOrAddSpPrXfrm,
	nonVisualCNvPr,
	LN_FILL_AFTER,
	SHAPE_AFTER_SPPR,
	SPPR_FILL_AFTER,
	SPPR_LN_AFTER,
	type ShapeProperties,
} from './oxml.js'
import { readBox, rotationDegrees, toEmu, transformFlipH, transformFlipV } from './geometry.js'
import type {
	AbsoluteFrame,
	Glow,
	InnerShadow,
	LineEnd,
	LineEnds,
	OuterShadow,
	PatternFill,
	Reflection,
	ShapeType,
	SoftEdge,
} from './types.js'
import { InternalError, PackageReadError, UnsupportedFeatureError } from '../../../errors.js'
import { ANGLE_UNITS_PER_DEGREE, EMU_PER_POINT, PERCENT_SCALE } from '../../../units.js'
import { ptFromEmu } from '../coords.js'

// Microsoft's "decorative" accessibility extension: p:cNvPr/a:extLst/a:ext
// (uri {C183D7F6-B498-43B3-948B-1728B52AA6E4}) / adec:decorative. Confirmed
// against real PowerPoint output ("Mark as decorative").
const ADEC_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative'

/** Common base for every shape in a shape tree — a slide's, a layout's, or a master's. */
export abstract class Shape {
	constructor(
		protected readonly element: Element,
		/**
		 * The part that owns this shape's tree — a `Slide`, a `SlideLayout`, or a
		 * `SlideMaster`. Narrow it with `instanceof` when you need the concrete class;
		 * {@link ShapeHost.partName} tells the tiers apart without one.
		 */
		readonly host: ShapeHost
	) {}

	/** Which concrete shape kind this is. */
	abstract readonly shapeType: ShapeType

	/**
	 * The transform element carrying this shape's geometry, or `null` if inherited.
	 *
	 * Defaults to `p:spPr/a:xfrm`, which is where three of the four shape kinds keep it
	 * (`p:sp`, `p:pic`, `p:cxnSp`). {@link GraphicFrame} overrides it: a `p:graphicFrame`
	 * has no `p:spPr` and carries a `p:xfrm` of its own directly.
	 */
	protected xfrm(): Element | null {
		const spPr = firstChild(this.element, 'p:spPr')
		return spPr ? firstChild(spPr, 'a:xfrm') : null
	}

	/** The transform element, creating it (and its container) in document order if absent. */
	protected getOrAddXfrm(): Element {
		return getOrAddSpPrXfrm(this.element)
	}

	/**
	 * Mark the owning host's part dirty so `save()` reserializes it. Public
	 * because {@link element_} hands out the live DOM node: the hatch and the
	 * obligation that comes with it belong on the same object.
	 */
	markDirty(): void {
		this.host.part.markDirty()
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
		if (!cNvPr)
			throw new PackageReadError('shape/no-non-visual-properties', 'cannot set description: shape has no p:cNvPr')
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
		return resolveInheritedFrame(ph, this.host.themeContext())
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
	 * `true` when the shape sets an explicit no-fill (`spPr/a:noFill`) — a
	 * deliberately transparent surface. The fill-side counterpart of
	 * {@link lineNoFill}, and the only accessor that separates it from a shape
	 * carrying no fill child at all (one inheriting through `p:style/a:fillRef`):
	 * every other fill accessor — {@link fillColor}, {@link fillSchemeColor},
	 * {@link resolvedFill}, {@link gradientFill}, {@link patternFill},
	 * {@link pictureFill} — reports `null` for both. The two paint completely
	 * differently, so a consumer that cannot tell them apart paints a transparent
	 * shape in the theme's accent colour.
	 */
	get fillNoFill(): boolean {
		const props = this.properties()
		return !!(props && firstChild(props, 'a:noFill'))
	}

	/**
	 * Set an explicit `<a:noFill/>` on the shape — a transparent surface. This is
	 * distinct from clearing the fill (`fillColor = null`), which removes the
	 * `a:solidFill` and lets the fill inherit from the shape's style/placeholder.
	 * Read it back with {@link fillNoFill}.
	 */
	noFill(): void {
		this.#requireFillSupport()
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
		return ptFromEmu(w)
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
	 * The line end cap (`spPr/a:ln/@cap`) as the raw OOXML token — `'flat'`,
	 * `'rnd'` (round) or `'sq'` (square) — or `null` when unset (PowerPoint's
	 * default is `flat`). The write API authors this attribute through
	 * `ShapeLineProps.cap`, so without the accessor a deck this library produced
	 * could not be read back without losing it.
	 *
	 * Not cosmetic on a thick dashed rule: the cap decides whether each dash reads
	 * as a rectangle or a lozenge, and it extends every dash by the stroke width.
	 * SVG's `stroke-linecap` is the exact equivalent (`flat`→`butt`, `rnd`→`round`,
	 * `sq`→`square`).
	 */
	get lineCap(): string | null {
		const ln = this.#line()
		return ln ? (attr(ln, 'cap') ?? null) : null
	}

	/**
	 * The line alignment (`spPr/a:ln/@algn`) as the raw OOXML token — `'ctr'`
	 * (centred on the shape's outline) or `'in'` (inset, drawn wholly inside it) —
	 * or `null` when unset. It shifts a thick outline by half its width, so it
	 * changes where the border sits relative to the fill.
	 */
	get lineAlign(): string | null {
		const ln = this.#line()
		return ln ? (attr(ln, 'algn') ?? null) : null
	}

	/**
	 * Preset geometry name (`spPr/a:prstGeom/@prst`, e.g. `rect`), or `null` for
	 * custom geometry or none. Not an auto-shape-only property: PowerPoint gives a
	 * picture and a connector a preset geometry too (a `p:pic` is `rect` unless it
	 * has been cropped to a shape), so it reads off whichever properties element
	 * this kind carries. A group has no geometry of its own and reads `null`.
	 */
	get presetGeometry(): string | null {
		const props = this.properties()
		const prstGeom = props && firstChild(props, 'a:prstGeom')
		return prstGeom ? attr(prstGeom, 'prst') : null
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
		return readGradientStops(container, this.host.themeContext())
	}

	/**
	 * Read the full `a:gradFill` (stops + linear angle / path shape) from a
	 * container (`spPr` for a fill or `a:ln` for a line stroke). `null` when the
	 * container has no gradient.
	 */
	#gradientFillIn(container: Element): GradientFill | null {
		return readGradientFill(container, this.host.themeContext())
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
	 * against the host's theme, or `null` when the shape has no outer shadow. The
	 * soft brand shadows the eye reads as "floating" panels live here and are
	 * invisible in geometry/fill alone.
	 */
	get shadow(): OuterShadow | null {
		const shdw = this.#effect('a:outerShdw')
		return shdw ? this.#readShadow(shdw) : null
	}

	/**
	 * The shape's **inner** shadow (`spPr/a:effectLst/a:innerShdw`), resolved against
	 * the host's theme, or `null` when the shape has no inner shadow. The inset
	 * counterpart of {@link shadow}: the write-side `shadow: { type: 'inner' }`
	 * emits it, and it is invisible in geometry/fill alone.
	 */
	get innerShadow(): InnerShadow | null {
		const shdw = this.#effect('a:innerShdw')
		return shdw ? this.#readShadow(shdw) : null
	}

	/**
	 * The shape's glow halo (`spPr/a:effectLst/a:glow`), resolved against the host's
	 * theme, or `null` when the shape has no glow. Same element the write-side text
	 * glow emits, so its {@link Glow.radiusPt} and colour round-trip.
	 */
	get glow(): Glow | null {
		const glow = this.#effect('a:glow')
		if (!glow) return null
		const out: Glow = { color: null }
		this.#applyEffectColor(out, firstChildElement(glow))
		const rad = intValue(attr(glow, 'rad'))
		if (rad !== null) out.radiusPt = rad / EMU_PER_POINT
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
		put('blurPt', 'blurRad', EMU_PER_POINT)
		put('offsetPt', 'dist', EMU_PER_POINT)
		put('angleDeg', 'dir', ANGLE_UNITS_PER_DEGREE)
		put('fadeAngleDeg', 'fadeDir', ANGLE_UNITS_PER_DEGREE)
		put('startAlpha', 'stA', PERCENT_SCALE)
		put('startPos', 'stPos', PERCENT_SCALE)
		put('endAlpha', 'endA', PERCENT_SCALE)
		put('endPos', 'endPos', PERCENT_SCALE)
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
		return { radiusPt: rad === null ? 0 : rad / EMU_PER_POINT }
	}

	/**
	 * The shape's pattern (hatch) fill (`spPr/a:pattFill`), or `null` when the fill
	 * is not a pattern. Surfaces the {@link PatternFill.preset} name and both
	 * colours resolved against the host's theme — the pattern counterpart of
	 * {@link resolvedFill}, which reports `null` for a non-solid fill and so drops a
	 * hatched surface entirely.
	 */
	get patternFill(): PatternFill | null {
		const props = this.properties()
		return props ? readPatternFill(props, this.host.themeContext()) : null
	}

	/**
	 * The shape's picture (image) fill (`spPr/a:blipFill`), or `null` when the fill
	 * is not a picture. The image-fill counterpart of {@link patternFill}: a shape
	 * whose *surface* is an image is not a {@link Picture} — it is an autoShape with
	 * a blip fill — and {@link resolvedFill} reports `null` for one, so without this
	 * an image-filled shape reads as unfilled. Carries the embedded image
	 * ({@link PictureFill.relId}/{@link PictureFill.partName}) plus the stretch/tile
	 * geometry.
	 */
	get pictureFill(): PictureFill | null {
		const props = this.properties()
		return props ? readPictureFill(props, this.host.relationships) : null
	}

	/** A named child of the shape's effect list (`spPr/a:effectLst/<qname>`), or `null`. */
	#effect(qname: string): Element | null {
		const props = this.properties()
		const effectLst = props && firstChild(props, 'a:effectLst')
		return effectLst ? firstChild(effectLst, qname) : null
	}

	/** Resolve `colorEl` against the theme and stamp `color`/`colorToken`/`alpha` onto an effect result. */
	#applyEffectColor(out: { color: string | null; colorToken?: string; alpha?: number }, colorEl: Element | null): void {
		const resolved = resolveColorElement(colorEl, this.host.themeContext())
		if (resolved) {
			out.color = resolved.effectiveHex
			if (resolved.alpha !== undefined) out.alpha = resolved.alpha
		}
		// A `schemeClr` with no `val` leaves `colorToken` off entirely, which is the read model's one
		// spelling of "not a theme colour" — the same invariant `compact()` keeps downstream.
		const token = colorEl && colorEl.localName === 'schemeClr' ? attr(colorEl, 'val') : null
		if (token !== null) out.colorToken = token
	}

	/** Decode a shadow element (`a:outerShdw`/`a:innerShdw` share the fields), resolving its colour. */
	#readShadow(shdw: Element): OuterShadow {
		const out: OuterShadow = { color: null }
		// `a:EG_ColorChoice` is a required, single-member group, so the colour element is the shadow's
		// only child and taking the first one is both correct and total — the same thing `glow` above
		// does. Naming `a:srgbClr` and `a:schemeClr` explicitly dropped the other four models on the
		// floor: `a:sysClr` resolves everywhere else in the read model, and this library emits
		// `a:prstClr` itself (`gen/slide/notes.ts`). `resolveColor` now answers for five of the six
		// (`a:scrgbClr` is the exception, and reports no colour rather than a guessed one).
		this.#applyEffectColor(out, firstChildElement(shdw))
		const blur = intValue(attr(shdw, 'blurRad'))
		const dist = intValue(attr(shdw, 'dist'))
		const dir = intValue(attr(shdw, 'dir'))
		if (blur !== null) out.blurPt = blur / EMU_PER_POINT
		if (dist !== null) out.offsetPt = dist / EMU_PER_POINT
		if (dir !== null) out.angleDeg = dir / ANGLE_UNITS_PER_DEGREE
		return out
	}

	/**
	 * The shape's solid fill resolved against the host's theme
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
		const ctx = this.host.themeContext()
		const props = this.properties()
		if (props && FILL_CHOICES.some((q) => firstChild(props, q))) return resolveSolidFillColor(props, ctx)
		return resolveStyleFillColor(this.element, ctx)
	}

	/**
	 * The shape's line/border solid fill resolved against the host's theme to a
	 * literal hex — the resolved counterpart of {@link lineColor}/{@link lineSchemeColor}.
	 * `null` when the shape has no `a:ln/a:solidFill` or it cannot be made literal.
	 * Like {@link resolvedFill}, the result carries `effectiveHex` (the base colour
	 * with its transforms applied) for the final rendered colour.
	 *
	 * When the shape carries no explicit `spPr/a:ln`, this falls back to the line
	 * the shape inherits from its `p:style` `a:lnRef` (the theme style matrix).
	 */
	get resolvedLine(): ResolvedColor | null {
		const ctx = this.host.themeContext()
		const ln = this.#line()
		return ln ? resolveSolidFillColor(ln, ctx) : resolveStyleLineColor(this.element, ctx)
	}

	/** The line element (`spPr/a:ln`), or `null` when absent. */
	#line(): Element | null {
		const props = this.properties()
		return props ? firstChild(props, 'a:ln') : null
	}

	/**
	 * Throw unless this shape kind has a fill to set. Both the explicit `<a:noFill/>` setter and
	 * the solid-fill one need it, and a kind that cannot be filled is a caller mistake rather than
	 * a no-op — silently accepting the call would report success for a shape that never changes.
	 */
	#requireFillSupport(): void {
		if (!this.supportsFill)
			throw new UnsupportedFeatureError(
				'shape/fill-unsupported',
				`${this.shapeType} shapes do not support a solid fill`
			)
	}

	#setFill(color: { qname: string; val: string } | null): void {
		if (color === null) {
			const props = this.properties()
			if (!props || !firstChild(props, 'a:solidFill')) return
			removeChildrenByQName(props, ['a:solidFill'])
			this.markDirty()
			return
		}
		this.#requireFillSupport()
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
		if (lnAfter === null)
			throw new UnsupportedFeatureError(
				'shape/line-unsupported',
				`${this.shapeType} shapes do not support a line colour`
			)
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
		if (!frame) throw new UnsupportedFeatureError('shape/no-text-frame', 'Shape has no text frame to set text on')
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
	 * Remove this shape from its parent (the host's shape tree, or an enclosing
	 * group) and mark the owning host's part dirty. The proxy is dead afterwards.
	 */
	delete(): void {
		const parent = this.element.parentNode
		if (!parent)
			throw new InternalError('oxml/node-has-no-parent', 'Shape is not attached to a parent and cannot be deleted')
		parent.removeChild(this.element)
		this.markDirty()
	}

	/** Escape hatch: the underlying shape element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.element
	}
}
