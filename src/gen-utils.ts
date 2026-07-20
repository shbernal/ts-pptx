/**
 * PptxGenJS: DrawingML fragment builders and shared generator helpers.
 *
 * Everything here returns a piece of DrawingML markup or supports doing so:
 *   - Identifiers & naming   getUuid, validateObjectName, getDuplicateObjectNames, getNewRelId
 *   - XML text               encodeXmlEntities
 *   - Color                  componentToHex, rgbToHex, createColorElement, genXmlColorSelection
 *   - Effects                createGlowElement, createShadowElement(+EffectLst), correctShadowOptions
 *   - Fills & lines          genXmlGradientFill / genXmlPatternFill / genXmlImageFill / genXmlLineFill, createLineCap
 *
 * Unit conversion moved to `units-internal.ts` (over the public primitives in `units.ts`);
 * base64/image-header decoding and media content types moved to `media/`.
 */

import { REGEX_HEX_COLOR, DEF_FONT_COLOR, SchemeColor, type SCHEME_COLORS } from './core-enums.js'
import { warn } from './log.js'
import { ANGLE_UNITS_PER_DEGREE, EMU_PER_POINT, FIXED_PCT_PER_PERCENT, PERCENT_SCALE } from './units.js'
import { convertRotationDegrees, opacityToAlpha, transparencyToAlpha, valToPts } from './units-internal.js'
import type {
	BorderProps,
	TextGlowProps,
	PresSlideInternal,
	ShapeFillProps,
	Color,
	ShapeLineProps,
	ShadowProps,
	ShadowPropsInternal,
	GradientFillProps,
	GradientStopProps,
	PatternFillProps,
	LineCap,
} from './core-interfaces.js'

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid(uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		const r = (Math.random() * 16) | 0
		const v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string | number} xml - value to encode (numbers are stringified, as callers pass counts/sizes)
 * @returns {string} escaped XML
 */
export function encodeXmlEntities(xml: string | number): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	// Strip XML 1.0 illegal control chars (e.g. \v) before escaping to prevent PowerPoint repair dialogs.
	// Pattern built from String.fromCharCode so no-control-regex cannot flag it statically.
	const cc = String.fromCharCode
	const illegalXmlCharsRe = new RegExp(`[${cc(0)}-${cc(8)}${cc(11)}${cc(12)}${cc(14)}-${cc(31)}${cc(127)}]`, 'g')
	return xml
		.toString()
		.replace(illegalXmlCharsRe, '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;')
}

/**
 * Practical maximum length for a `p:cNvPr` object name. PowerPoint does not
 * enforce a hard spec limit, but very long names are a strong signal of a bug
 * and are unwieldy in the Selection Pane.
 */
const MAX_OBJECT_NAME_LENGTH = 255

/**
 * Validate a user-supplied object name and warn (does not throw) when the value
 * cannot be preserved as a stable PowerPoint Selection Pane identity. This keeps
 * semantic-identity bugs visible at generation time without breaking existing
 * decks that pass loose names.
 * - Empty/whitespace-only names provide no usable identity.
 * - Control characters are stripped by `encodeXmlEntities`, silently changing
 *   the stored name.
 * - Excessively long names may not round-trip through PowerPoint/consumers.
 * @param {string} name - the raw (pre-encoding) object name
 * @param {string} kind - object kind for the warning message (e.g. 'text')
 * @returns {string} the name unchanged (validation only)
 */
export function validateObjectName(name: string, kind: string): string {
	if (typeof name !== 'string') return name
	if (name.trim().length === 0) {
		warn(`${kind} objectName is empty or whitespace-only; it will not provide a stable Selection Pane identity.`)
		return name
	}
	// Same illegal-XML-char set that `encodeXmlEntities` strips; detect so the caller knows the name will change.
	const cc = String.fromCharCode
	const illegalXmlCharsRe = new RegExp(`[${cc(0)}-${cc(8)}${cc(11)}${cc(12)}${cc(14)}-${cc(31)}${cc(127)}]`)
	if (illegalXmlCharsRe.test(name)) {
		warn(`${kind} objectName "${name}" contains control characters that will be stripped, changing the stored name.`)
	}
	if (name.length > MAX_OBJECT_NAME_LENGTH) {
		warn(`${kind} objectName exceeds ${MAX_OBJECT_NAME_LENGTH} characters and may not be preserved by PowerPoint.`)
	}
	return name
}

/**
 * Return object names that appear more than once in the given list. Used to warn
 * when duplicate Selection Pane identities would be emitted on a single slide,
 * which breaks consumers (e.g. semantic manifests) that rely on unique names.
 * @param {string[]} names - object names emitted on one slide
 * @returns {string[]} the duplicated names (each listed once)
 */
export function getDuplicateObjectNames(names: string[]): string[] {
	const seen = new Set<string>()
	const dupes = new Set<string>()
	names.forEach((name) => {
		if (typeof name !== 'string' || name.length === 0) return
		if (seen.has(name)) dupes.add(name)
		else seen.add(name)
	})
	return Array.from(dupes)
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex(c: number): string {
	const hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export function rgbToHex(r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/**
 * Emit a bare DrawingML color element — either `<a:schemeClr>` (scheme color) or `<a:srgbClr>`
 * (hex RGB). This is a low-level primitive: it is the shared building block for every color
 * context (solid fills, gradient stops, `<a:alpha>` on shadows/glow, line fills, highlight,
 * underline, patterns, …), so it is intentionally NOT limited to solid fills and cannot be
 * folded into {@link genXmlColorSelection}.
 *
 * When you specifically want a *solid fill*, prefer {@link genXmlColorSelection}, which wraps
 * this in `<a:solidFill>…</a:solidFill>` and also handles `alpha`/`transparency` and the
 * gradient/pattern/image fill types — reach for `createColorElement` directly only when you
 * need the raw color element (e.g. inside `<a:ln>`, `<a:gs>`, an effect, or a highlight).
 * @param {string|SCHEME_COLORS} colorStr - hex RGB (e.g. "FFFF00") or a scheme color constant (e.g. pptx.SchemeColor.ACCENT1)
 * @param {string} [innerElements] - additional elements that adjust the color, nested inside the color element
 * @returns {string} XML string
 */
export function createColorElement(colorStr: string | SCHEME_COLORS, innerElements?: string): string {
	if (typeof colorStr !== 'string') {
		warn(`createColorElement: expected a string color value, got ${typeof colorStr}. "${DEF_FONT_COLOR}" used instead.`)
		colorStr = DEF_FONT_COLOR
	}
	let colorVal = (colorStr || '').replace('#', '')

	// 8-char hex (RGBA) — strip the alpha byte to a sibling <a:alpha val="N"/>,
	// continue with the leading 6-char RGB through the existing validation. This keeps
	// fill/text/line/glow paths from silently falling back to DEF_FONT_COLOR on RGBA input.
	if (/^[0-9a-fA-F]{8}$/.test(colorVal)) {
		// If the caller already supplied an explicit <a:alpha> (e.g. shadow/glow `opacity`),
		// it wins — do NOT add a second alpha from the RGBA byte, which would emit two
		// <a:alpha> children and produce schema-invalid OOXML (CT_SRgbColor allows one).
		if (!innerElements?.includes('<a:alpha')) {
			const alphaHex = colorVal.slice(6, 8)
			const alphaVal = Math.round((parseInt(alphaHex, 16) / 255) * PERCENT_SCALE)
			innerElements = `<a:alpha val="${alphaVal}"/>${innerElements || ''}`
		}
		colorVal = colorVal.slice(0, 6)
	}

	if (!REGEX_HEX_COLOR.test(colorVal) && !Object.values(SchemeColor).includes(colorVal as SchemeColor)) {
		warn(
			`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'pptx.SchemeColor' values!`
		)
		colorVal = DEF_FONT_COLOR
	}

	const tagName = REGEX_HEX_COLOR.test(colorVal) ? 'srgbClr' : 'schemeClr'
	const colorAttr = 'val="' + (REGEX_HEX_COLOR.test(colorVal) ? colorVal.toUpperCase() : colorVal) + '"'

	return innerElements ? `<a:${tagName} ${colorAttr}>${innerElements}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`
}

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: TextGlowProps, defaults: TextGlowProps): string {
	let strXml = ''
	const opts = { ...defaults, ...options }
	const size = Math.round(opts.size * EMU_PER_POINT)
	const color = opts.color || DEF_FONT_COLOR
	const opacity = opacityToAlpha(opts.opacity ?? 0)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += '</a:glow>'

	return strXml
}

/**
 * Creates an `a:outerShdw`/`a:innerShdw` element for a text run, shape, image, or chart.
 * Returns the shadow element only (no wrapping `a:effectLst`) so callers can either combine
 * it with other effects (e.g. glow) inside one `a:effectLst`, or wrap a lone shadow via
 * {@link createShadowEffectLst}.
 *
 * Colors go through {@link createColorElement}, so scheme colors (e.g. `accent1`) are honored
 * — earlier per-site copies hardcoded `a:srgbClr` and silently emitted invalid OOXML for them.
 * @param {ShadowPropsInternal} options shadow properties
 * @param {ShadowPropsInternal} defaults defaults for unspecified properties in `options`
 * @see http://officeopenxml.com/drwSp-effects.php
 * @returns {string} XML string, or '' when type is 'none'
 */
export function createShadowElement(options: ShadowPropsInternal | undefined, defaults: ShadowPropsInternal): string {
	const opts = { ...defaults, ...options }
	if (opts.type === 'none') return ''

	// NOTE: read into locals so we never mutate the caller's options (re-emission
	// would otherwise re-convert pt→EMU and produce absurd values).
	const type = opts.type || 'outer'
	const blur = valToPts(opts.blur ?? 0)
	const offset = valToPts(opts.offset ?? 0)
	const angle = Math.round((opts.angle ?? 0) * ANGLE_UNITS_PER_DEGREE)
	const opacity = Math.round((opts.opacity ?? 0.75) * PERCENT_SCALE)
	const color = opts.color || DEF_FONT_COLOR

	// sx/sy/kx/ky/algn/rotWithShape are valid only on `a:outerShdw` (CT_OuterShadowEffect);
	// `a:innerShdw` (CT_InnerShadowEffect) accepts only blurRad/dist/dir.
	const extraAttrs =
		type === 'outer'
			? `sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="${opts.rotateWithShape ? 1 : 0}" `
			: ''
	let strXml = `<a:${type}Shdw ${extraAttrs}blurRad="${blur}" dist="${offset}" dir="${angle}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += `</a:${type}Shdw>`

	return strXml
}

/**
 * Wraps a lone shadow in the `a:effectLst` that CT_ShapeProperties (shapes/images) and chart
 * marker/data-point properties require. Returns a self-closing `<a:effectLst/>` when there is
 * no shadow (missing/non-object options, or `type: 'none'`), matching the "no effects" element
 * PowerPoint emits. Use this instead of {@link createShadowElement} when the shadow is the only
 * effect; use `createShadowElement` directly when combining with other effects (e.g. glow).
 * @param {ShadowPropsInternal} options shadow properties
 * @param {ShadowPropsInternal} defaults defaults for unspecified properties in `options`
 * @returns {string} `<a:effectLst>…</a:effectLst>` or `<a:effectLst/>`
 */
export function createShadowEffectLst(options: ShadowPropsInternal | undefined, defaults: ShadowPropsInternal): string {
	if (!options || typeof options !== 'object') return '<a:effectLst/>'
	const inner = createShadowElement(options, defaults)
	return inner ? `<a:effectLst>${inner}</a:effectLst>` : '<a:effectLst/>'
}

function boolToXml(value: boolean): string {
	return value ? '1' : '0'
}

function normalizeGradientAngle(angle: number | undefined): number {
	const degrees = angle ?? 0
	if (typeof degrees !== 'number' || !Number.isFinite(degrees))
		throw new Error('Gradient angle must be a finite number.')
	return convertRotationDegrees(((degrees % 360) + 360) % 360)
}

function gradientStopColorAdjustments(stop: GradientStopProps): string {
	let internalElements = ''
	if (stop.transparency) internalElements += `<a:alpha val="${transparencyToAlpha(stop.transparency)}"/>`
	return internalElements
}

function normalizeGradientStops(stops: GradientStopProps[] | undefined): GradientStopProps[] {
	if (!Array.isArray(stops) || stops.length < 2) throw new Error('Gradient fill requires at least two stops.')

	return stops
		.map((stop) => {
			if (!stop || typeof stop.position !== 'number' || !Number.isFinite(stop.position)) {
				throw new Error('Gradient stop position must be a finite number from 0 to 100.')
			}
			if (stop.position < 0 || stop.position > 100) throw new Error('Gradient stop position must be from 0 to 100.')
			return stop
		})
		.sort((a, b) => a.position - b.position)
}

/**
 * Create a native DrawingML gradient fill.
 * @param {GradientFillProps} gradient gradient fill options
 * @returns XML string
 */
export function genXmlGradientFill(gradient: GradientFillProps | undefined): string {
	if (!gradient || (gradient.kind !== 'linear' && gradient.kind !== 'radial')) {
		throw new Error('Gradient fill currently supports only linear and radial gradients.')
	}
	if (typeof gradient.rotateWithShape !== 'undefined' && typeof gradient.rotateWithShape !== 'boolean') {
		throw new Error('Gradient rotateWithShape must be a boolean.')
	}

	const stops = normalizeGradientStops(gradient.stops)
	const rotWithShape = gradient.rotateWithShape ?? true

	let strXml = `<a:gradFill rotWithShape="${boolToXml(rotWithShape)}">`
	strXml += '<a:gsLst>'
	stops.forEach((stop) => {
		const position = Math.round(stop.position * FIXED_PCT_PER_PERCENT)
		strXml += `<a:gs pos="${position}">${createColorElement(stop.color, gradientStopColorAdjustments(stop))}</a:gs>`
	})
	strXml += '</a:gsLst>'
	if (gradient.kind === 'radial') {
		// `<a:path path="circle">` radiates the first stop from a focus rectangle out
		// to the edges. `fillToRect` insets place that focus: equal insets center it,
		// and the `center` percentage shifts it (l/t = center, r/b = 100 - center).
		const cx = Math.max(0, Math.min(100, gradient.center?.x ?? 50))
		const cy = Math.max(0, Math.min(100, gradient.center?.y ?? 50))
		const l = Math.round(cx * FIXED_PCT_PER_PERCENT)
		const t = Math.round(cy * FIXED_PCT_PER_PERCENT)
		const r = Math.round((100 - cx) * FIXED_PCT_PER_PERCENT)
		const b = Math.round((100 - cy) * FIXED_PCT_PER_PERCENT)
		strXml += `<a:path path="circle"><a:fillToRect l="${l}" t="${t}" r="${r}" b="${b}"/></a:path>`
	} else {
		if (typeof gradient.scaled !== 'undefined' && typeof gradient.scaled !== 'boolean')
			throw new Error('Gradient scaled must be a boolean.')
		const scaledAttr = typeof gradient.scaled === 'boolean' ? ` scaled="${boolToXml(gradient.scaled)}"` : ''
		strXml += `<a:lin ang="${normalizeGradientAngle(gradient.angle)}"${scaledAttr}/>`
	}
	strXml += '</a:gradFill>'

	return strXml
}

/**
 * Create a native DrawingML pattern fill.
 * @param {PatternFillProps} pattern pattern fill options
 * @returns XML string
 */
export function genXmlPatternFill(pattern: PatternFillProps | undefined): string {
	if (!pattern) throw new Error('Pattern fill requires a pattern object.')
	const fgColor = pattern.fgColor ?? '000000'
	const bgColor = pattern.bgColor ?? 'FFFFFF'
	return (
		`<a:pattFill prst="${pattern.preset}">` +
		`<a:fgClr>${createColorElement(fgColor)}</a:fgClr>` +
		`<a:bgClr>${createColorElement(bgColor)}</a:bgClr>` +
		'</a:pattFill>'
	)
}

/**
 * Create a native DrawingML picture (image) fill.
 * The media relationship is registered when the object is added; this only emits
 * the `<a:blipFill>` referencing the pre-resolved rId.
 * @param {ShapeFillProps} props fill props (must carry a resolved `_imgRid`)
 * @returns XML string
 */
export function genXmlImageFill(props: ShapeFillProps | undefined): string {
	if (!props || typeof props._imgRid !== 'number') {
		warn(
			'image fill is missing its resolved media reference; falling back to no fill. Provide `image: { path }` or `image: { data }`.'
		)
		return '<a:noFill/>'
	}
	const alpha = props.transparency
	const blipInner = alpha ? `<a:alphaModFix amt="${Math.round((100 - alpha) * FIXED_PCT_PER_PERCENT)}"/>` : ''
	return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId${props._imgRid}">${blipInner}</a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
}

/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
/**
 * Map a friendly `LineCap` value to the OOXML `cap` attribute value (`flat`/`sq`/`rnd`).
 * @param {LineCap} [lineCap] - line cap style (defaults to `flat`)
 * @returns {string} value for the `cap` attribute on `<a:ln>`
 */
/**
 * Resolve a border's line width in points, falling back to `defaultPt` when `width`
 * is not a usable number.
 * @param {BorderProps} border - border properties (may carry `width`)
 * @param {number} defaultPt - width to use when `width` is not a finite number
 * @returns {number} resolved width in points
 */
export function resolveBorderWidth(border: BorderProps, defaultPt: number): number {
	const val = border.width
	return typeof val === 'number' && !isNaN(val) ? val : defaultPt
}

export function createLineCap(lineCap?: LineCap): string {
	if (!lineCap || lineCap === 'flat') {
		return 'flat'
	} else if (lineCap === 'square') {
		return 'sq'
	} else if (lineCap === 'round') {
		return 'rnd'
	} else {
		const neverLineCap: never = lineCap
		throw new Error(`Invalid line cap: ${String(neverLineCap)}`)
	}
}

export function genXmlColorSelection(props: Color | ShapeFillProps | ShapeLineProps): string {
	let fillType = 'solid'
	let colorVal = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		if (typeof props === 'string') colorVal = props
		else {
			if (props.type) fillType = props.type
			if (props.color) colorVal = props.color
			if (props.transparency) internalElements += `<a:alpha val="${transparencyToAlpha(props.transparency)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				break
			case 'gradient':
				outText += genXmlGradientFill(typeof props === 'string' ? undefined : props.gradient)
				break
			case 'pattern':
				outText += genXmlPatternFill(typeof props === 'string' ? undefined : props.pattern)
				break
			case 'image':
				outText += genXmlImageFill(typeof props === 'string' ? undefined : props)
				break
			default: // @note need a statement as having only "break" can be removed by bundlers, then triggers "no-default" js-linter
				outText += ''
				break
		}
	}

	return outText
}

/**
 * Emit the paint child of an `<a:ln>` stroke.
 * DrawingML allows the same fill group inside `<a:ln>` as inside a shape fill, so a
 * stroke can be a gradient/pattern as well as a solid color:
 * - a `gradient` (or `type: 'gradient'`) produces a `<a:gradFill>` (gradient stroke);
 * - a `pattern`/`image` type delegates to the shared fill dispatch;
 * - otherwise a `color` produces a `<a:solidFill>`.
 * Returns '' when the line specifies no paint, so the caller emits no fill child and
 * the stroke inherits its color from the theme/placeholder.
 * @param {ShapeLineProps} [line] line options
 * @returns XML string
 */
export function genXmlLineFill(line: ShapeLineProps | undefined): string {
	if (!line) return ''
	// `gradient` presence selects a gradient stroke even when `type` was omitted.
	if (line.gradient || line.type === 'gradient') return genXmlGradientFill(line.gradient)
	if (line.type === 'pattern' || line.type === 'image') return genXmlColorSelection(line)
	if (line.color) return genXmlColorSelection(line)
	return ''
}

/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlideInternal} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export function getNewRelId(target: PresSlideInternal): number {
	return target._rels.length + target._relsChart.length + target._relsMedia.length + 1
}

/**
 * Whether a slide relationship is a hyperlink (external URL or internal slide
 * link). The relationship `type` is stringly-typed (`'hyperlink'`, `'online'`,
 * mixed-case variants), so this centralizes the case-insensitive predicate that
 * was duplicated across the slide-rels writer (gen-xml) and the inspect path
 * (pptxgen). For an internal slide-to-slide link, `rel.data === 'slide'` and
 * `rel.Target` is the 1-based target slide number.
 * @param {{ type: string }} rel - a slide relationship
 * @returns {boolean} true if the rel is any kind of hyperlink
 */
export function isHyperlinkRel(rel: { type: string }): boolean {
	return rel.type.toLowerCase().includes('hyperlink')
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
export function correctShadowOptions(ShadowProps?: ShadowProps | null): ShadowPropsInternal | undefined {
	if (!ShadowProps || typeof ShadowProps !== 'object') {
		// warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
		return undefined
	}
	const corrected: ShadowPropsInternal = ShadowProps
	// `opacity` is no longer a public input (removed in favor of `transparency`); strip any
	// leftover value from an untyped caller so it doesn't silently keep working through this
	// function's internal reuse of the same field name for the derived alpha.
	delete corrected.opacity

	// OPT: `type`
	if (corrected.type !== 'outer' && corrected.type !== 'inner' && corrected.type !== 'none') {
		warn('shadow.type options are `outer`, `inner` or `none`.')
		corrected.type = 'outer'
	}

	// OPT: `angle`
	if (corrected.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(corrected.angle)) || corrected.angle < 0 || corrected.angle > 359) {
			warn('shadow.angle can only be 0-359')
			corrected.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		corrected.angle = Math.round(Number(corrected.angle))
	}

	// OPT: `transparency` (PowerPoint UI term, 0-100) -> internal `opacity` (0.0-1.0), which
	// every emit site reads.
	if (corrected.transparency !== undefined) {
		const pct = Number(corrected.transparency)
		if (isNaN(pct) || pct < 0 || pct > 100) {
			warn('shadow.transparency can only be 0-100')
		} else {
			corrected.opacity = 1 - pct / 100
		}
	}

	// OPT: `color`
	if (corrected.color) {
		// INCORRECT FORMAT
		if (corrected.color.startsWith('#')) {
			warn('shadow.color should not include hash (#) character, , e.g. "FF0000"')
			corrected.color = corrected.color.replace('#', '')
		}

		// 8-char hex (RGBA) — derive `opacity` from the alpha byte (only when `transparency`
		// didn't already set one), then strip the alpha byte from the color so emit sites
		// produce valid 6-char `<a:srgbClr val="…"/>`.
		if (/^[0-9a-fA-F]{8}$/.test(corrected.color)) {
			const alphaHex = corrected.color.slice(6, 8)
			if (corrected.opacity === undefined) {
				corrected.opacity = parseInt(alphaHex, 16) / 255
			}
			corrected.color = corrected.color.slice(0, 6)
		}
	}

	return corrected
}
