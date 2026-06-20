/**
 * PptxGenJS: Utility Methods
 */

import { REGEX_HEX_COLOR, DEF_FONT_COLOR, ONEPT, SchemeColor, SCHEME_COLORS } from './core-enums.js'
import { coordToEmu, inchesToEmu, type Emu } from './units.js'
import type { PresLayout, TextGlowProps, PresSlideInternal, ShapeFillProps, Color, ShapeLineProps, Coord, ShadowProps, GradientFillProps, GradientStopProps, PatternFillProps, LineCap } from './core-interfaces.js'

/**
 * Resolve a user `Coord` (x/y/w/h) to EMU — the single user-coordinate → EMU boundary.
 * - bare `number` → **inches** (no magnitude guessing); `"<n>%"` → percent of the slide axis;
 *   `"<n>in"`/`"<n>pt"`/`"<n>emu"` → explicit units (see {@link Coord} / {@link coordToEmu})
 * - `null`/`undefined` → 0 (callers may omit a coordinate)
 * - throws on a non-finite number rather than silently collapsing the object to zero size
 * @param {Coord|null|undefined} size - user coordinate
 * @param {'X' | 'Y'} xyDir - axis (selects slide width vs height for percentages)
 * @param {PresLayout} layout - presentation layout (EMU dimensions)
 * @returns {Emu} resolved EMU value
 */
export function getSmartParseNumber (size: Coord | null | undefined, xyDir: 'X' | 'Y', layout: PresLayout): Emu {
	if (size === null || size === undefined) return 0 as Emu

	// GUARD: A NaN/Infinity coordinate is always a mistake (commonly arithmetic on an
	// `undefined` layout dimension). Fail loud with a targeted hint instead of the generic
	// converter message, since this is the most common way a deck collapses to zero-size.
	if (typeof size === 'number' && !isFinite(size)) {
		throw new Error(
			`Invalid ${xyDir || 'coordinate'} value: expected a finite number but received ${String(size)}. ` +
				'This usually means a layout dimension was read from a missing property (e.g. `layout.width` returning `undefined`). ' +
				'Use `slide.width`/`slide.height` or `STANDARD_LAYOUTS.<NAME>.width`/`.height` (inches).'
		)
	}

	return coordToEmu(size, xyDir === 'Y' ? layout.height : layout.width)
}

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid (uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		const r = (Math.random() * 16) | 0
		const v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities (xml: string): string {
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
export function validateObjectName (name: string, kind: string): string {
	if (typeof name !== 'string') return name
	if (name.trim().length === 0) {
		console.warn(`Warning: ${kind} objectName is empty or whitespace-only; it will not provide a stable Selection Pane identity.`)
		return name
	}
	// Same illegal-XML-char set that `encodeXmlEntities` strips; detect so the caller knows the name will change.
	const cc = String.fromCharCode
	const illegalXmlCharsRe = new RegExp(`[${cc(0)}-${cc(8)}${cc(11)}${cc(12)}${cc(14)}-${cc(31)}${cc(127)}]`)
	if (illegalXmlCharsRe.test(name)) {
		console.warn(`Warning: ${kind} objectName "${name}" contains control characters that will be stripped, changing the stored name.`)
	}
	if (name.length > MAX_OBJECT_NAME_LENGTH) {
		console.warn(`Warning: ${kind} objectName exceeds ${MAX_OBJECT_NAME_LENGTH} characters and may not be preserved by PowerPoint.`)
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
export function getDuplicateObjectNames (names: string[]): string[] {
	const seen = new Set<string>()
	const dupes = new Set<string>()
	names.forEach(name => {
		if (typeof name !== 'string' || name.length === 0) return
		if (seen.has(name)) dupes.add(name)
		else seen.add(name)
	})
	return Array.from(dupes)
}

/**
 * Convert inches into EMU.
 * - accepts a number (inches) or a numeric/`"<n>in"` string
 * - no magnitude guessing: values are always treated as inches (use {@link coordToEmu} for
 *   user coordinates that may carry other units)
 * @param {number|string} inches - inches as number or string
 * @returns {Emu} EMU value
 */
export function inch2Emu (inches: number | string): Emu {
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return inchesToEmu(inches)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts (pt: number | string): number {
	const points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert a transparency percentage (0-100) into a schema-valid `<a:alpha>` value
 * (ST_PositiveFixedPercentage, 0-100000). Out-of-range transparency yields an
 * alpha that PowerPoint rejects as needing repair, so clamp into range and warn.
 */
export function transparencyToAlpha (transparency: number): number {
	const pct = Math.min(100, Math.max(0, transparency))
	if (pct !== transparency) console.warn(`Warning: transparency ${transparency} is outside the valid range 0-100; using ${pct}.`)
	return Math.round((100 - pct) * 1000)
}

/** Convert an opacity (0-1) into a schema-valid `<a:alpha>` value (0-100000); clamps + warns on out-of-range input. */
export function opacityToAlpha (opacity: number): number {
	const o = Math.min(1, Math.max(0, opacity))
	if (o !== opacity) console.warn(`Warning: opacity ${opacity} is outside the valid range 0-1; using ${o}.`)
	return Math.round(o * 100000)
}

/**
 * Convert a line width (points) to EMU clamped into ST_LineWidth (0..20116800 EMU,
 * i.e. 0-1584pt). Out-of-range widths make PowerPoint report the package as needing
 * repair, so clamp into range and warn.
 */
export function lineWidthToEmu (widthPts: number | string): number {
	const raw = valToPts(widthPts)
	const clamped = Math.min(20116800, Math.max(0, raw))
	if (clamped !== raw) console.warn(`Warning: line width ${widthPts} is outside the valid range 0-1584pt; using ${clamped / ONEPT}.`)
	return clamped
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 * @param {number} d degrees
 * @returns {number} calculated `rot` value
 */
export function convertRotationDegrees (d: number): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * 60000)
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex (c: number): string {
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
export function rgbToHex (r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/**  TODO: FUTURE: TODO-4.0:
 * @date 2022-04-10
 * @tldr this s/b a private method with all current calls switched to `genXmlColorSelection()`
 * @desc lots of code calls this method
 * @example [gen-charts.tx] `strXml += '<a:solidFill>' + createColorElement(seriesColor, `<a:alpha val="${Math.round(opts.chartColorsOpacity * 1000)}"/>`) + '</a:solidFill>'`
 * Thi sis wrong. We s/b calling `genXmlColorSelection()` instead as it returns `<a:solidfill>BLAH</a:solidFill>`!!
 */
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export function createColorElement (colorStr: string | SCHEME_COLORS, innerElements?: string): string {
	if (typeof colorStr !== 'string') {
		console.warn(`createColorElement: expected a string color value, got ${typeof colorStr}. "${DEF_FONT_COLOR}" used instead.`)
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
			const alphaVal = Math.round((parseInt(alphaHex, 16) / 255) * 100000)
			innerElements = `<a:alpha val="${alphaVal}"/>${innerElements || ''}`
		}
		colorVal = colorVal.slice(0, 6)
	}

	if (
		!REGEX_HEX_COLOR.test(colorVal) &&
		colorVal !== SchemeColor.background1 &&
		colorVal !== SchemeColor.background2 &&
		colorVal !== SchemeColor.text1 &&
		colorVal !== SchemeColor.text2 &&
		colorVal !== SchemeColor.accent1 &&
		colorVal !== SchemeColor.accent2 &&
		colorVal !== SchemeColor.accent3 &&
		colorVal !== SchemeColor.accent4 &&
		colorVal !== SchemeColor.accent5 &&
		colorVal !== SchemeColor.accent6
	) {
		console.warn(`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'pptx.SchemeColor' values!`)
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
export function createGlowElement (options: TextGlowProps, defaults: TextGlowProps): string {
	let strXml = ''
	const opts = { ...defaults, ...options }
	const size = Math.round(opts.size * ONEPT)
	const color = opts.color || DEF_FONT_COLOR
	const opacity = opacityToAlpha(opts.opacity ?? 0)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += '</a:glow>'

	return strXml
}

/**
 * Creates an `a:outerShdw`/`a:innerShdw` element for a text run or shape.
 * Returns the shadow element only (no wrapping `a:effectLst`) so callers can
 * combine it with other effects (e.g. glow) inside a single `a:effectLst`.
 * @param {ShadowProps} options shadow properties
 * @param {ShadowProps} defaults defaults for unspecified properties in `options`
 * @see http://officeopenxml.com/drwSp-effects.php
 * @returns {string} XML string, or '' when type is 'none'
 */
export function createShadowElement (options: ShadowProps, defaults: ShadowProps): string {
	const opts = { ...defaults, ...options }
	if (opts.type === 'none') return ''

	// NOTE: read into locals so we never mutate the caller's options (re-emission
	// would otherwise re-convert pt→EMU and produce absurd values).
	const type = opts.type || 'outer'
	const blur = valToPts(opts.blur ?? 0)
	const offset = valToPts(opts.offset ?? 0)
	const angle = Math.round((opts.angle ?? 0) * 60000)
	const opacity = Math.round((opts.opacity ?? 0.75) * 100000)
	const color = opts.color || DEF_FONT_COLOR

	const extraAttrs = type === 'outer' ? 'sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="0" ' : ''
	let strXml = `<a:${type}Shdw ${extraAttrs}blurRad="${blur}" dist="${offset}" dir="${angle}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += `</a:${type}Shdw>`

	return strXml
}

function boolToXml (value: boolean): string {
	return value ? '1' : '0'
}

function normalizeGradientAngle (angle: number | undefined): number {
	const degrees = angle ?? 0
	if (typeof degrees !== 'number' || !Number.isFinite(degrees)) throw new Error('Gradient angle must be a finite number.')
	return convertRotationDegrees(((degrees % 360) + 360) % 360)
}

function gradientStopColorAdjustments (stop: GradientStopProps): string {
	let internalElements = ''
	if (stop.alpha) internalElements += `<a:alpha val="${transparencyToAlpha(stop.alpha)}"/>` // DEPRECATED: @deprecated v3.3.0
	if (stop.transparency) internalElements += `<a:alpha val="${transparencyToAlpha(stop.transparency)}"/>`
	return internalElements
}

function normalizeGradientStops (stops: GradientStopProps[] | undefined): GradientStopProps[] {
	if (!Array.isArray(stops) || stops.length < 2) throw new Error('Gradient fill requires at least two stops.')

	return stops
		.map(stop => {
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
export function genXmlGradientFill (gradient: GradientFillProps | undefined): string {
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
	stops.forEach(stop => {
		const position = Math.round(stop.position * 1000)
		strXml += `<a:gs pos="${position}">${createColorElement(stop.color, gradientStopColorAdjustments(stop))}</a:gs>`
	})
	strXml += '</a:gsLst>'
	if (gradient.kind === 'radial') {
		// `<a:path path="circle">` radiates the first stop from a focus rectangle out
		// to the edges. `fillToRect` insets place that focus: equal insets center it,
		// and the `center` percentage shifts it (l/t = center, r/b = 100 - center).
		const cx = Math.max(0, Math.min(100, gradient.center?.x ?? 50))
		const cy = Math.max(0, Math.min(100, gradient.center?.y ?? 50))
		const l = Math.round(cx * 1000)
		const t = Math.round(cy * 1000)
		const r = Math.round((100 - cx) * 1000)
		const b = Math.round((100 - cy) * 1000)
		strXml += `<a:path path="circle"><a:fillToRect l="${l}" t="${t}" r="${r}" b="${b}"/></a:path>`
	} else {
		if (typeof gradient.scaled !== 'undefined' && typeof gradient.scaled !== 'boolean') throw new Error('Gradient scaled must be a boolean.')
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
export function genXmlPatternFill (pattern: PatternFillProps | undefined): string {
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
export function genXmlImageFill (props: ShapeFillProps | undefined): string {
	if (!props || typeof props._imgRid !== 'number') {
		console.warn('Warning: image fill is missing its resolved media reference; falling back to no fill. Provide `image: { path }` or `image: { data }`.')
		return '<a:noFill/>'
	}
	const alpha = props.transparency ?? props.alpha
	const blipInner = alpha ? `<a:alphaModFix amt="${Math.round((100 - alpha) * 1000)}"/>` : ''
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
export function createLineCap (lineCap?: LineCap): string {
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

export function genXmlColorSelection (props: Color | ShapeFillProps | ShapeLineProps): string {
	let fillType = 'solid'
	let colorVal = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		if (typeof props === 'string') colorVal = props
		else {
			if (props.type) fillType = props.type
			if (props.color) colorVal = props.color
			if (props.alpha) internalElements += `<a:alpha val="${transparencyToAlpha(props.alpha)}"/>` // DEPRECATED: @deprecated v3.3.0
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
export function genXmlLineFill (line: ShapeLineProps | undefined): string {
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
export function getNewRelId (target: PresSlideInternal): number {
	return target._rels.length + target._relsChart.length + target._relsMedia.length + 1
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
export function correctShadowOptions (ShadowProps?: ShadowProps | null): ShadowProps | undefined {
	if (!ShadowProps || typeof ShadowProps !== 'object') {
		// console.warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
		return
	}

	// OPT: `type`
	if (ShadowProps.type !== 'outer' && ShadowProps.type !== 'inner' && ShadowProps.type !== 'none') {
		console.warn('Warning: shadow.type options are `outer`, `inner` or `none`.')
		ShadowProps.type = 'outer'
	}

	// OPT: `angle`
	if (ShadowProps.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.angle)) || ShadowProps.angle < 0 || ShadowProps.angle > 359) {
			console.warn('Warning: shadow.angle can only be 0-359')
			ShadowProps.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.angle = Math.round(Number(ShadowProps.angle))
	}

	// OPT: `opacity`
	if (ShadowProps.opacity) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.opacity)) || ShadowProps.opacity < 0 || ShadowProps.opacity > 1) {
			console.warn('Warning: shadow.opacity can only be 0-1')
			ShadowProps.opacity = 0.75
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.opacity = Number(ShadowProps.opacity)
	}

	// OPT: `color`
	if (ShadowProps.color) {
		// INCORRECT FORMAT
		if (ShadowProps.color.startsWith('#')) {
			console.warn('Warning: shadow.color should not include hash (#) character, , e.g. "FF0000"')
			ShadowProps.color = ShadowProps.color.replace('#', '')
		}

		// 8-char hex (RGBA) — derive `opacity` from the alpha byte (only when caller
		// did not pass an explicit opacity), then strip the alpha byte from the color so
		// emit sites produce valid 6-char `<a:srgbClr val="…"/>`.
		if (/^[0-9a-fA-F]{8}$/.test(ShadowProps.color)) {
			const alphaHex = ShadowProps.color.slice(6, 8)
			if (ShadowProps.opacity === undefined) {
				ShadowProps.opacity = parseInt(alphaHex, 16) / 255
			}
			ShadowProps.color = ShadowProps.color.slice(0, 6)
		}
	}

	return ShadowProps
}

/**
 * Encode raw SVG markup as a base64 `image/svg+xml` data URI.
 * - lets callers pass inline SVG to `addImage({ svg })` without hand-rolling base64
 * - isomorphic and UTF-8 safe: uses the global `TextEncoder`/`btoa` (Node >=16, browsers)
 * @param {string} svg - SVG markup, e.g. `'<svg ...>...</svg>'`
 * @returns {string} a `data:image/svg+xml;base64,...` URI
 */
export function svgMarkupToDataUri (svg: string): string {
	const bytes = new TextEncoder().encode(svg)
	let binary = ''
	for (let i = 0; i < bytes.length; i++) {
		binary += String.fromCharCode(bytes[i])
	}
	return `data:image/svg+xml;base64,${btoa(binary)}`
}

/**
 * Decode a base64 image payload (raw base64 or a `data:` URI) to bytes.
 * - tolerant of the `data:[mime];base64,` prefix and of whitespace in the payload
 * @param {string} b64 - base64 string or data URI
 * @returns {Uint8Array | null} decoded bytes, or `null` when the payload is empty/undecodable
 */
export function decodeBase64ToBytes (b64: string): Uint8Array | null {
	if (!b64) return null
	// Strip any `data:...;base64,` prefix and surrounding whitespace
	const comma = b64.indexOf('base64,')
	const payload = (comma >= 0 ? b64.slice(comma + 'base64,'.length) : b64).replace(/\s/g, '')
	if (!payload) return null
	try {
		const binary = atob(payload)
		const bytes = new Uint8Array(binary.length)
		for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
		return bytes
	} catch {
		return null
	}
}

/**
 * Read the intrinsic dimensions of an image from its header bytes.
 * - synchronous: parses only file-format headers, never decodes pixels
 * - raster: PNG, JPEG, GIF, BMP, and WebP (VP8 / VP8L / VP8X) — natural pixels
 * - vector: SVG — intrinsic size from the root `<svg>` width/height or viewBox
 * - unrecognized formats return `null` (no measurable intrinsic size)
 *
 * Used by image `sizing: 'cover' | 'contain'` to compute an aspect-correct
 * `<a:srcRect>` crop from the *natural* image ratio rather than the displayed box.
 * @param {string} dataB64 - base64 image payload or `data:` URI
 * @returns {{ w: number, h: number } | null} natural size, or `null` when unmeasurable
 */
export function getImageSizeFromBase64 (dataB64: string): { w: number, h: number } | null {
	const b = decodeBase64ToBytes(dataB64)
	return b ? getImageSizeFromBytes(b) : null
}

/**
 * Read the intrinsic dimensions of an image from raw header bytes — the
 * byte-level core shared by {@link getImageSizeFromBase64} and the read API's
 * `Picture.setImage({ fit })`, which already holds the media bytes.
 * @param {Uint8Array} b - image bytes
 * @returns {{ w: number, h: number } | null} natural size, or `null` when unmeasurable
 */
export function getImageSizeFromBytes (b: Uint8Array): { w: number, h: number } | null {
	if (!b || b.length < 24) return null

	// PNG: 8-byte signature, then IHDR with width@16 / height@20 (big-endian uint32)
	if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
		const w = (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19]
		const h = (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23]
		return w > 0 && h > 0 ? { w, h } : null
	}

	// GIF: "GIF87a"/"GIF89a", width@6 / height@8 (little-endian uint16)
	if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
		const w = b[6] | (b[7] << 8)
		const h = b[8] | (b[9] << 8)
		return w > 0 && h > 0 ? { w, h } : null
	}

	// BMP: "BM", width@18 / height@22 (little-endian int32; height may be negative for top-down)
	if (b[0] === 0x42 && b[1] === 0x4d) {
		const w = b[18] | (b[19] << 8) | (b[20] << 16) | (b[21] << 24)
		const h = b[22] | (b[23] << 8) | (b[24] << 16) | (b[25] << 24)
		const aw = Math.abs(w)
		const ah = Math.abs(h)
		return aw > 0 && ah > 0 ? { w: aw, h: ah } : null
	}

	// WebP: "RIFF"...."WEBP" then a VP8 / VP8L / VP8X chunk
	if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) {
		const fourCC = String.fromCharCode(b[12], b[13], b[14], b[15])
		if (fourCC === 'VP8 ' && b.length >= 30) {
			// Lossy: 14-bit width/height at offset 26/28 (little-endian, mask off scale bits)
			const w = ((b[26] | (b[27] << 8)) & 0x3fff)
			const h = ((b[28] | (b[29] << 8)) & 0x3fff)
			return w > 0 && h > 0 ? { w, h } : null
		}
		if (fourCC === 'VP8L' && b.length >= 25) {
			// Lossless: 14-bit width/height packed starting at bit 0 of offset 21
			const bits = b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24)
			const w = (bits & 0x3fff) + 1
			const h = ((bits >> 14) & 0x3fff) + 1
			return w > 0 && h > 0 ? { w, h } : null
		}
		if (fourCC === 'VP8X' && b.length >= 30) {
			// Extended: 24-bit canvas width/height minus one at offset 24/27 (little-endian)
			const w = (b[24] | (b[25] << 8) | (b[26] << 16)) + 1
			const h = (b[27] | (b[28] << 8) | (b[29] << 16)) + 1
			return w > 0 && h > 0 ? { w, h } : null
		}
		return null
	}

	// JPEG: "FFD8", scan segment markers for a Start-Of-Frame (SOFn) and read height@5 / width@7
	if (b[0] === 0xff && b[1] === 0xd8) {
		let i = 2
		while (i + 9 < b.length) {
			if (b[i] !== 0xff) { i++; continue }
			const marker = b[i + 1]
			// SOF0..SOF15 carry frame dimensions, excluding DHT(C4)/JPG(C8)/DAC(CC)
			if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
				const h = (b[i + 5] << 8) | b[i + 6]
				const w = (b[i + 7] << 8) | b[i + 8]
				return w > 0 && h > 0 ? { w, h } : null
			}
			// Standalone markers (RSTn / SOI / EOI / TEM) have no length payload
			if ((marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) { i += 2; continue }
			// Otherwise skip this segment using its 2-byte big-endian length
			const segLen = (b[i + 2] << 8) | b[i + 3]
			if (segLen < 2) break
			i += 2 + segLen
		}
		return null
	}

	// SVG: text-based vector with no binary signature. When the payload is an
	// `<svg>` document, read its intrinsic size from the root element so that
	// `sizing: 'cover' | 'contain'` is aspect-correct for SVG, not just rasters.
	const text = utf8Decode(b)
	if (/<svg[\s>]/i.test(text)) return getSvgSizeFromMarkup(text)

	return null
}

/**
 * Compute the `<a:srcRect>` crop percentages (each in 1/1000 of a percent, the
 * OOXML unit) for fitting an image of natural size `img` into a display `box`,
 * assuming the cropped region is then stretched to fill the box (`<a:stretch>`).
 *
 * - `cover`: fill the box, cropping the overflowing axis (positive l/r or t/b)
 * - `contain`: fit inside the box, letterboxing the short axis (negative l/r or t/b)
 *
 * Single source of truth for the crop math shared by the write side
 * (`ImageSizingXml`) and the read API's `Picture.setImage({ fit })`. `l`/`r` and
 * `t`/`b` are symmetric (centered crop).
 * @param {'cover' | 'contain'} type - fit mode
 * @param {{ w: number, h: number }} img - natural image pixel size
 * @param {{ w: number, h: number }} box - displayed frame size (any consistent unit)
 * @returns {{ l: number, r: number, t: number, b: number }} srcRect percentages
 */
export function fitSrcRectPercents (
	type: 'cover' | 'contain',
	img: { w: number, h: number },
	box: { w: number, h: number },
): { l: number, r: number, t: number, b: number } {
	const imgRatio = img.h / img.w
	const boxRatio = box.h / box.w
	let width: number
	let height: number
	if (type === 'cover') {
		const isBoxBased = boxRatio > imgRatio
		width = isBoxBased ? box.h / imgRatio : box.w
		height = isBoxBased ? box.h : box.w * imgRatio
	} else {
		const widthBased = boxRatio > imgRatio
		width = widthBased ? box.w : box.h / imgRatio
		height = widthBased ? box.w * imgRatio : box.h
	}
	const hz = Math.round(1e5 * 0.5 * (1 - box.w / width))
	const vz = Math.round(1e5 * 0.5 * (1 - box.h / height))
	return { l: hz, r: hz, t: vz, b: vz }
}

/**
 * Read the intrinsic size of an SVG document from its root `<svg>` element.
 * Follows the SVG sizing model: an explicit absolute `width`/`height` pair wins;
 * otherwise the `viewBox` width/height defines the size (and thus aspect ratio).
 * Percentage or missing `width`/`height` fall through to `viewBox`.
 * @param {string} svg - SVG markup
 * @returns {{ w: number, h: number } | null} intrinsic size, or `null` when undeterminable
 */
function getSvgSizeFromMarkup (svg: string): { w: number, h: number } | null {
	const openTag = /<svg\b[^>]*>/i.exec(svg)?.[0]
	if (!openTag) return null
	const attr = (name: string): string | null => new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`, 'i').exec(openTag)?.[1] ?? null
	// Leading number with an optional absolute unit; a percentage is not an intrinsic length.
	const absLength = (val: string | null): number => {
		if (val == null || /%\s*$/.test(val)) return NaN
		const m = /^\s*\+?(\d*\.?\d+)/.exec(val)
		return m ? parseFloat(m[1]) : NaN
	}
	let w = absLength(attr('width'))
	let h = absLength(attr('height'))
	if (!(w > 0 && h > 0)) {
		const vb = attr('viewBox')
		const p = vb ? vb.trim().split(/[\s,]+/).map(Number) : []
		if (p.length === 4 && p[2] > 0 && p[3] > 0) { w = p[2]; h = p[3] }
	}
	return w > 0 && h > 0 ? { w, h } : null
}

/** Decode UTF-8 bytes to a string, isomorphic across Node and browsers. */
function utf8Decode (bytes: Uint8Array): string {
	return new TextDecoder().decode(bytes)
}
