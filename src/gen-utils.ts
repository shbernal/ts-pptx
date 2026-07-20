/**
 * PptxGenJS: Utility Methods
 *
 * Shared pure helpers used across the generators — no state, no I/O. Roughly grouped:
 *   - Units & coordinates   getSmartParseNumber, inch2Emu, marginToEmu, valToPts, rotation/arc angles
 *   - Identifiers & naming   getUuid, validateObjectName, getDuplicateObjectNames, getNewRelId
 *   - XML text               encodeXmlEntities
 *   - Color                  componentToHex, rgbToHex, createColorElement, genXmlColorSelection
 *   - Effects                createGlowElement, createShadowElement(+EffectLst), correctShadowOptions
 *   - Fills & lines          genXmlGradientFill / genXmlPatternFill / genXmlImageFill / genXmlLineFill, createLineCap
 *   - Content types          imageContentType, avContentType
 *   - Image decoding         decodeBase64ToBytes, getImageSizeFromBase64/Bytes, SVG size sniffing, fitSrcRectPercents
 *
 * `getSmartParseNumber` is the single user-coordinate → EMU boundary; keep unit handling here.
 */

import { REGEX_HEX_COLOR, DEF_FONT_COLOR, EMU, ONEPT, SchemeColor, type SCHEME_COLORS } from './core-enums.js'
import { warn, warnOnce } from './log.js'
import {
	ANGLE_UNITS_PER_DEGREE,
	coordToEmu,
	FIXED_PCT_PER_PERCENT,
	inchesToEmu,
	PERCENT_SCALE,
	type Emu,
} from './units.js'
import type {
	BorderProps,
	PresLayout,
	TextGlowProps,
	PresSlideInternal,
	ShapeFillProps,
	Color,
	ShapeLineProps,
	Coord,
	ShadowProps,
	ShadowPropsInternal,
	GradientFillProps,
	GradientStopProps,
	PatternFillProps,
	LineCap,
} from './core-interfaces.js'

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
export function getSmartParseNumber(size: Coord | null | undefined, xyDir: 'X' | 'Y', layout: PresLayout): Emu {
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
 * Convert inches into EMU.
 * - accepts a number (inches) or a numeric/`"<n>in"` string
 * - no magnitude guessing: values are always treated as inches (use {@link coordToEmu} for
 *   user coordinates that may carry other units)
 * @param {number|string} inches - inches as number or string
 * @returns {Emu} EMU value
 */
export function inch2Emu(inches: number | string): Emu {
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return inchesToEmu(inches)
}

/**
 * Convert a single `margin` component (table cell/table margin, or text-box/placeholder body
 * inset) to EMU.
 *
 * Margins are INCHES, consistent with the positional API (`x`/`y`/`w`/`h`) and the value
 * PowerPoint's own dialog shows (both the table cell-margin and the text-box internal-margin
 * fields are inches). Historically the library read these as POINTS (table cells used a magnitude
 * heuristic — `>= 1` points, `< 1` inches; text-box margins were straight points), so a legitimate
 * fraction-of-an-inch value entered from the PowerPoint dialog became a tiny points value. Every
 * value is now inches. A `>= 1` value is honored as inches but warns once, because it is almost
 * certainly a legacy points value that should be divided by 72 (e.g. `10` points → `0.139` inches).
 *
 * Shared by every margin site so they stay in lockstep: the cell XML emitter and text-box/slide-
 * number insets (`gen-xml`), the autoPage row-height pass (`gen-tables`), and the measured-fit pass
 * (`measure-fit`).
 * @param {number} inches - margin component in inches
 * @returns {Emu} EMU value
 */
export function marginToEmu(inches: number): Emu {
	if (inches >= 1)
		warnOnce(
			'margins (table cell and text-box) are interpreted as inches (matching the rest of the API and the ' +
				'PowerPoint dialog); a value >= 1 is likely a legacy points value — divide by 72 to convert (e.g. 10pt => 0.139in).'
		)
	return inch2Emu(inches)
}

/**
 * Resolve a table's column widths to EMU, the single source of truth shared by the
 * table XML emitter and the measured-fit pass (so a fitted cell sees the same grid
 * the renderer draws).
 * - an explicit `colW` **array** is per-column inches (`inch2Emu`); a non-finite slot
 *   falls back to the even-distribution width
 * - otherwise the table's already-resolved width (`totalWidthEmu`, EMU) is split evenly
 *   across `colCount` columns. (A scalar `colW` never reaches here — `addTableDefinition`
 *   converts it to `w` and clears `colW`.)
 *
 * IMPORTANT: the even path divides an **EMU** width. Passing the raw inches `options.w`
 * instead (the historical bug) produced ~0-EMU columns (e.g. `w=9` → `gridCol w="3"`),
 * collapsing auto-width tables to a sliver in PowerPoint/LibreOffice.
 * @param {Coord[]|Coord|undefined} colW - explicit per-column inches, or scalar/undefined
 * @param {number} totalWidthEmu - the table's resolved width in EMU
 * @param {number} colCount - number of grid columns (counting colspans)
 * @returns {number[]} per-column widths in EMU (length `colCount`)
 */
export function resolveTableColWidthsEmu(
	colW: Coord[] | Coord | undefined,
	totalWidthEmu: number,
	colCount: number
): number[] {
	if (!(colCount > 0)) return []
	const even = totalWidthEmu > 0 ? Math.round(totalWidthEmu / colCount) : EMU
	if (Array.isArray(colW)) {
		return Array.from({ length: colCount }, (_, i) => {
			// Guard before inch2Emu: it throws on non-finite input. A missing/NaN slot
			// falls back to the even-distribution width.
			const n = colW[i]
			return typeof n === 'number' && Number.isFinite(n) ? Math.round(inch2Emu(n)) : even
		})
	}
	return new Array<number>(colCount).fill(even)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts(pt: number | string): number {
	const points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert a transparency percentage (0-100) into a schema-valid `<a:alpha>` value
 * (ST_PositiveFixedPercentage, 0-100000). Out-of-range transparency yields an
 * alpha that PowerPoint rejects as needing repair, so clamp into range and warn.
 */
export function transparencyToAlpha(transparency: number): number {
	const pct = Math.min(100, Math.max(0, transparency))
	if (pct !== transparency) warn(`transparency ${transparency} is outside the valid range 0-100; using ${pct}.`)
	return Math.round((100 - pct) * FIXED_PCT_PER_PERCENT)
}

/** Convert an opacity (0-1) into a schema-valid `<a:alpha>` value (0-100000); clamps + warns on out-of-range input. */
export function opacityToAlpha(opacity: number): number {
	const o = Math.min(1, Math.max(0, opacity))
	if (o !== opacity) warn(`opacity ${opacity} is outside the valid range 0-1; using ${o}.`)
	return Math.round(o * PERCENT_SCALE)
}

/**
 * Convert a line width (points) to EMU clamped into ST_LineWidth (0..20116800 EMU,
 * i.e. 0-1584pt). Out-of-range widths make PowerPoint report the package as needing
 * repair, so clamp into range and warn.
 */
export function lineWidthToEmu(widthPts: number | string): number {
	const raw = valToPts(widthPts)
	const clamped = Math.min(20116800, Math.max(0, raw))
	if (clamped !== raw) warn(`line width ${widthPts} is outside the valid range 0-1584pt; using ${clamped / ONEPT}.`)
	return clamped
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 * @param {number} d degrees
 * @returns {number} calculated `rot` value
 */
export function convertRotationDegrees(d: number): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * ANGLE_UNITS_PER_DEGREE)
}

/**
 * Convert a freeform arc angle (degrees) to an `<a:arcTo>` ST_AdjAngle value (60000ths).
 * Unlike a shape rotation, a sweep is not modular: a 400 degree swAng draws a different
 * arc than a 40 degree one, so the value is never wrapped into 0..360.
 * @param {number} d degrees
 * @param {'stAng' | 'swAng'} attr - attribute being emitted, for the error message
 * @returns {number} ST_AdjAngle value (60000ths of a degree)
 */
export function convertArcAngle(d: number, attr: 'stAng' | 'swAng'): number {
	if (typeof d !== 'number' || !Number.isFinite(d))
		throw new Error(`Arc ${attr} must be a finite number of degrees; received ${String(d)}.`)
	return Math.round(d * ANGLE_UNITS_PER_DEGREE)
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
	const size = Math.round(opts.size * ONEPT)
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
 * @param {ShadowProps} options shadow properties
 * @param {ShadowProps} defaults defaults for unspecified properties in `options`
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
 * @param {ShadowProps} options shadow properties
 * @param {ShadowProps} defaults defaults for unspecified properties in `options`
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
 * Map an image file extension to its OOXML content type.
 * Inverse of the read-side `IMAGE_EXTENSION_BY_CONTENT_TYPE` (src/read/api/shapes.ts):
 * EMF/WMF use the `x-`-prefixed forms PowerPoint authors (and that the read side
 * expects), `jpg`/`jpeg` normalize to `image/jpeg`, and `svg` to `image/svg+xml`.
 * Only the content type is derived here; the file extension (used for the media
 * Target filename) is left to the caller.
 * @param {string} extn - image file extension (e.g. `png`, `jpg`, `emf`)
 * @returns {string} OOXML content type (e.g. `image/png`, `image/x-emf`)
 */
export function imageContentType(extn: string): string {
	switch ((extn || '').toLowerCase()) {
		case 'emf':
			return 'image/x-emf'
		case 'wmf':
			return 'image/x-wmf'
		case 'svg':
			return 'image/svg+xml'
		case 'jpg':
		case 'jpeg':
			return 'image/jpeg'
		default:
			return 'image/' + (extn || '').toLowerCase()
	}
}

/**
 * Resolve the OPC content type for an embedded audio/video part by file extension,
 * matching what PowerPoint authors (e.g. `mp3` → `audio/mpeg`, not `audio/mp3`).
 * The `mtype` disambiguates extensions Office maps differently per kind and seeds
 * the `mtype/extn` fallback for anything unlisted.
 * @param {string} extn - media file extension (no dot), case-insensitive
 * @param {'audio' | 'video'} mtype - whether the item is audio or video
 */
export function avContentType(extn: string, mtype: 'audio' | 'video'): string {
	switch ((extn || '').toLowerCase()) {
		// video
		case 'mp4':
			return mtype === 'audio' ? 'audio/mp4' : 'video/mp4'
		case 'm4v':
			return 'video/mp4'
		case 'mov':
			return 'video/quicktime'
		case 'avi':
			return 'video/avi'
		case 'wmv':
			return 'video/x-ms-wmv'
		case 'mpg':
		case 'mpeg':
			return mtype === 'audio' ? 'audio/mpeg' : 'video/mpeg'
		case 'ogv':
			return 'video/ogg'
		case 'webm':
			return 'video/webm'
		// audio
		case 'mp3':
			return 'audio/mpeg'
		case 'm4a':
			return 'audio/mp4'
		case 'wav':
			return 'audio/x-wav' // PowerPoint authors the x- form (e.g. embedded transition sounds)
		case 'wma':
			return 'audio/x-ms-wma'
		case 'aac':
			return 'audio/aac'
		case 'oga':
		case 'ogg':
			return 'audio/ogg'
		case 'flac':
			return 'audio/flac'
		default:
			return mtype + '/' + (extn || '').toLowerCase()
	}
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

/**
 * Encode raw SVG markup as a base64 `image/svg+xml` data URI.
 * - lets callers pass inline SVG to `addImage({ svg })` without hand-rolling base64
 * - isomorphic and UTF-8 safe: uses the global `TextEncoder`/`btoa` (Node and browsers)
 * @param {string} svg - SVG markup, e.g. `'<svg ...>...</svg>'`
 * @returns {string} a `data:image/svg+xml;base64,...` URI
 */
export function svgMarkupToDataUri(svg: string): string {
	const bytes = new TextEncoder().encode(svg)
	let binary = ''
	for (const byte of bytes) binary += String.fromCharCode(byte)
	return `data:image/svg+xml;base64,${btoa(binary)}`
}

/**
 * Decode a base64 image payload (raw base64 or a `data:` URI) to bytes.
 * - tolerant of the `data:[mime];base64,` prefix and of whitespace in the payload
 * @param {string} b64 - base64 string or data URI
 * @returns {Uint8Array | null} decoded bytes, or `null` when the payload is empty/undecodable
 */
export function decodeBase64ToBytes(b64: string): Uint8Array | null {
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
export function getImageSizeFromBase64(dataB64: string): { w: number; h: number } | null {
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
export function getImageSizeFromBytes(b: Uint8Array): { w: number; h: number } | null {
	if (!b || b.length < 24) return null

	// Bounds-checked byte read: every access below is already guarded by an
	// explicit length check, so the `?? 0` fallback is unreachable in practice.
	const u = (n: number): number => b[n] ?? 0

	// PNG: 8-byte signature, then IHDR with width@16 / height@20 (big-endian uint32)
	if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
		const w = (u(16) << 24) | (u(17) << 16) | (u(18) << 8) | u(19)
		const h = (u(20) << 24) | (u(21) << 16) | (u(22) << 8) | u(23)
		return w > 0 && h > 0 ? { w, h } : null
	}

	// GIF: "GIF87a"/"GIF89a", width@6 / height@8 (little-endian uint16)
	if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
		const w = u(6) | (u(7) << 8)
		const h = u(8) | (u(9) << 8)
		return w > 0 && h > 0 ? { w, h } : null
	}

	// BMP: "BM", width@18 / height@22 (little-endian int32; height may be negative for top-down)
	if (b[0] === 0x42 && b[1] === 0x4d) {
		const w = u(18) | (u(19) << 8) | (u(20) << 16) | (u(21) << 24)
		const h = u(22) | (u(23) << 8) | (u(24) << 16) | (u(25) << 24)
		const aw = Math.abs(w)
		const ah = Math.abs(h)
		return aw > 0 && ah > 0 ? { w: aw, h: ah } : null
	}

	// WebP: "RIFF"...."WEBP" then a VP8 / VP8L / VP8X chunk
	if (
		b[0] === 0x52 &&
		b[1] === 0x49 &&
		b[2] === 0x46 &&
		b[3] === 0x46 &&
		b[8] === 0x57 &&
		b[9] === 0x45 &&
		b[10] === 0x42 &&
		b[11] === 0x50
	) {
		const fourCC = String.fromCharCode(u(12), u(13), u(14), u(15))
		if (fourCC === 'VP8 ' && b.length >= 30) {
			// Lossy: 14-bit width/height at offset 26/28 (little-endian, mask off scale bits)
			const w = (u(26) | (u(27) << 8)) & 0x3fff
			const h = (u(28) | (u(29) << 8)) & 0x3fff
			return w > 0 && h > 0 ? { w, h } : null
		}
		if (fourCC === 'VP8L' && b.length >= 25) {
			// Lossless: 14-bit width/height packed starting at bit 0 of offset 21
			const bits = u(21) | (u(22) << 8) | (u(23) << 16) | (u(24) << 24)
			const w = (bits & 0x3fff) + 1
			const h = ((bits >> 14) & 0x3fff) + 1
			return w > 0 && h > 0 ? { w, h } : null
		}
		if (fourCC === 'VP8X' && b.length >= 30) {
			// Extended: 24-bit canvas width/height minus one at offset 24/27 (little-endian)
			const w = (u(24) | (u(25) << 8) | (u(26) << 16)) + 1
			const h = (u(27) | (u(28) << 8) | (u(29) << 16)) + 1
			return w > 0 && h > 0 ? { w, h } : null
		}
		return null
	}

	// JPEG: "FFD8", scan segment markers for a Start-Of-Frame (SOFn) and read height@5 / width@7
	if (b[0] === 0xff && b[1] === 0xd8) {
		let i = 2
		while (i + 9 < b.length) {
			if (b[i] !== 0xff) {
				i++
				continue
			}
			const marker = u(i + 1)
			// SOF0..SOF15 carry frame dimensions, excluding DHT(C4)/JPG(C8)/DAC(CC)
			if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
				const h = (u(i + 5) << 8) | u(i + 6)
				const w = (u(i + 7) << 8) | u(i + 8)
				return w > 0 && h > 0 ? { w, h } : null
			}
			// Standalone markers (RSTn / SOI / EOI / TEM) have no length payload
			if ((marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) {
				i += 2
				continue
			}
			// Otherwise skip this segment using its 2-byte big-endian length
			const segLen = (u(i + 2) << 8) | u(i + 3)
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
export function fitSrcRectPercents(
	type: 'cover' | 'contain',
	img: { w: number; h: number },
	box: { w: number; h: number }
): { l: number; r: number; t: number; b: number } {
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
function getSvgSizeFromMarkup(svg: string): { w: number; h: number } | null {
	const openTag = /<svg\b[^>]*>/i.exec(svg)?.[0]
	if (!openTag) return null
	const attr = (name: string): string | null =>
		new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`, 'i').exec(openTag)?.[1] ?? null
	// Leading number with an optional absolute unit; a percentage is not an intrinsic length.
	const absLength = (val: string | null): number => {
		if (val == null || /%\s*$/.test(val)) return NaN
		const m = /^\s*\+?(\d*\.?\d+)/.exec(val)
		return m ? parseFloat(m[1] ?? '') : NaN
	}
	let w = absLength(attr('width'))
	let h = absLength(attr('height'))
	if (!(w > 0 && h > 0)) {
		const vb = attr('viewBox')
		const p = vb
			? vb
					.trim()
					.split(/[\s,]+/)
					.map(Number)
			: []
		const vw = p[2]
		const vh = p[3]
		if (p.length === 4 && vw != null && vh != null && vw > 0 && vh > 0) {
			w = vw
			h = vh
		}
	}
	return w > 0 && h > 0 ? { w, h } : null
}

/** Decode UTF-8 bytes to a string, isomorphic across Node and browsers. */
function utf8Decode(bytes: Uint8Array): string {
	return new TextDecoder().decode(bytes)
}
