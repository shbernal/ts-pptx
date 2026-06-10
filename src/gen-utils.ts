/**
 * PptxGenJS: Utility Methods
 */

import { EMU, REGEX_HEX_COLOR, DEF_FONT_COLOR, ONEPT, SchemeColor, SCHEME_COLORS } from './core-enums.js'
import type { PresLayout, TextGlowProps, PresSlideInternal, ShapeFillProps, Color, ShapeLineProps, Coord, ShadowProps, GradientFillProps, GradientStopProps } from './core-interfaces.js'

/**
 * Translates any type of `x`/`y`/`w`/`h` prop to EMU
 * - guaranteed to return a result regardless of undefined, null, etc. (0)
 * - {number} - 12800 (EMU)
 * - {number} - 0.5 (inches)
 * - {string} - "75%"
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
export function getSmartParseNumber (size: Coord | null | undefined, xyDir: 'X' | 'Y', layout: PresLayout): number {
	// FIRST: Convert string numeric value if reqd
	if (typeof size === 'string' && !isNaN(Number(size))) size = Number(size)

	// GUARD: A NaN/Infinity coordinate is always a mistake (commonly arithmetic on an
	// `undefined` layout dimension). Fail loud instead of silently emitting 0 EMU, which
	// collapses the object to zero size/position and produces a broken-looking deck.
	if (typeof size === 'number' && !isFinite(size)) {
		throw new Error(
			`Invalid ${xyDir || 'coordinate'} value: expected a finite number but received ${String(size)}. ` +
				'This usually means a layout dimension was read from a missing property (e.g. `layout.width` returning `undefined`). ' +
				'Use `slide.width`/`slide.height` or `STANDARD_LAYOUTS.<NAME>.width`/`.height` (inches).'
		)
	}

	// CASE 1: Number in inches
	// Assume any number less than 100 is inches
	if (typeof size === 'number' && size < 100) return inch2Emu(size)

	// CASE 2: Number is already converted to something other than inches
	// Assume any number greater than 100 sure isnt inches! Just return it (assume value is EMU already).
	if (typeof size === 'number' && size >= 100) return size

	// CASE 3: Percentage (ex: '50%')
	if (typeof size === 'string' && size.includes('%')) {
		if (xyDir && xyDir === 'X') return Math.round((parseFloat(size) / 100) * layout.width)
		if (xyDir && xyDir === 'Y') return Math.round((parseFloat(size) / 100) * layout.height)

		// Default: Assume width (x/cx)
		return Math.round((parseFloat(size) / 100) * layout.width)
	}

	// LAST: Default value
	return 0
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
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export function inch2Emu (inches: number | string): number {
	// NOTE: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any value over 100 damn sure isnt inches, so lets assume its in EMU already, therefore, just return the same value
	if (typeof inches === 'number' && inches > 100) return inches
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return Math.round(EMU * inches)
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
		const alphaHex = colorVal.slice(6, 8)
		const alphaVal = Math.round((parseInt(alphaHex, 16) / 255) * 100000)
		innerElements = `<a:alpha val="${alphaVal}"/>${innerElements || ''}`
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
	const opacity = Math.round((opts.opacity ?? 0) * 100000)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += '</a:glow>'

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
	if (stop.alpha) internalElements += `<a:alpha val="${Math.round((100 - stop.alpha) * 1000)}"/>` // DEPRECATED: @deprecated v3.3.0
	if (stop.transparency) internalElements += `<a:alpha val="${Math.round((100 - stop.transparency) * 1000)}"/>`
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
	if (!gradient || gradient.kind !== 'linear') throw new Error('Gradient fill currently supports only linear gradients.')
	if (typeof gradient.rotateWithShape !== 'undefined' && typeof gradient.rotateWithShape !== 'boolean') {
		throw new Error('Gradient rotateWithShape must be a boolean.')
	}
	if (typeof gradient.scaled !== 'undefined' && typeof gradient.scaled !== 'boolean') throw new Error('Gradient scaled must be a boolean.')

	const stops = normalizeGradientStops(gradient.stops)
	const rotWithShape = gradient.rotateWithShape ?? true
	const scaledAttr = typeof gradient.scaled === 'boolean' ? ` scaled="${boolToXml(gradient.scaled)}"` : ''

	let strXml = `<a:gradFill rotWithShape="${boolToXml(rotWithShape)}">`
	strXml += '<a:gsLst>'
	stops.forEach(stop => {
		const position = Math.round(stop.position * 1000)
		strXml += `<a:gs pos="${position}">${createColorElement(stop.color, gradientStopColorAdjustments(stop))}</a:gs>`
	})
	strXml += '</a:gsLst>'
	strXml += `<a:lin ang="${normalizeGradientAngle(gradient.angle)}"${scaledAttr}/>`
	strXml += '</a:gradFill>'

	return strXml
}

/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
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
			if (props.alpha) internalElements += `<a:alpha val="${Math.round((100 - props.alpha) * 1000)}"/>` // DEPRECATED: @deprecated v3.3.0
			if (props.transparency) internalElements += `<a:alpha val="${Math.round((100 - props.transparency) * 1000)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				break
			case 'gradient':
				outText += genXmlGradientFill(typeof props === 'string' ? undefined : props.gradient)
				break
			default: // @note need a statement as having only "break" can be removed by bundlers, then triggers "no-default" js-linter
				outText += ''
				break
		}
	}

	return outText
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
function decodeBase64ToBytes (b64: string): Uint8Array | null {
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
 * Read the intrinsic pixel dimensions of a raster image from its header bytes.
 * - synchronous: parses only file-format headers, never decodes pixels
 * - supports PNG, JPEG, GIF, BMP, and WebP (VP8 / VP8L / VP8X)
 * - vector (SVG) and unrecognized formats return `null` (no intrinsic pixel size)
 *
 * Used by image `sizing: 'cover' | 'contain'` to compute an aspect-correct
 * `<a:srcRect>` crop from the *natural* image ratio rather than the displayed box.
 * @param {string} dataB64 - base64 image payload or `data:` URI
 * @returns {{ w: number, h: number } | null} natural pixel size, or `null` when unmeasurable
 */
export function getImageSizeFromBase64 (dataB64: string): { w: number, h: number } | null {
	const b = decodeBase64ToBytes(dataB64)
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

	return null
}
