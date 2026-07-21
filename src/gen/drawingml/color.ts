/**
 * PptxGenJS: DrawingML color elements
 *
 * Emit the bare `<a:srgbClr>` / `<a:schemeClr>` color element that every other
 * color context builds on (solid fills, gradient stops, effect alphas, line
 * fills, highlights, patterns), plus the RGB->hex helpers the HTML table
 * importer needs to get CSS colors into that form.
 */

import { SchemeColor, type SCHEME_COLORS } from '../../core-enums.js'
import { REGEX_HEX_COLOR, DEF_FONT_COLOR } from '../../core-enums-internal.js'
import { warn } from '../../log.js'
import { PERCENT_SCALE } from '../../units.js'

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
 * folded into `genXmlColorSelection`.
 *
 * When you specifically want a *solid fill*, prefer `genXmlColorSelection` (`fill.ts`), which
 * wraps this in `<a:solidFill>…</a:solidFill>` and also handles `alpha`/`transparency` and the
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
