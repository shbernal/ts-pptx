/**
 * ts-pptx: DrawingML color elements
 *
 * Emit the bare `<a:srgbClr>` / `<a:schemeClr>` color element that every other
 * color context builds on (solid fills, gradient stops, effect alphas, line
 * fills, highlights, patterns).
 *
 * The RGB->hex helpers this module used to carry now live in
 * `gen/table/html-dom.ts`, their only caller: they are reachable exclusively from
 * the browser-only `tableToSlides` path, so keeping them here made them permanent
 * dead weight in the Node chunk's coverage report (33% functions on a file that is
 * otherwise fully exercised) while html-dom.ts is coverage-excluded wholesale.
 */

import { SchemeColor, type SCHEME_COLORS } from '../../enums.js'
import { DEF_FONT_COLOR } from '../../constants-internal.js'
import { isHexColor, splitRgbaHex, stripHash } from '../../hex-color.js'
import { warn } from '../../diagnostics.js'
import { PERCENT_SCALE } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'

/**
 * Opening delimiter of an `<a:alpha>` child, for detecting one a caller already
 * supplied. A prefix rather than a whole element: the value is the caller's, and
 * this asks only whether the element is present.
 */
const ALPHA_TAG = '<a:alpha'

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
 * @param {string|SCHEME_COLORS} colorStr - hex RGB (e.g. "FFFF00") or a scheme color constant (e.g. SchemeColor.ACCENT1)
 * @param {string} [innerElements] - additional elements that adjust the color, nested inside the color element
 * @returns {string} XML string
 */
export function createColorElement(colorStr: string | SCHEME_COLORS, innerElements?: string): string {
	if (typeof colorStr !== 'string') {
		warn(
			'color/not-a-string',
			`createColorElement: expected a string color value, got ${typeof colorStr}. "${DEF_FONT_COLOR}" used instead.`
		)
		colorStr = DEF_FONT_COLOR
	}
	let colorVal = stripHash(colorStr || '')

	// 8-char hex (RGBA) — strip the alpha byte to a sibling <a:alpha val="N"/>,
	// continue with the leading 6-char RGB through the existing validation. This keeps
	// fill/text/line/glow paths from silently falling back to DEF_FONT_COLOR on RGBA input.
	const rgba = splitRgbaHex(colorVal)
	if (rgba.alpha !== undefined) {
		// If the caller already supplied an explicit <a:alpha> (e.g. shadow/glow `opacity`),
		// it wins — do NOT add a second alpha from the RGBA byte, which would emit two
		// <a:alpha> children and produce schema-invalid OOXML (CT_SRgbColor allows one).
		if (!innerElements?.includes(ALPHA_TAG))
			innerElements = alphaEl(Math.round(rgba.alpha * PERCENT_SCALE)) + (innerElements || '')
		colorVal = rgba.rgb
	}

	if (!isHexColor(colorVal) && !Object.values(SchemeColor).includes(colorVal as SchemeColor)) {
		warn(
			'color/invalid-value',
			`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'SchemeColor' values!`
		)
		colorVal = DEF_FONT_COLOR
	}

	const isHex = isHexColor(colorVal)
	const name = isHex ? 'a:srgbClr' : 'a:schemeClr'
	const attrs = { val: isHex ? colorVal.toUpperCase() : colorVal }

	// Paired vs self-closing is decided by whether there is anything to nest, so this
	// is one of the few places `el`/`voidEl` are chosen at runtime rather than by tag.
	return innerElements ? el(name, attrs, raw(innerElements)) : voidEl(name, attrs)
}

/**
 * The `<a:alpha val>` child a colour element carries.
 *
 * Four sites built it: the RGBA split above, the gradient stop's `transparency`, the glow and
 * shadow `opacity`, and the chart's `chartColorsOpacity`. Each converts to the fixed
 * percentage its own option speaks first; this is only the element.
 * @param fixedPct - the alpha in `ST_PositiveFixedPercentage` units (`100%` is `100000`)
 */
export function alphaEl(fixedPct: number): string {
	return voidEl('a:alpha', { val: fixedPct })
}
