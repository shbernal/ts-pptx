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
import { warn, warnOnce } from '../../diagnostics.js'
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
	// An empty string reaches here only from a slot that *requires* a colour — a gradient
	// stop, a duotone half, a `buClr`. There is no "inherit" state to fall back to, so this
	// is the one place `''` still paints: it is reported under its own code, with a message
	// that names the real problem rather than `"" is not a valid scheme color`.
	if (colorStr === '') {
		warnOnce(
			'color/empty-string',
			`An empty string is not a color. "${DEF_FONT_COLOR}" used instead: this position requires one, so there is nothing to inherit. Omit the whole option if you meant to paint nothing.`
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
 * Reject an empty colour string, reporting it once per option that named one.
 *
 * `''` is not a spelling of "no paint": omission already spells that, and `'inherit'` and
 * `'none'` spell the other two states a paint can be in. It reaches the library from the
 * caller's own missing value — an unset template field, `row.accent` on a row that has none
 * — so the useful answer is to say so and then behave exactly as omitting the option would
 * have. What it must not do is coerce to {@link DEF_FONT_COLOR}, which paints visible black
 * on a shape the caller expected to keep the theme's paint.
 *
 * Every site that consumes a caller-supplied colour calls this *before* its own fallback, so
 * the diagnostic names the option that carried the empty string rather than whichever
 * emitter happened to see it last. {@link createColorElement} is the exception and reports
 * the same code itself: its slot has no absent state to fall back to.
 * @param value - a colour option's value, whatever shape the option accepts
 * @param option - the option's name, for the message
 * @returns true when `value` was the empty string
 */
export function rejectEmptyColor(value: unknown, option: string): boolean {
	if (value !== '') return false
	warnOnce(
		'color/empty-string',
		`\`${option}\` is an empty string, which is not a color — the option was ignored, exactly as omitting it would be. Pass 6-digit hex RGB (e.g. "FF3399") or a SchemeColor value.`
	)
	return true
}

/**
 * The colour the caller named for `option`, or `fallback` when they named none.
 *
 * The `x ?? DEFAULT` and `x || DEFAULT` this replaces disagreed about `''`: the first passed
 * it through to be painted black or dropped, the second quietly resolved it to the default.
 * Here it is one rule — an empty string is a missing value, reported and then resolved the
 * way an absent option is.
 * @param value - the caller's colour option, possibly absent or empty
 * @param fallback - what an absent option resolves to at this site
 * @param option - the option's name, for the message
 */
export function namedColorOr<T>(value: string | undefined, fallback: T, option: string): string | T {
	if (value === undefined || rejectEmptyColor(value, option)) return fallback
	return value
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
