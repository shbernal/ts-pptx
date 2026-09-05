/**
 * ts-pptx: hex colour text, in the one spelling both halves use.
 *
 * A caller may write a colour with or without a leading `#`, and the library accepts both
 * everywhere — which means every site that validates or parses one strips it first. Those
 * strips had drifted into three spellings across six sites, and one of them was not
 * anchored: `.replace('#', '')` takes the `#` out of `'FF00#0'` and hands `'FF000'` on, which
 * then fails the six-digit test and takes a different branch than the anchored forms would.
 * No caller is known to reach that, which is exactly why it wants one spelling rather than a
 * test at each site.
 *
 * This lives at the root rather than under `gen/` or `read/` because both halves need it:
 * the emitters strip on the way in, the readers strip on the way back out — and for the same
 * reason the *validation* belongs here too. `REGEX_HEX_COLOR` used to live in
 * `constants-internal.ts`, which is a write-side module the read path cannot import, so the read
 * path re-spelled the regex; so did one write-side site. Six sites, three spellings of one fact.
 */

/**
 * A colour written as six hex digits, `#` already stripped.
 *
 * Not exported: {@link isHexColor} is the whole of what the call sites want, and a shared mutable
 * `RegExp` object is a worse export than a predicate.
 */
const REGEX_HEX_COLOR = /^[0-9a-fA-F]{6}$/

/** An 8-hex RGBA colour, `#` already stripped — {@link splitRgbaHex}'s test. */
const REGEX_ARGB_HEX = /^[0-9a-fA-F]{8}$/

/** Strip a single leading `#` from a colour, leaving any other character alone. */
export function stripHash(value: string): string {
	return value.startsWith('#') ? value.slice(1) : value
}

/**
 * Is `value` a 6-hex RGB colour? A leading `#` is stripped first, so both spellings agree.
 *
 * This is the *shape* test only. Whether the caller may also name a scheme colour there is the
 * call site's question, and each still asks it: `createColorElement` falls back to the default
 * font colour, `makeXmlTheme` keeps the Office default and warns, `normalizeHex` throws.
 * @param value - the caller's colour, with or without a leading `#`
 */
export function isHexColor(value: string): boolean {
	return REGEX_HEX_COLOR.test(stripHash(value))
}

/**
 * Is `value` an 8-hex RGBA colour? A leading `#` is stripped first.
 *
 * The 8-digit form is RGB followed by one alpha byte; {@link splitRgbaHex} is what takes it apart.
 * @param value - the caller's colour, with or without a leading `#`
 */
function isArgbHex(value: string): boolean {
	return REGEX_ARGB_HEX.test(stripHash(value))
}

/**
 * Split an 8-hex RGBA colour into its 6-hex RGB and a 0–1 alpha; anything else passes through
 * unchanged with no alpha.
 *
 * `#` is stripped first, so both spellings reach the same answer. Two sites parsed this by
 * hand — `createColorElement`, which turns the alpha into an `<a:alpha>` sibling, and
 * `normalizeShadowOptions`, which turns it into a shadow's `_alpha` — and they divided by 255 at
 * different points in their arithmetic.
 * @param value - the caller's colour, with or without a leading `#`
 */
export function splitRgbaHex(value: string): { rgb: string; alpha?: number } {
	const hex = stripHash(value)
	if (!isArgbHex(hex)) return { rgb: hex }
	return { rgb: hex.slice(0, 6), alpha: parseInt(hex.slice(6, 8), 16) / 255 }
}
