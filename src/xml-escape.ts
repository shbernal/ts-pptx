/**
 * ts-pptx: turning a caller's string into something that can sit inside XML.
 *
 * Three escapers with three notions of what is dangerous used to live here-and-there — one under
 * `gen/`, one under `read/oxml/`, and a fourth-of-a-job local one in `embedded-fonts.ts` — and the
 * weakest was the one reached by the least-checked input. `serializeEmbeddedFontLst` puts
 * `font.typeface` straight from `pptx.embedFont({ typeface })` through an escaper that handled
 * `& < > "` and nothing else, so `embedFont({ typeface: 'Foo\v' })` wrote a vertical tab into
 * `<p:font typeface="…">` — a character XML 1.0 forbids outright, and exactly what the write side
 * strips at every other emission site — and `'Foo\nBar'` lost its newline to attribute-value
 * normalisation.
 *
 * The copies had a real reason: `embedded-fonts.ts` is imported by `read/api/ops/`, so it must not
 * pull in `src/gen/`, and `read/oxml/dom.ts` is on the same side of that line. The answer to
 * "both halves need this" is a dependency-free root module, the same placement and rationale as
 * `hex-color.ts` — not a third, weaker copy.
 *
 * This module imports nothing.
 */

/**
 * The XML 1.0 control characters no document may carry, as a character class.
 *
 * Built from `String.fromCharCode` so `no-control-regex` cannot flag it statically, and built
 * ONCE: it was constructed inside the text escaper, which runs for every attribute value and
 * text child in the package, and again inside `validateObjectName` under a comment saying it was
 * the same set.
 */
const ILLEGAL_XML_CHARS_CLASS = ((cc: (n: number) => string) =>
	`[${cc(0)}-${cc(8)}${cc(11)}${cc(12)}${cc(14)}-${cc(31)}${cc(127)}]`)(String.fromCharCode)

/** The stripping form. Separate from {@link ILLEGAL_XML_CHARS} because a `g` regex is stateful under `.test`. */
const ILLEGAL_XML_CHARS_G = new RegExp(ILLEGAL_XML_CHARS_CLASS, 'g')

/** The detecting form; see {@link ILLEGAL_XML_CHARS_G} for why the two are not one regex. */
const ILLEGAL_XML_CHARS = new RegExp(ILLEGAL_XML_CHARS_CLASS)

/** Does `value` carry a character XML 1.0 forbids? The escapers strip them; `validateObjectName` reports them. */
export function hasIllegalXmlChars(value: string): boolean {
	return ILLEGAL_XML_CHARS.test(value)
}

/**
 * Replace special XML characters with HTML-encoded strings.
 *
 * ELEMENT TEXT ONLY. A literal tab, carriage return or line feed is left alone here because in
 * character data it *is* the content (`<a:t>` line breaks depend on it). Inside an attribute value
 * the same three characters are destroyed by the parser rather than preserved — see
 * {@link encodeXmlAttrValue}, which every attribute-emitting path must use instead.
 * @param {string | number} xml - value to encode (numbers are stringified, as callers pass counts/sizes)
 * @returns {string} escaped XML
 */
export function encodeXmlEntities(xml: string | number): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	// Strip XML 1.0 illegal control chars (e.g. \v) before escaping to prevent PowerPoint repair dialogs.
	return xml
		.toString()
		.replace(ILLEGAL_XML_CHARS_G, '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;')
}

/**
 * Escape a value destined for an XML **attribute**.
 *
 * Everything {@link encodeXmlEntities} does, plus the three whitespace characters that survive in
 * element text but not in an attribute: XML 1.0 section 3.3.3 requires a parser to normalise a
 * literal tab, carriage return or line feed inside an attribute value to a single space *before any
 * consumer sees it*. Carrying one across therefore requires a character reference, so a caller's
 * `objectName: 'Abschnitts-\nuberschrift'` reads back with its line break rather than a space.
 *
 * Deliberately a separate function rather than a widening of `encodeXmlEntities`: that helper also
 * escapes element text, where a raw newline is meaningful content emitted as-is, and escaping it
 * there would change bytes across every text-bearing part in the package.
 * @param {string | number} xml - value to encode (numbers are stringified, as callers pass ids/sizes)
 * @returns {string} escaped XML, safe to place between attribute quotes
 */
export function encodeXmlAttrValue(xml: string | number): string {
	return encodeXmlEntities(xml).replace(/\t/g, '&#9;').replace(/\n/g, '&#10;').replace(/\r/g, '&#13;')
}
