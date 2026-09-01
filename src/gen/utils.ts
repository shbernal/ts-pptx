/**
 * ts-pptx: shared generator helpers.
 *
 * The small cross-cutting pieces the OOXML writers all need and that belong to
 * no single part:
 *   - XML text               encodeXmlEntities
 *   - Identifiers & naming    getUuid, validateObjectName, getDuplicateObjectNames
 *   - Slide relationships     getNewRelId, isHyperlinkRel, mediaSlideKey
 *   - Option normalization    setOrClear
 *
 * DrawingML fragment builders moved to `gen/drawingml/{color,effect,fill,line}.ts`;
 * unit conversion to `units-internal.ts` (over the public primitives in `units.ts`);
 * base64/image-header decoding and media content types to `media/`.
 */

import { warn } from '../diagnostics.js'
import type { OptionalKeysOf, PresSlideInternal } from '../types/internal.js'

/**
 * Write a normalized value onto an option bag, spelling "no value" as an *absent* key.
 *
 * The generator's option bags are read with plain truthiness or `?.`, so absent and
 * present-but-`undefined` look the same to a reader — but not to a spread, and these bags are
 * spread: a placeholder's options onto a slide's, a column default under a cell's own, a combo
 * subchart's overrides onto the chart's. There the side that wins is decided by whether the key
 * exists, not by what it holds, so a normalizer that rejects a value has to remove it rather than
 * paper over it. One spelling of absent, which is the invariant `compact()` keeps on the read
 * side (`script/from-read/values.ts`).
 *
 * Only optional keys are accepted, so this cannot be used to unset something required.
 * @param bag - the options bag to write onto
 * @param key - the key to set or remove
 * @param value - the normalized value, or `undefined` to remove the key
 */
export function setOrClear<T extends object, K extends OptionalKeysOf<T> & keyof T>(bag: T, key: K, value: T[K]): void {
	if (value === undefined) delete bag[key]
	else bag[key] = value
}

/**
 * Fill the `x`/`y` placeholders of a hex pattern with random nibbles.
 *
 * **Not replaceable by `crypto.randomUUID()`**, which exists on Node 24 and in every
 * browser. This takes a *format string*, and half its callers pass a partial one:
 * `gen/chart/plot-scatter.ts` asks for a bare tail (`-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)
 * to splice onto a padded point index, and its other two calls want a GUID shape with
 * no version nibble. Only `gen/define/zoom.ts` and `presentation.ts`, which pass a
 * v4-shaped pattern, could switch, and switching only those would leave two generators
 * where there is now one. The analysis is recorded here so it does not get re-derived:
 * the swap is low value and non-zero risk, and it is not on the roadmap.
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - the pattern; `x` becomes a random nibble, `y` one of 8/9/a/b
 * @returns {string} the filled pattern
 */
export function getUuid(uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		const r = (Math.random() * 16) | 0
		const v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
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

/**
 * Reverse {@link encodeXmlAttrValue}, so a stored attribute value reads back in the spelling its
 * caller used. Entity replacement runs in the mirror order, with `&amp;` last: encoded output only
 * ever opens an entity with a `&` it wrote itself, so no earlier replacement can manufacture one.
 *
 * Full invertibility is not on offer and is not what a consumer needs — the encoder *strips* XML
 * 1.0 illegal control characters, and no inverse brings those back. The guarantee is the round
 * trip: `encodeXmlAttrValue(decodeXmlAttrValue(stored)) === stored` for every `stored` the encoder
 * produced. It holds for a caller who authored `&amp;` literally (decode yields `&amp;`, which
 * re-encodes to `&amp;amp;`, the stored value) and for one whose name carried a stripped control
 * character (stripping is idempotent).
 *
 * That round trip is the whole point. A name read back off a stored object has to resolve when it
 * is handed to a lookup that escapes before comparing — `groupObjects()` is the one that does —
 * and the *stored* spelling would escape a second time and match nothing.
 * @param {string} value - a stored attribute value, as written by `encodeXmlAttrValue`
 * @returns {string} the spelling the caller passed
 */
export function decodeXmlAttrValue(value: string): string {
	return value
		.replace(/&#9;/g, '\t')
		.replace(/&#10;/g, '\n')
		.replace(/&#13;/g, '\r')
		.replace(/&apos;/g, "'")
		.replace(/&quot;/g, '"')
		.replace(/&gt;/g, '>')
		.replace(/&lt;/g, '<')
		.replace(/&amp;/g, '&')
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
		warn(
			'object-name/empty',
			`${kind} objectName is empty or whitespace-only; it will not provide a stable Selection Pane identity.`
		)
		return name
	}
	// Same illegal-XML-char set that `encodeXmlEntities` strips; detect so the caller knows the name will change.
	const cc = String.fromCharCode
	const illegalXmlCharsRe = new RegExp(`[${cc(0)}-${cc(8)}${cc(11)}${cc(12)}${cc(14)}-${cc(31)}${cc(127)}]`)
	if (illegalXmlCharsRe.test(name)) {
		warn(
			'object-name/control-characters',
			`${kind} objectName "${name}" contains control characters that will be stripped, changing the stored name.`
		)
	}
	if (name.length > MAX_OBJECT_NAME_LENGTH) {
		warn(
			'object-name/too-long',
			`${kind} objectName exceeds ${MAX_OBJECT_NAME_LENGTH} characters and may not be preserved by PowerPoint.`
		)
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
 * Get a new rel ID (rId) for charts, media, etc.
 *
 * Counting is the normal case: a slide's rels are minted in order, so one past the
 * total is free. It is not enough on its own, because a rel can arrive on a slide
 * carrying an id minted somewhere else — auto-paging re-registers a repeated header
 * row's hyperlink on each overflow slide under the id stamped on the shared
 * hyperlink object (`gen/define/hyperlinks.ts`). That id is then held but not
 * counted-to, so the next mint could hand out the same number and the part would
 * declare two `Relationship` elements sharing an `Id` — invalid OPC, and PowerPoint
 * resolves such a reference to whichever comes first. Hence the scan: start where
 * the count says, then step over anything the slide already holds.
 * @param {PresSlideInternal} target - the slide to use
 * @returns {number} the lowest free rId at or above the current rel count plus 1
 */
export function getNewRelId(target: PresSlideInternal): number {
	const held = new Set<number>()
	target._rels.forEach((rel) => held.add(rel.rId))
	target._relsChart.forEach((rel) => held.add(rel.rId))
	target._relsMedia.forEach((rel) => held.add(rel.rId))

	let rId = target._rels.length + target._relsChart.length + target._relsMedia.length + 1
	while (held.has(rId)) rId++
	return rId
}

/**
 * Whether a slide relationship is a hyperlink (external URL or internal slide
 * link). The relationship `type` is stringly-typed (`'hyperlink'`, `'online'`,
 * mixed-case variants), so this centralizes the case-insensitive predicate that
 * was duplicated across the slide-rels writer (`gen/slide/slide.ts`) and the inspect path
 * (presentation.ts). For an internal slide-to-slide link, `rel.data === 'slide'` and
 * `rel.Target` is the 1-based target slide number.
 * @param {{ type: string }} rel - a slide relationship
 * @returns {boolean} true if the rel is any kind of hyperlink
 */
export function isHyperlinkRel(rel: { type: string }): boolean {
	return rel.type.toLowerCase().includes('hyperlink')
}

/**
 * The slide-scoped segment of a media part's name.
 *
 * Media targets are namespaced by the slide that registered them so that slide
 * master (`sm`) and slide layout (`sl-N`) media never collide with regular slide
 * media names in large decks. A layout is identified by the `_slideNum >= 1000`
 * convention the layout builder stamps on it; a master has no `_slideNum` at all.
 * @param {PresSlideInternal} target - the slide, layout or master registering the media
 * @returns {string} the key to splice into the media part name
 */
export function mediaSlideKey(target: PresSlideInternal): string {
	if (target._slideNum == null) return 'sm'
	return target._slideNum >= 1000 ? `sl-${target._slideNum}` : `${target._slideNum}`
}
