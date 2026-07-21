/**
 * PptxGenJS: shared generator helpers.
 *
 * The small cross-cutting pieces the OOXML writers all need and that belong to
 * no single part:
 *   - XML text               encodeXmlEntities
 *   - Identifiers & naming    getUuid, validateObjectName, getDuplicateObjectNames
 *   - Slide relationships     getNewRelId, isHyperlinkRel
 *
 * DrawingML fragment builders moved to `gen/drawingml/{color,effect,fill,line}.ts`;
 * unit conversion to `units-internal.ts` (over the public primitives in `units.ts`);
 * base64/image-header decoding and media content types to `media/`.
 */

import { warn } from './log.js'
import type { PresSlideInternal } from './types/internal.js'

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
 * was duplicated across the slide-rels writer (`gen/slide/slide.ts`) and the inspect path
 * (pptxgen). For an internal slide-to-slide link, `rel.data === 'slide'` and
 * `rel.Target` is the 1-based target slide number.
 * @param {{ type: string }} rel - a slide relationship
 * @returns {boolean} true if the rel is any kind of hyperlink
 */
export function isHyperlinkRel(rel: { type: string }): boolean {
	return rel.type.toLowerCase().includes('hyperlink')
}
