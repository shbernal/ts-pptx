/**
 * Base64 payload handling for media parts — isomorphic across Node and browsers
 * (`TextEncoder`/`atob`/`btoa` only, no `Buffer`).
 */

/**
 * Encode raw SVG markup as a base64 `image/svg+xml` data URI.
 * - lets callers pass inline SVG to `addImage({ svg })` without hand-rolling base64
 * - isomorphic and UTF-8 safe: uses the global `TextEncoder`/`btoa` (Node and browsers)
 * @param {string} svg - SVG markup, e.g. `'<svg ...>...</svg>'`
 * @returns {string} a `data:image/svg+xml;base64,...` URI
 */
import { bytesToBinaryString } from './binary-string.js'

export function svgMarkupToDataUri(svg: string): string {
	return `data:image/svg+xml;base64,${bytesToBase64(new TextEncoder().encode(svg))}`
}

/**
 * Encode raw bytes as base64 (no `data:` prefix) using only `btoa`, so the result is the same
 * on Node, in a browser, and in a runtime that has neither `Buffer` nor `FileReader`.
 * @param {Uint8Array} bytes - the payload
 * @returns {string} base64 text
 */
export function bytesToBase64(bytes: Uint8Array): string {
	return btoa(bytesToBinaryString(bytes))
}

/**
 * Present a freshly loaded media payload as a `data:` URI.
 *
 * A {@link RuntimeAdapter.loadMedia} implementation returns whichever encoding its host
 * makes cheap — the browser adapter a full URI, because `FileReader.readAsDataURL` is what
 * decodes a blob without blocking; Node and the neutral adapter raw base64, because neither
 * has a `FileReader`. This is where the two become one, so nothing downstream of the load
 * has to ask which it got.
 *
 * The mime label is documentary: {@link decodeBase64ToBytes} ignores it, and the one consumer
 * that needs a real URI (`image.src`, in the browser adapter's SVG preview) is handed the
 * adapter's own. It is taken from the rel anyway rather than hard-coded, so an already-correct
 * label is not overwritten with a wrong one.
 * @param {string} payload - what the adapter returned: raw base64, or an already-formed URI
 * @param {string} contentType - the rel's content type, e.g. `image/png`
 * @returns {string} a `data:` URI
 */
export function toMediaDataUri(payload: string, contentType: string): string {
	return payload.startsWith('data:') ? payload : `data:${contentType};base64,${payload}`
}

/**
 * Where the payload begins in an inline `data:` URI, or `-1` when the value carries no
 * `base64,` header and is therefore raw base64 already.
 *
 * The one reading of that header, because the two functions below used to spell it separately and
 * disagree: {@link hasBase64Header} lower-cased before looking, {@link decodeBase64ToBytes} did
 * not. RFC 2397 does not case the `;base64` token, so `data:image/png;BASE64,…` passed every
 * definer's check and then decoded from index 0, which `atob` rejects. The media part was written
 * empty, with nothing said. One search means the two can no longer answer differently.
 *
 * Matched with a regular expression rather than by lower-casing and indexing, because
 * `toLowerCase` is not length-preserving for every code point and a shifted index would slice the
 * payload apart.
 */
function base64PayloadStart(data: string): number {
	const match = /base64,/i.exec(data)
	return match ? match.index + match[0].length : -1
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
	const start = base64PayloadStart(b64)
	const payload = (start >= 0 ? b64.slice(start) : b64).replace(/\s/g, '')
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
 * Whether an inline payload carries the `base64,` header the definers ask image and media bytes
 * to arrive with. A value that is not a string has none.
 * @param data - the caller's `data`
 */
export function hasBase64Header(data: unknown): boolean {
	return typeof data === 'string' && base64PayloadStart(data) >= 0
}
