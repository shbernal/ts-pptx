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
export function svgMarkupToDataUri(svg: string): string {
	return `data:image/svg+xml;base64,${bytesToBase64(new TextEncoder().encode(svg))}`
}

/**
 * Encode raw bytes as base64 (no `data:` prefix) using only `btoa`, so the result is the same
 * on Node, in a browser, and in a runtime that has neither `Buffer` nor `FileReader`.
 *
 * Chunked because `String.fromCharCode(...spread)` blows the argument-count limit on a payload
 * of any real size — a multi-megabyte video rel would throw rather than encode.
 * @param {Uint8Array} bytes - the payload
 * @returns {string} base64 text
 */
export function bytesToBase64(bytes: Uint8Array): string {
	const CHUNK = 0x8000
	let binary = ''
	for (let i = 0; i < bytes.length; i += CHUNK) binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
	return btoa(binary)
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
