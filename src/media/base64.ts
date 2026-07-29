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
