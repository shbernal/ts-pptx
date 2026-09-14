/**
 * Raw bytes as a latin1 binary string, one character per byte: what `btoa` takes, and what the
 * `zip` entry's `binarystring` output returns as it is.
 *
 * A module of its own, importing nothing, for the reason `ooxml/pptx-content-type.ts` states: the
 * `zip` entry needs this one function, and imported from `base64.ts` it brought that module's other
 * exports along and pushed the entry past its size budget.
 *
 * Chunked because `String.fromCharCode(...spread)` blows the argument-count limit on a payload
 * of any real size — a multi-megabyte video rel would throw rather than encode.
 * @param {Uint8Array} bytes - the payload
 * @returns {string} one character per byte
 */
export function bytesToBinaryString(bytes: Uint8Array): string {
	const CHUNK = 0x8000
	let binary = ''
	for (let i = 0; i < bytes.length; i += CHUNK) binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
	return binary
}
