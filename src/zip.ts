import { strToU8, unzipSync, zipSync, type Unzipped, type Zippable, type ZipOptions } from 'fflate'
import type { JSZIP_OUTPUT_TYPE } from './core-enums.js'

/**
 * ZIP backend seam for the write path.
 *
 * fflate is functional — one `{path: bytes}` map in, one `Uint8Array` out — so
 * every contact with the backend is isolated here rather than spread across the
 * incremental builder calls the JSZip API encouraged. Callers accumulate entries
 * with {@link ZipWriter.add} (strings are UTF-8 encoded; already-decoded media
 * passes bytes directly) and finalize with {@link ZipWriter.generate}, which maps
 * fflate's single `Uint8Array` to the requested output shape.
 *
 * Unlike JSZip there is no folder/directory concept: keys are full slash-paths and
 * fflate emits no directory entries, so the package carries only real parts.
 */
/**
 * Fixed modification time (2001-01-01 UTC) stamped on every entry so archive
 * bytes are reproducible across runs. fflate encodes DOS dates, which are bounded
 * to 1980-2099, so 0/epoch is not usable here.
 */
const FIXED_MTIME = Date.UTC(2001, 0, 1)

export class ZipWriter {
	readonly #entries: Zippable = {}

	/**
	 * Add a package part.
	 * @param path - full zip path, e.g. `ppt/slides/slide1.xml`
	 * @param data - XML string (UTF-8 encoded) or already-decoded bytes (media)
	 * @param opts.store - skip DEFLATE for this entry (already-compressed media, #1006)
	 */
	add (path: string, data: string | Uint8Array, opts?: { store?: boolean }): void {
		const bytes = typeof data === 'string' ? strToU8(data) : data
		// Pin mtime so archive bytes are reproducible across runs (stable fixtures).
		const fileOpts: ZipOptions = { mtime: FIXED_MTIME }
		if (opts?.store) fileOpts.level = 0
		this.#entries[path] = [bytes, fileOpts]
	}

	/**
	 * Compress all accumulated entries to raw zip bytes.
	 * @param compression - false stores every entry uncompressed (level 0)
	 */
	toBytes (compression = true): Uint8Array {
		// Global level is the per-entry default; entries added with `store` keep their level 0.
		return zipSync(this.#entries, { level: compression ? 6 : 0, mtime: FIXED_MTIME })
	}

	/**
	 * Finalize the archive in the requested output shape.
	 * @param type - JSZip-compatible output type
	 * @param opts.compression - false stores every entry uncompressed
	 */
	async generate (type: JSZIP_OUTPUT_TYPE, opts: { compression: boolean }): Promise<string | ArrayBuffer | Blob | Uint8Array> {
		const bytes = this.toBytes(opts.compression)
		return convertZipOutput(bytes, type)
	}
}

/**
 * Inputs the read path accepts, mirroring what JSZip's `loadAsync` auto-detected.
 * A `Promise` wrapper is allowed so callers can forward an unawaited byte source.
 */
type ZipInputValue = string | number[] | Uint8Array | ArrayBuffer | Blob
export type ZipInput = ZipInputValue | Promise<ZipInputValue>

/**
 * Decompress a zip archive to a `path → bytes` map in central-directory order.
 *
 * Read-path counterpart to {@link ZipWriter}: fflate's `unzipSync` needs a
 * `Uint8Array`, so the JSZip-style input auto-detection that `loadAsync` did is
 * reproduced by {@link toUint8Array} first. Directory markers (keys ending in
 * `/`) are dropped — fflate surfaces them but no consumer wants empty-dir
 * entries, and the write path emits none.
 */
export async function readZip (input: ZipInput): Promise<Map<string, Uint8Array>> {
	const bytes = await toUint8Array(input)
	let entries: Unzipped
	try {
		entries = unzipSync(bytes)
	} catch (cause) {
		throw new Error('Not a valid ZIP archive', { cause })
	}
	const out = new Map<string, Uint8Array>()
	for (const [path, body] of Object.entries(entries)) {
		if (path.endsWith('/')) continue
		out.set(path, body)
	}
	return out
}

/**
 * Normalize a read-path input to `Uint8Array`, matching JSZip's default
 * `loadAsync` handling: strings are treated as latin1 binary strings (not
 * base64), `ArrayBuffer`/`Blob`/`number[]` are copied/wrapped, and a `Buffer`
 * passes through as the `Uint8Array` it already is.
 */
async function toUint8Array (input: ZipInput): Promise<Uint8Array> {
	const data = await input
	if (data instanceof Uint8Array) return data
	if (data instanceof ArrayBuffer) return new Uint8Array(data)
	if (typeof data === 'string') {
		const bytes = new Uint8Array(data.length)
		for (let i = 0; i < data.length; i++) bytes[i] = data.charCodeAt(i) & 0xff
		return bytes
	}
	if (Array.isArray(data)) return Uint8Array.from(data)
	if (typeof Blob !== 'undefined' && data instanceof Blob) return new Uint8Array(await data.arrayBuffer())
	throw new Error('Unsupported zip input type; expected string, number[], Uint8Array, ArrayBuffer, or Blob')
}

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

/** Map fflate's `Uint8Array` to a JSZip-compatible output type. */
function convertZipOutput (bytes: Uint8Array, type: JSZIP_OUTPUT_TYPE): string | ArrayBuffer | Blob | Uint8Array {
	switch (type) {
		case 'uint8array':
			return bytes
		case 'arraybuffer': {
			// Copy into a fresh (non-shared) ArrayBuffer; fflate's output buffer is
			// typed ArrayBufferLike, which may be a SharedArrayBuffer.
			const out = new ArrayBuffer(bytes.byteLength)
			new Uint8Array(out).set(bytes)
			return out
		}
		case 'nodebuffer':
			// Buffer.from(Uint8Array) copies, so the result owns its memory.
			return Buffer.from(bytes)
		case 'blob':
			// Copy into an ArrayBuffer-backed view so it is a valid BlobPart.
			return new Blob([new Uint8Array(bytes)], { type: PPTX_MIME })
		case 'base64':
			return bytesToBinaryString(bytes, true)
		case 'binarystring':
			return bytesToBinaryString(bytes, false)
		default: {
			const exhaustive: never = type
			throw new Error(`Unsupported zip output type: ${String(exhaustive)}`)
		}
	}
}

/**
 * Build a latin1 binary string from bytes (chunked to dodge the argument-count
 * limit of `String.fromCharCode(...spread)` on large archives), optionally
 * base64-encoding it. `btoa` is isomorphic (Node >=16 and browsers).
 */
function bytesToBinaryString (bytes: Uint8Array, base64: boolean): string {
	let binary = ''
	const CHUNK = 0x8000
	for (let i = 0; i < bytes.length; i += CHUNK) {
		binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
	}
	return base64 ? btoa(binary) : binary
}
