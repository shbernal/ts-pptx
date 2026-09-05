/**
 * ts-pptx: getting font bytes out of whatever the caller handed over.
 *
 * `registerFontMetrics` and `embedFont` are two public methods on one class, forty lines apart,
 * and each resolved its source itself. They agreed on `Uint8Array` and `ArrayBuffer` and
 * disagreed about what a `string` means: `registerFontMetrics('X', someBase64)` handed the
 * base64 to the runtime's file loader and failed to open it, while
 * `embedFont({ data: someBase64 })` decoded it. Neither doc said which, because each was only
 * describing its own half.
 *
 * The two really do want different string semantics -- `embedFont` names its path in `path` and
 * its bytes in `data`, so a string in `data` cannot be a filename -- so the answer is one
 * resolver with the difference stated as a flag rather than two resolvers with the difference
 * left implicit.
 */

import { InvalidOptionError } from './errors.js'
import { decodeBase64ToBytes } from './media/base64.js'
import type { RuntimeAdapter } from './runtime/types.js'

/** Anything a caller may hand a font-taking method: a path/URL, raw bytes, or base64 text. */
type FontSource = string | Uint8Array | ArrayBuffer

/**
 * Resolve a caller's font source to bytes.
 *
 * A `Uint8Array` is used as it stands, an `ArrayBuffer` is wrapped, and a `string` is read
 * according to `base64`: **false or absent** means a path or URL for the runtime's loader (what
 * `registerFontMetrics(face, source)` and `embedFont({ path })` take), **true** means base64
 * text (what `embedFont({ data })` takes, with or without a `data:` prefix).
 * @param source - the caller's value
 * @param runtime - the entry point's adapter, for the path/URL case
 * @param opts - `base64: true` to read a string as base64 rather than as a path
 * @param opts.base64 - see above
 * @param label - how to name the offending argument if the source is unusable
 */
export async function resolveFontBytes(
	source: FontSource | undefined,
	runtime: RuntimeAdapter,
	opts: { base64?: boolean } | undefined,
	label: string
): Promise<Uint8Array> {
	if (source instanceof Uint8Array) return source
	if (source instanceof ArrayBuffer) return new Uint8Array(source)
	if (typeof source === 'string') {
		if (!opts?.base64) return await runtime.loadFontData(source)
		// A caller may paste either a bare base64 body or a whole data URL; the decoder wants the
		// second, so a body with no comma gets the `fntdata` header the parts themselves carry.
		const decoded = decodeBase64ToBytes(source.includes(',') ? source : `application/x-fontdata;base64,${source}`)
		if (!decoded) throw new InvalidOptionError('font/invalid-base64', `${label} is not valid base64`)
		return decoded
	}
	throw new InvalidOptionError('font/missing-source', `${label} must be a font path, raw bytes, or base64 data`)
}
