// Contract for `TsPptx.toParts()` (src/presentation.ts → buildPackageParts): the unzipped
// parts it returns must be the SAME parts `write()` compresses into the `.pptx` — same path
// set, same emission order, and byte-identical per part. `toParts` is the public seam over
// the assembly pipeline, so a drift between it and `write()` (a part only one path emits, a
// reordering, a byte difference) would be a silent contract break. JSZip reads the `write()`
// output back as an independent oracle (the write path zips with fflate), mirroring helpers.js.
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import TsPptx from '../../../dist/node.js'
import { assert } from '../../helpers.js'

/** Author an identical text-only deck each call so two builds differ only in core.xml timestamps. */
function makePres() {
	const pres = new TsPptx()
	pres.addSlide().addText('to-parts contract', { x: 1, y: 1, w: 4, h: 1 })
	pres.addSlide().addText('second slide', { x: 1, y: 1, w: 4, h: 1 })
	return pres
}

// docProps/core.xml carries `new Date()` dcterms timestamps; blank them so a build straddling a
// clock tick doesn't make byte-equality flaky. Every other part is deterministic for this deck.
const decoder = new TextDecoder()
function stripCoreTimestamps(bytes) {
	return decoder
		.decode(bytes)
		.replace(/(<dcterms:(?:created|modified)[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/g, '$1$2')
}

describe('toParts()', () => {
	test('returns the same parts write() emits: path set, order, and per-part bytes', async () => {
		const parts = await makePres().toParts()

		// Independent oracle: unzip a real write() output with JSZip (not the fflate write path).
		const buf = /** @type {Uint8Array} */ (await makePres().write({ outputType: 'uint8array' }))
		const zip = await JSZip.loadAsync(buf)
		// JSZip preserves central-directory order, which is the write path's insertion order.
		const zipPaths = Object.keys(zip.files).filter((p) => !zip.files[p].dir)

		// Order (hence path set) must match exactly.
		assert(
			JSON.stringify(parts.map((p) => p.path)) === JSON.stringify(zipPaths),
			`part order/paths differ.\n toParts: ${parts.map((p) => p.path).join(', ')}\n write:   ${zipPaths.join(', ')}`
		)

		// Per-part bytes must match (core.xml compared with its dcterms timestamps blanked).
		for (const part of parts) {
			assert(part.data instanceof Uint8Array, `part ${part.path} data must be a Uint8Array`)
			const oracle = await zip.files[part.path].async('uint8array')
			if (part.path === 'docProps/core.xml') {
				assert(
					stripCoreTimestamps(part.data) === stripCoreTimestamps(oracle),
					`core.xml bytes differ once timestamps are normalized`
				)
			} else {
				assert(
					Buffer.from(part.data).equals(Buffer.from(oracle)),
					`bytes differ for part ${part.path} (${part.data.length} vs ${oracle.length})`
				)
			}
		}
	})

	test('each call returns fresh Uint8Array buffers (no shared mutable state)', async () => {
		const pres = makePres()
		const a = await pres.toParts()
		const b = await pres.toParts()
		const byPath = new Map(b.map((p) => [p.path, p]))
		for (const part of a) {
			const other = byPath.get(part.path)
			assert(other, `part ${part.path} missing from the second toParts() call`)
			assert(part.data !== other.data, `part ${part.path} shares a Uint8Array across calls`)
			assert(part.data.buffer !== other.data.buffer, `part ${part.path} shares an ArrayBuffer across calls`)
		}
	})

	test('the public part shape is { path, data } only — no internal store hint leaks', async () => {
		const [part] = await makePres().toParts()
		assert(
			JSON.stringify(Object.keys(part).sort()) === JSON.stringify(['data', 'path']),
			`unexpected public part keys: ${Object.keys(part).join(', ')}`
		)
	})
})
