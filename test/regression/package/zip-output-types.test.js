// Coverage for ZipWriter.generate / convertZipOutput (src/zip.ts): every
// ZIP_OUTPUT_TYPE the public `write({ outputType })` accepts must map fflate's
// Uint8Array onto the documented return shape, and an unknown type must throw
// rather than emit a degenerate value. The default write path only exercises
// 'nodebuffer', so the other shapes are otherwise unmeasured.
import { Buffer } from 'node:buffer'
import { describe, test } from 'vitest'
import TsPptx from '../../../dist/node.js'
import { assert, assertRejects } from '../../helpers.js'

const PK_MAGIC = [0x50, 0x4b] // "PK" — local file header of any zip

function makePres() {
	const pres = new TsPptx()
	pres.addSlide().addText('hi', { x: 1, y: 1, w: 2, h: 1 })
	return pres
}

/** First two bytes of `bytes` are the ZIP local-file-header magic "PK". */
function isZipBytes(bytes) {
	return bytes[0] === PK_MAGIC[0] && bytes[1] === PK_MAGIC[1]
}

describe('zip output types', () => {
	test('toBytes() returns a portable Uint8Array rather than a Node Buffer', async () => {
		const out = await makePres().toBytes()
		assert(out instanceof Uint8Array, `expected Uint8Array, got ${out?.constructor?.name}`)
		assert(!Buffer.isBuffer(out), 'toBytes() must not copy the archive into a Node Buffer')
		assert(isZipBytes(out), 'toBytes() output must start with the PK zip magic')
	})

	test("'uint8array' returns a Uint8Array of the archive", async () => {
		const out = await makePres().write({ outputType: 'uint8array' })
		assert(out instanceof Uint8Array, `expected Uint8Array, got ${out?.constructor?.name}`)
		assert(isZipBytes(out), 'Uint8Array output must start with the PK zip magic')
	})

	test("'arraybuffer' returns a standalone (non-shared) ArrayBuffer", async () => {
		const out = await makePres().write({ outputType: 'arraybuffer' })
		assert(out instanceof ArrayBuffer, `expected ArrayBuffer, got ${out?.constructor?.name}`)
		assert(isZipBytes(new Uint8Array(out)), 'ArrayBuffer output must start with the PK zip magic')
	})

	test("'nodebuffer' returns a Buffer", async () => {
		const out = await makePres().write({ outputType: 'nodebuffer' })
		assert(Buffer.isBuffer(out), `expected Buffer, got ${out?.constructor?.name}`)
		assert(isZipBytes(out), 'Buffer output must start with the PK zip magic')
	})

	test("'blob' returns a pptx-typed Blob", async () => {
		const out = await makePres().write({ outputType: 'blob' })
		assert(out instanceof Blob, `expected Blob, got ${out?.constructor?.name}`)
		assert(
			out.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			`unexpected Blob MIME type: ${out.type}`
		)
		assert(isZipBytes(new Uint8Array(await out.arrayBuffer())), 'Blob bytes must start with the PK zip magic')
	})

	test("'base64' returns a base64 string that decodes to the archive", async () => {
		const out = await makePres().write({ outputType: 'base64' })
		assert(typeof out === 'string', `expected string, got ${typeof out}`)
		assert(isZipBytes(Buffer.from(out, 'base64')), 'base64 output must decode to the PK zip magic')
	})

	test("'binarystring' returns a latin1 string of the archive bytes", async () => {
		const out = await makePres().write({ outputType: 'binarystring' })
		assert(typeof out === 'string', `expected string, got ${typeof out}`)
		assert(out.charCodeAt(0) === PK_MAGIC[0] && out.charCodeAt(1) === PK_MAGIC[1], 'binarystring must begin with "PK"')
	})

	test('an unsupported output type throws rather than emitting garbage', async () => {
		await assertRejects(
			() => makePres().write({ outputType: /** @type {any} */ ('bogus') }),
			/Unsupported zip output type/,
			'an unknown outputType'
		)
	})
})
