// Input-matrix coverage for the read path's zip entrypoint (`readZip` /
// `toUint8Array` in src/zip.ts) and the OPC loader on top of it.
//
// The declared input union — filesystem path (string), number[], Uint8Array,
// ArrayBuffer, Blob, and a Promise wrapper — is the exact surface where the
// "string is latin1 binary content" footgun lived: a path silently became
// garbage bytes → the opaque "Not a valid ZIP archive". These tests drive every
// branch of that union with the same real deck bytes and assert each error
// branch reports a clear, specific message rather than the opaque zip error.

import { mkdtempSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { readZip } from '../../dist/zip.js'
import { OpcPackage } from '../../dist/read.js'
import { build, assert, assertEqual } from '../helpers.js'

// One real .pptx worth of bytes, shared across the input-shape cases so each
// branch is proven to decode identical content to the same part set.
const { buf } = await build((p) => {
	p.addSlide().addText('zip input matrix', { x: 1, y: 1, w: 3, h: 0.5 })
})
const SLIDE_PATH = 'ppt/slides/slide1.xml'
const CONTENT_TYPES_PATH = '[Content_Types].xml'

function bytes() {
	// Fresh copy per case so no test can observe another's mutation/detachment.
	return Uint8Array.from(buf)
}

function assertDecodesDeck(entries, label) {
	assert(entries instanceof Map, `${label}: readZip returns a Map`)
	assert(entries.has(CONTENT_TYPES_PATH), `${label}: [Content_Types].xml present`)
	assert(entries.has(SLIDE_PATH), `${label}: slide part present`)
	assert(entries.get(SLIDE_PATH).length > 0, `${label}: slide part has bytes`)
}

describe('readZip input matrix', () => {
	test('Uint8Array input decodes the archive', async () => {
		assertDecodesDeck(await readZip(bytes()), 'Uint8Array')
	})

	test('Node Buffer input decodes the archive', async () => {
		// Buffer is a Uint8Array subclass; the common Node caller passes one.
		assertDecodesDeck(await readZip(Buffer.from(bytes())), 'Buffer')
	})

	test('ArrayBuffer input decodes the archive', async () => {
		const b = bytes()
		const ab = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)
		assert(ab instanceof ArrayBuffer, 'precondition: ArrayBuffer built')
		assertDecodesDeck(await readZip(ab), 'ArrayBuffer')
	})

	test('number[] input decodes the archive', async () => {
		assertDecodesDeck(await readZip(Array.from(bytes())), 'number[]')
	})

	test('Blob input decodes the archive', async () => {
		assert(typeof Blob !== 'undefined', 'precondition: Blob is available in this runtime')
		assertDecodesDeck(await readZip(new Blob([bytes()])), 'Blob')
	})

	test('Promise-wrapped input is awaited and decoded', async () => {
		assertDecodesDeck(await readZip(Promise.resolve(bytes())), 'Promise<Uint8Array>')
	})

	test('filesystem path (string) input is read from disk, not treated as bytes', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'pptx-zip-input-'))
		try {
			const filePath = join(dir, 'deck.pptx')
			writeFileSync(filePath, bytes())
			assertDecodesDeck(await readZip(filePath), 'string path')
		} finally {
			rmSync(dir, { recursive: true, force: true })
		}
	})

	test('directory markers are dropped from the entry map', async () => {
		// fflate surfaces `foo/` directory keys; readZip filters them so consumers
		// only ever see real parts. Build an archive that contains one.
		const zip = new JSZip()
		zip.folder('emptydir')
		zip.file('real.xml', '<x/>')
		const withDir = await zip.generateAsync({ type: 'uint8array' })
		const entries = await readZip(withDir)
		assert(entries.has('real.xml'), 'real part survives')
		assert(
			[...entries.keys()].every((k) => !k.endsWith('/')),
			`no directory markers survive; got: ${[...entries.keys()].join(', ')}`
		)
	})
})

describe('readZip error branches report a specific message', () => {
	test('corrupt bytes throw "Not a valid ZIP archive" with the decode cause attached', async () => {
		let error = null
		try {
			await readZip(new Uint8Array([1, 2, 3, 4, 5]))
		} catch (err) {
			error = err
		}
		assert(error, 'corrupt bytes throw')
		assert(error.message.includes('Not a valid ZIP archive'), `got: ${error.message}`)
		assert('cause' in error && error.cause, 'the underlying decode error is attached as cause')
	})

	test('unsupported input type names the accepted shapes', async () => {
		let error = null
		try {
			// A bare number is none of the accepted input shapes.
			await readZip(42)
		} catch (err) {
			error = err
		}
		assert(error, 'unsupported type throws')
		assert(error.message.includes('Unsupported zip input type'), `got: ${error.message}`)
		assert(
			!error.message.includes('Not a valid ZIP archive'),
			'an unsupported type is not misreported as a corrupt archive'
		)
	})

	test('a missing filesystem path names the path and is not misreported as a corrupt archive', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'pptx-zip-input-'))
		let error = null
		try {
			await readZip(join(dir, 'does-not-exist.pptx'))
		} catch (err) {
			error = err
		} finally {
			rmSync(dir, { recursive: true, force: true })
		}
		assert(error, 'a missing path throws')
		assert(error.message.includes('does-not-exist.pptx'), `error names the path; got: ${error.message}`)
		assert(!error.message.includes('Not a valid ZIP archive'), 'a missing path is not misreported as a corrupt archive')
	})
})

describe('OpcPackage.load over the same input surface', () => {
	test('loads from a filesystem path (string)', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'pptx-zip-input-'))
		try {
			const filePath = join(dir, 'deck.pptx')
			writeFileSync(filePath, bytes())
			const pkg = await OpcPackage.load(filePath)
			assert(pkg.part('/ppt/slides/slide1.xml'), 'slide part loaded from path')
		} finally {
			rmSync(dir, { recursive: true, force: true })
		}
	})

	test('loads from an ArrayBuffer', async () => {
		const b = bytes()
		const ab = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)
		const pkg = await OpcPackage.load(ab)
		assert(pkg.part('/ppt/slides/slide1.xml'), 'slide part loaded from ArrayBuffer')
	})

	test('a valid zip that is not an OPC package is rejected by name', async () => {
		// A structurally-valid archive with no [Content_Types].xml is a zip but not
		// an OPC package; the loader must say so rather than fail obscurely later.
		const zip = new JSZip()
		zip.file('hello.txt', 'not an OPC package')
		const notOpc = await zip.generateAsync({ type: 'uint8array' })
		let error = null
		try {
			await OpcPackage.load(notOpc)
		} catch (err) {
			error = err
		}
		assert(error, 'a non-OPC zip is rejected')
		assertEqual(error.message, 'Not an OPC package: missing [Content_Types].xml', 'the loader names the missing part')
	})
})
