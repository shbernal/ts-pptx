import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, test } from 'vitest'
import TsPptx from '../../../dist/node.js'
import { assert, assertRejects, captureDiagnostics } from '../../helpers.js'

const BROKEN_SVG = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../browser/harness/broken.svg')

// Acceptance: a media asset that fails to load must, by default, reject the export with an
// actionable error that names the failing asset (the raw fs/network error alone does not say
// which path broke). An opt-in `onMediaError: 'placeholder'` degrades gracefully instead, so a
// single missing asset does not abort a best-effort/batch deck.

const BAD_PATH = '/definitely/does/not/exist/missing-image.png'

function deckWithMissingImage() {
	const pptx = new TsPptx()
	pptx.addSlide().addImage({ path: BAD_PATH, x: 1, y: 1, w: 2, h: 2 })
	return pptx
}

describe('media load failure policy', () => {
	test('default export rejects with an error naming the failing asset', async () => {
		const error = await assertRejects(
			() => deckWithMissingImage().write({ outputType: 'nodebuffer' }),
			new RegExp(BAD_PATH.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
			'export with a media asset that fails to load'
		)
		assert(error.cause !== undefined, 'wrapped error must chain the original cause')
	})

	test("onMediaError:'placeholder' substitutes a placeholder and resolves", async () => {
		// write({ outputType: 'nodebuffer' }) resolves to a Buffer; the return type is the union of all targets.
		const buf = /** @type {Buffer} */ (
			await deckWithMissingImage().write({ outputType: 'nodebuffer', onMediaError: 'placeholder' })
		)
		assert(buf && buf.length > 0, 'placeholder mode must produce a non-empty package')
	})

	test("toBytes() forwards onMediaError:'placeholder'", async () => {
		const bytes = await deckWithMissingImage().toBytes({ onMediaError: 'placeholder' })
		assert(bytes.length > 0, 'toBytes placeholder mode must produce a non-empty package')
	})
})

// Node has no rasterizer, so every SVG's PNG fallback is the placeholder: that is its normal output,
// not a failure, or every SVG would fail the default export here. So an SVG a browser cannot decode
// warns there and nothing here, and under either policy its own bytes are kept.
describe('SVG preview on a runtime with no rasterizer', () => {
	for (const { label, image } of [
		{ label: 'by path', image: { path: BROKEN_SVG } },
		{ label: 'inline', image: { svg: 'this is not SVG either' } },
	]) {
		test(`an SVG given ${label} is written with its own bytes and no warning`, async () => {
			const pptx = new TsPptx()
			pptx.addSlide().addImage({ ...image, x: 1, y: 1, w: 1, h: 1 })
			const { result, codes } = await captureDiagnostics(() =>
				pptx.write({ outputType: 'nodebuffer', onMediaError: 'placeholder' })
			)
			expect(codes).toEqual([])

			const zip = await JSZip.loadAsync(/** @type {Buffer} */ (result))
			const svgParts = Object.keys(zip.files).filter((name) => /^ppt\/media\/.*\.svg$/.test(name))
			expect(svgParts).toHaveLength(1)
			const written = await zip.file(svgParts[0]).async('string')
			const expected = 'path' in image ? await readFile(image.path, 'utf8') : image.svg
			expect(written).toBe(expected)
		})
	}
})
