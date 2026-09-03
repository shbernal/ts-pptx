import { describe, test } from 'vitest'
import TsPptx from '../../../dist/node.js'
import { assert, assertRejects } from '../../helpers.js'

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
