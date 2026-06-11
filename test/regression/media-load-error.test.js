import { describe, test } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { assert } from '../helpers.js'

// Acceptance: a media asset that fails to load must, by default, reject the export with an
// actionable error that names the failing asset (the raw fs/network error alone does not say
// which path broke). An opt-in `onMediaError: 'placeholder'` degrades gracefully instead, so a
// single missing asset does not abort a best-effort/batch deck (upstream gitbrent/PptxGenJS#1310).

const BAD_PATH = '/definitely/does/not/exist/missing-image.png'

function deckWithMissingImage() {
	const pptx = new PptxGenJS()
	pptx.addSlide().addImage({ path: BAD_PATH, x: 1, y: 1, w: 2, h: 2 })
	return pptx
}

describe('media load failure policy (upstream #1310)', () => {
	test('default export rejects with an error naming the failing asset', async () => {
		let threw = false
		try {
			await deckWithMissingImage().write({ outputType: 'nodebuffer' })
		} catch (ex) {
			threw = true
			assert(String(ex.message).includes(BAD_PATH), `error must name the failing media path; got: ${ex.message}`)
			assert(ex.cause !== undefined, 'wrapped error must chain the original cause')
		}
		assert(threw, 'export must reject by default when a media asset fails to load')
	})

	test("onMediaError:'placeholder' substitutes a placeholder and resolves", async () => {
		const buf = await deckWithMissingImage().write({ outputType: 'nodebuffer', onMediaError: 'placeholder' })
		assert(buf && buf.length > 0, 'placeholder mode must produce a non-empty package')
	})
})
