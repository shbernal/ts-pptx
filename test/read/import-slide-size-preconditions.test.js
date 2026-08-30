// Every import entry point starts by asking whether the two decks share a canvas, and the five
// that asked disagreed about what an *unknown* size means. `importSlide` reported a deck with no
// `p:sldSz` as `import/slide-size-unknown`, but only when a rescale had been requested; the other
// four folded it into `import/slide-size-mismatch` and printed the word "unknown" in the message.
// So a consumer branching on `err.code` could not tell "these decks are different sizes" from "I
// could not read a size" anywhere but one method, and `import/slide-size-unknown` existed in
// `codes.ts` for a case one of five call sites could raise.
//
// The two conditions are separate everywhere now, and the codes mean what their names say. This
// pins both halves at every entry point, because "they all agree" is the whole claim.
import { describe, test } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assertEqual } from '../helpers.js'

/** A one-page deck at the given layout. */
async function deck(layout = 'LAYOUT_16x9') {
	const pptx = new TsPptx()
	pptx.layout = layout
	pptx.addSlide().addText('page', { x: 1, y: 1, w: 4, h: 1 })
	return pptx.write({ outputType: 'uint8array' })
}

/**
 * The same deck with its `<p:sldSz/>` removed — the only way a real package produces a
 * `slideSize` of `null`, and not something the authoring API can express.
 */
async function sizelessDeck() {
	const zip = await JSZip.loadAsync(await deck())
	const xml = await zip.file('ppt/presentation.xml').async('string')
	zip.file('ppt/presentation.xml', xml.replace(/<p:sldSz[^>]*\/>/, ''))
	return zip.generateAsync({ type: 'uint8array' })
}

/** The `code` of whatever `fn` throws, or `null` if it returned. */
async function codeOf(fn) {
	try {
		await fn()
		return null
	} catch (err) {
		return err?.code ?? null
	}
}

describe('slide-size preconditions agree across every import entry point', () => {
	// The four entry points whose source is another loaded `Presentation`. `appendSlides` is
	// the fifth and takes a *generator* rather than a package, so it gets its own cases below.
	/** @type {[string, (target: any, source: any) => unknown][]} */
	const ENTRY_POINTS = [
		['importSlide', (target, source) => target.importSlide(source, 0)],
		['importSlides', (target, source) => target.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])],
		['importSlideMasters', (target, source) => target.importSlideMasters(source)],
		['importShapes', (target, source) => target.importShapes(target.slides[0], source.slides[0], [0])],
	]

	/** A generator deck at the given layout — what `appendSlides` takes as its source. */
	function generator(layout = 'LAYOUT_16x9') {
		const pptx = new TsPptx()
		pptx.layout = layout
		pptx.addSlide().addText('page', { x: 1, y: 1, w: 4, h: 1 })
		return pptx
	}

	test('two known sizes that differ are a mismatch, everywhere', async () => {
		for (const [name, run] of ENTRY_POINTS) {
			const target = await Presentation.load(await deck('LAYOUT_16x9'))
			const source = await Presentation.load(await deck('LAYOUT_4x3'))
			assertEqual(await codeOf(() => run(target, source)), 'import/slide-size-mismatch', `${name} on a size difference`)
		}
	})

	test('a deck that declares no size is unknown, not a mismatch, everywhere', async () => {
		// This is the half that used to differ. A size that is not there cannot be compared and
		// cannot be rescaled onto, so no option the caller passes makes it answerable — which is
		// exactly what separates it from a mismatch.
		for (const [name, run] of ENTRY_POINTS) {
			const target = await Presentation.load(await deck())
			const source = await Presentation.load(await sizelessDeck())
			assertEqual(await codeOf(() => run(target, source)), 'import/slide-size-unknown', `${name} on a missing p:sldSz`)
		}
	})

	test('appendSlides makes the same two calls, from the other side', async () => {
		// Its source is a generator, which always declares a size, so the unknown half can only
		// ever be about the *target* package. Both are still the same two conditions.
		const target = await Presentation.load(await deck('LAYOUT_16x9'))
		const layout = target.layouts()[0].name
		assertEqual(
			await codeOf(() => target.appendSlides(generator('LAYOUT_4x3'), { layout })),
			'import/slide-size-mismatch',
			'appendSlides on a size difference'
		)

		const sizeless = await Presentation.load(await sizelessDeck())
		assertEqual(
			await codeOf(() => sizeless.appendSlides(generator(), { layout: sizeless.layouts()[0].name })),
			'import/slide-size-unknown',
			'appendSlides into a deck with no p:sldSz'
		)
	})

	test('a rescale answers a mismatch but not an unknown size', async () => {
		// The escape hatches are for the condition that *is* answerable.
		const target = await Presentation.load(await deck('LAYOUT_16x9'))
		const differing = await Presentation.load(await deck('LAYOUT_4x3'))
		assertEqual(
			await codeOf(() => target.importSlide(differing, 0, { rescale: 'fit' })),
			null,
			'importSlide rescales a known size difference'
		)

		const sizeless = await Presentation.load(await sizelessDeck())
		assertEqual(
			await codeOf(() => target.importSlide(sizeless, 0, { rescale: 'fit' })),
			'import/slide-size-unknown',
			'but a rescale cannot be computed from a size that is not there'
		)
		assertEqual(
			await codeOf(() => target.importSlideMasters(sizeless, { requireEqualSize: false })),
			null,
			'importSlideMasters { requireEqualSize: false } skips the check entirely, as documented'
		)
	})
})
