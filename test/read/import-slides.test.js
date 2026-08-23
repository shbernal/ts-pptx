// Batch slide import: `Presentation.importSlides(requests)` — the multi-page
// sibling of `importSlide`. Contract under test:
//   - every request lands at its `outputIndex` in the final slide list;
//   - validation happens up front, so a rejected batch changes no byte;
//   - a `slide → slide` link on an imported page resolves to another *selected*
//     page's fresh partname (never back into the source package), and a link to
//     an unselected page is refused rather than dragging that page across;
//   - results survive a save → reopen round-trip with no dangling relationships.

import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual, bytesEqual, throws } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { openFixture } from './corpus.js'
import { assertNoDanglingRels } from './opc.js'

const validatorInstalled = await validatorAvailable()

const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

/** The internal SLIDE_REL targets of one slide part, resolved via its own rels. */
function slideLinkTargets(opc, partName) {
	const rels = opc.relationshipsFor(partName)
	const out = []
	for (const rel of rels) {
		if (rel.type !== SLIDE_REL || rel.targetMode === 'External') continue
		out.push(rels.resolveTarget(rel.id))
	}
	return out
}

/** Catch a synchronous throw and return its stable `code`, or null when nothing threw. */
function catchCode(fn) {
	try {
		fn()
		return null
	} catch (err) {
		return err.code ?? null
	}
}

/** A generated two-page deck whose first page optionally links to the second. */
async function generatedDeck(firstLinksToSecond = false) {
	const pptx = new TsPptx()
	const first = pptx.addSlide()
	first.addText('first', { x: 1, y: 1, w: 4, h: 1 })
	if (firstLinksToSecond) first.addText('jump', { x: 1, y: 3, w: 4, h: 1, hyperlink: { slide: 2, tooltip: 'onward' } })
	pptx.addSlide().addText('second', { x: 1, y: 1, w: 4, h: 1 })
	return Presentation.load(await pptx.write({ outputType: 'uint8array' }))
}

describe('Presentation.importSlides', () => {
	test('places each selected page at its requested final position', async () => {
		const target = await openFixture('mixed') // 11 slides
		const before = target.slides.length
		const originalIds = new Set(target.slides.map((s) => s.slideId))
		const sourceA = await openFixture('mixed')
		const sourceB = await openFixture('mixed')

		target.importSlides([
			{ source: sourceA, sourceIndex: 0, outputIndex: 0 },
			{ source: sourceB, sourceIndex: 0, outputIndex: before + 1 },
			{ source: sourceB, sourceIndex: 1, outputIndex: before },
		])

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, before + 3, 'slide count grew by exactly the batch size')
		assertNoDanglingRels(reopened.opc)

		// The three imported pages mint fresh (highest) slide ids; they sit at
		// exactly the requested positions, originals keeping their relative order.
		const reopenedIds = reopened.slides.map((s) => s.slideId)
		const insertedAt = reopenedIds.flatMap((id, index) => (originalIds.has(id) ? [] : [index]))
		assertEqual(
			JSON.stringify(insertedAt),
			JSON.stringify([0, before, before + 1]),
			'imports land at their outputIndexes in ascending order'
		)
	})

	test('a rejected batch leaves the deck byte-identical — including a refused link', async () => {
		const target = await openFixture('mixed')
		const linked = await generatedDeck(true)

		// Page 0 links to page 1, which is not selected: refused up front.
		const beforeBytes = await target.save()
		assert(
			throws(() => target.importSlides([{ source: linked, sourceIndex: 0, outputIndex: 0 }])),
			'a link to an unselected page is refused'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'the refused batch changed no byte of the deck')

		// Duplicate final positions are likewise caught before anything is copied.
		assert(
			throws(() =>
				target.importSlides([
					{ source: plain, sourceIndex: 0, outputIndex: 2 },
					{ source: plain, sourceIndex: 1, outputIndex: 2 },
				])
			),
			'duplicate output positions are rejected'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'and again, no byte changed')
	})

	test('selecting one source page twice is rejected', () => {
		return openFixture('mixed').then(async (target) => {
			const source = await openFixture('image')
			assertEqual(
				catchCode(() =>
					target.importSlides([
						{ source, sourceIndex: 0, outputIndex: 0 },
						{ source, sourceIndex: 0, outputIndex: 1 },
					])
				),
				'import/slide-selected-twice',
				'a source page may appear in one request only'
			)
		})
	})

	test('an out-of-range or negative outputIndex is rejected', () => {
		return openFixture('mixed').then(async (target) => {
			const count = target.slides.length
			const source = await openFixture('image')
			assertEqual(
				catchCode(() => target.importSlides([{ source, sourceIndex: 0, outputIndex: count + 5 }])),
				'import/output-index-out-of-range',
				'an outputIndex past the final list is rejected'
			)
			assertEqual(
				catchCode(() => target.importSlides([{ source, sourceIndex: 0, outputIndex: -1 }])),
				'import/output-index-out-of-range',
				'a negative outputIndex is rejected'
			)
		})
	})

	test('a link between two selected pages is rewritten to the imported parts', async () => {
		const targetDeck = new TsPptx()
		targetDeck.addSlide().addText('a', { x: 1, y: 1, w: 2, h: 1 })
		targetDeck.addSlide().addText('b', { x: 1, y: 1, w: 2, h: 1 })
		const target = await Presentation.load(await targetDeck.write({ outputType: 'uint8array' }))
		const source = await generatedDeck(true)

		target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 0 },
			{ source, sourceIndex: 1, outputIndex: 1 },
		])

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)

		// Both pages came across under fresh partnames; the first page's slide
		// link must resolve to the SECOND IMPORTED partname, not into the source.
		const [importedFirst, importedSecond] = reopened.slides.map((s) => s.partName)
		const targets = slideLinkTargets(reopened.opc, importedFirst)
		assertEqual(targets.length, 1, 'the generated jump link survived the import')
		assertEqual(targets[0], importedSecond, 'the link resolves to the second imported page')
	})

	test.skipIf(!validatorInstalled)('a batch-imported deck stays schema-valid', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 3 },
			{ source, sourceIndex: 1, outputIndex: 9 },
		])
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
