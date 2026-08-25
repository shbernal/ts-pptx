// Batch slide import: `Presentation.importSlides(requests)` — the multi-page
// sibling of `importSlide`. Contract under test:
//   - every request lands at its `outputIndex` in the final slide list, and the
//     returned array stays parallel to `requests`;
//   - validation happens up front — including a dry run of the copy itself — so
//     a rejected batch changes no byte, whichever rule did the rejecting;
//   - a `slide → slide` link on an imported page resolves to another *selected*
//     page's fresh partname (never back into the source package), and a link to
//     an unselected page is refused rather than dragging that page across;
//   - one request is one output page, so naming a source page twice returns two
//     independent pages over one shared subgraph, and a jump link out of either
//     copy still lands inside the batch;
//   - results survive a save → reopen round-trip with no dangling relationships.

import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual, bytesEqual } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { openFixture } from './corpus.js'
import { assertNoDanglingRels } from './opc.js'

const validatorInstalled = await validatorAvailable()

const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

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

/** Every internal relationship target of one part, sorted, for comparing two pages' dependencies. */
function depTargets(opc, partName) {
	const rels = opc.relationshipsFor(partName)
	const out = []
	for (const rel of rels) {
		if (rel.targetMode === 'External') continue
		out.push(rels.resolveTarget(rel.id))
	}
	return out.sort()
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

/**
 * A generated two-page deck whose first page optionally links to the second and
 * always carries an external hyperlink, so the copy has a rel of each target mode
 * to route.
 */
async function generatedDeck(firstLinksToSecond = false) {
	const pptx = new TsPptx()
	const first = pptx.addSlide()
	first.addText('first', { x: 1, y: 1, w: 4, h: 1 })
	first.addText('outward', { x: 1, y: 2, w: 4, h: 1, hyperlink: { url: 'https://example.invalid/' } })
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
		// Generated decks share one layout, so sizes match by construction and a
		// rejection below can only come from the rule under test.
		const target = await generatedDeck(false)
		const linked = await generatedDeck(true)
		const plain = await generatedDeck(false)

		// Page 0 links to page 1, which is not selected: refused up front.
		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() => target.importSlides([{ source: linked, sourceIndex: 0, outputIndex: 0 }])),
			'import/unresolved-slide-link',
			'a link to an unselected page is refused'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'the refused batch changed no byte of the deck')

		// Duplicate final positions are likewise caught before anything is copied.
		assertEqual(
			catchCode(() =>
				target.importSlides([
					{ source: plain, sourceIndex: 0, outputIndex: 2 },
					{ source: plain, sourceIndex: 1, outputIndex: 2 },
				])
			),
			'import/output-index-conflict',
			'duplicate output positions are rejected by the conflict rule, not by some other one'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'and again, no byte changed')
	})

	test('a source whose dependency graph is broken is refused before anything is copied', async () => {
		// The copy phase used to be the last place a batch could fail, and it failed
		// with parts already added and the copied master already registered in
		// presentation.xml. A dry run of the traversal now runs first, so a damaged
		// source is a validation error like any other.
		const target = await generatedDeck(false)
		const good = await generatedDeck(false)
		const broken = await generatedDeck(false)
		const rels = broken.opc.relationshipsFor(broken.slides[0].partName)
		const layoutRel = [...rels].find((rel) => rel.type === SLIDE_LAYOUT_REL)
		assert(layoutRel !== undefined, 'the generated page has a layout to snap')
		broken.opc.removePart(rels.resolveTarget(layoutRel.id))

		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() =>
				target.importSlides([
					{ source: good, sourceIndex: 0, outputIndex: 0 },
					{ source: broken, sourceIndex: 0, outputIndex: 1 },
				])
			),
			'package/part-missing',
			'a part the copy would have reached is missing from the source'
		)
		// The good request sits before the broken one, so a batch that copied as it
		// went would have left its page, layout, master and theme behind here.
		assert(bytesEqual(beforeBytes, await target.save()), 'the refused batch changed no byte')
		assertEqual(target.slides.length, 2, 'and added no slide')
	})

	test('the returned array is parallel to the requests, not to the output order', async () => {
		const target = await generatedDeck(false)
		const source = await generatedDeck(false)

		// Request 0 asks for the last position, request 1 for the first: sorting by
		// outputIndex to insert must not reorder what the caller gets back.
		const [first, second] = target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 3 },
			{ source, sourceIndex: 1, outputIndex: 0 },
		])
		assertEqual(first.index, 3, 'requests[0] landed at its outputIndex 3')
		assertEqual(second.index, 0, 'requests[1] landed at its outputIndex 0')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides[3].partName, first.partName, 'and the deck agrees about the last page')
		assertEqual(reopened.slides[0].partName, second.partName, 'and about the first')
	})

	test('a source with a different slide size is rejected', async () => {
		const target = await generatedDeck(false)
		const wide = new TsPptx()
		wide.layout = 'LAYOUT_4x3'
		wide.addSlide().addText('other canvas', { x: 1, y: 1, w: 4, h: 1 })
		const source = await Presentation.load(await wide.write({ outputType: 'uint8array' }))
		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() => target.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])),
			'import/slide-size-mismatch',
			'importSlides has no rescale escape hatch, so a size difference is fatal'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'and the deck is untouched')
	})

	test('one source page requested twice yields two independent pages', async () => {
		// One request is one output page. The page part is the one thing an import
		// never shares, so the two copies must be distinct parts with distinct slide
		// ids -- while everything under them (layout, master, theme, media) is copied
		// once, exactly as a pair of `importSlide` calls would.
		const target = await openFixture('mixed')
		const before = target.slides.length
		const source = await openFixture('mixed')

		const [first, second] = target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 0 },
			{ source, sourceIndex: 0, outputIndex: before + 1 },
		])
		assert(first.partName !== second.partName, 'the two requests got parts of their own')
		assert(first.slideId !== second.slideId, 'and slide ids of their own')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, before + 2, 'both copies joined the deck')
		assertNoDanglingRels(reopened.opc)
		assertEqual(reopened.slides[0].partName, first.partName, 'the first copy landed at outputIndex 0')
		assertEqual(reopened.slides[before + 1].partName, second.partName, 'the second at the end')

		// Same bytes, same dependencies: the duplicate is a second page, not a
		// second copy of the subgraph underneath it.
		assert(
			bytesEqual(reopened.opc.part(first.partName).bytes, reopened.opc.part(second.partName).bytes),
			'the two pages are byte-identical copies of the one source page'
		)
		assertEqual(
			JSON.stringify(depTargets(reopened.opc, first.partName)),
			JSON.stringify(depTargets(reopened.opc, second.partName)),
			'and they share every part they depend on'
		)
	})

	test('a page duplicated beside a linked page keeps each copy linked within the batch', async () => {
		// Page 0 links to page 1. Asking for page 0 twice and page 1 once must leave
		// both copies of page 0 pointing at the single imported page 1 -- never back
		// into the source package, and without a third page appearing.
		const target = await generatedDeck(false)
		const source = await generatedDeck(true)
		const before = target.slides.length

		const [linkOwner, linkTarget, secondOwner] = target.importSlides([
			{ source, sourceIndex: 0, outputIndex: before },
			{ source, sourceIndex: 1, outputIndex: before + 1 },
			{ source, sourceIndex: 0, outputIndex: before + 2 },
		])
		assert(linkOwner.partName !== secondOwner.partName, 'the repeated page got two parts')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, before + 3, 'the batch added exactly its three pages')
		assertNoDanglingRels(reopened.opc)
		for (const owner of [linkOwner, secondOwner]) {
			assertEqual(
				JSON.stringify(slideLinkTargets(reopened.opc, owner.partName)),
				JSON.stringify([linkTarget.partName]),
				'each copy of the linking page resolves to the imported link target'
			)
		}
	})

	test('two pages duplicated together link to their own round-mate', async () => {
		// Both pages of a linked pair, each asked for twice: the copies are made in
		// rounds, so the batch produces two self-contained pairs rather than three
		// pages pointing at one.
		const target = await generatedDeck(false)
		const source = await generatedDeck(true)
		const before = target.slides.length

		const [ownerA, targetA, ownerB, targetB] = target.importSlides([
			{ source, sourceIndex: 0, outputIndex: before },
			{ source, sourceIndex: 1, outputIndex: before + 1 },
			{ source, sourceIndex: 0, outputIndex: before + 2 },
			{ source, sourceIndex: 1, outputIndex: before + 3 },
		])

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, before + 4, 'four pages joined the deck')
		assertNoDanglingRels(reopened.opc)
		assertEqual(
			JSON.stringify(slideLinkTargets(reopened.opc, ownerA.partName)),
			JSON.stringify([targetA.partName]),
			'the first pair links inside itself'
		)
		assertEqual(
			JSON.stringify(slideLinkTargets(reopened.opc, ownerB.partName)),
			JSON.stringify([targetB.partName]),
			'and the second pair inside itself'
		)
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

	test('a link into a page an earlier batch already imported resolves to that copy', async () => {
		// The rule is not "selected in this batch" but "already in this deck from
		// this source": importing the link target first has to satisfy it, and the
		// registry must hand back the earlier copy rather than a second one.
		const targetDeck = new TsPptx()
		targetDeck.addSlide().addText('own', { x: 1, y: 1, w: 2, h: 1 })
		const target = await Presentation.load(await targetDeck.write({ outputType: 'uint8array' }))
		const source = await generatedDeck(true)

		const [linkTarget] = target.importSlides([{ source, sourceIndex: 1, outputIndex: 1 }])
		const [linkOwner] = target.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, 3, 'the second batch added exactly its own page')
		assertNoDanglingRels(reopened.opc)
		assertEqual(
			JSON.stringify(slideLinkTargets(reopened.opc, linkOwner.partName)),
			JSON.stringify([linkTarget.partName]),
			'the jump link points at the page the first batch brought across'
		)
	})

	test.skipIf(!validatorInstalled)('a batch-imported deck stays schema-valid', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 3 },
			{ source, sourceIndex: 1, outputIndex: 9 },
			{ source, sourceIndex: 0, outputIndex: 0 }, // the same page a second time
		])
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
