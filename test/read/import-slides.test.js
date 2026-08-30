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
//   - `{ importNotes: true }` carries a page's notes across per request, under the
//     one-notesMaster rule, and the dry run covers that graph too;
//   - `{ embedFonts: true }` carries the *source deck's* whole face list once, and
//     `{ rescale }` puts a differently-sized source on this canvas -- both are
//     deck-level decisions in a per-page spelling, so the batch reconciles them up
//     front and refuses a source whose requests disagree about the rescale;
//   - results survive a save → reopen round-trip with no dangling relationships.

import JSZip from 'jszip'
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
const NOTES_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
const TAGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'
const TAGS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml'
const FONT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'

/** Every `notesMaster` a deck registers in presentation.xml, resolved to partnames. */
function notesMasters(pres) {
	const rels = pres.opc.relationshipsFor(pres.presentationPart.partName)
	return rels.byType(NOTES_MASTER_REL).map((rel) => rels.resolveTarget(rel.id))
}

/** The notesMaster part a notes slide binds to, or null when it names none. */
function notesMasterOf(pres, notesPartName) {
	const rels = pres.opc.relationshipsFor(notesPartName)
	const rel = rels.byType(NOTES_MASTER_REL)[0]
	return rel ? rels.resolveTarget(rel.id) : null
}

/** The `tags` targets of one part, resolved -- the owned part the notes copies must not share. */
function tagTargets(opc, partName) {
	const rels = opc.relationshipsFor(partName)
	return rels.byType(TAGS_REL).map((rel) => rels.resolveTarget(rel.id))
}

/**
 * `notes-slide-image`, with one part hung off its notes slide that page copies
 * must not share -- a `tags` part, which `page-owned.js` classes as owned. The
 * fixture's own notes reference only their master and their page, so without this
 * there is nothing under the notes for a second copy to collide on.
 */
async function sourceWithOwnedNotesPart() {
	const source = await openFixture('notes-slide-image')
	const notesPartName = source.slides[0].notesSlide.partName
	source.opc.addPart(
		'/ppt/tags/tag1.xml',
		TAGS_CONTENT_TYPE,
		new TextEncoder().encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
				'<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
				'<p:tag name="OWNED" val="1"/></p:tagLst>'
		)
	)
	source.opc.relationshipsFor(notesPartName).add(TAGS_REL, '../tags/tag1.xml')
	return source
}

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

/** A two-page 4x3 deck: a source on a canvas the 16x9 destinations do not share. */
async function otherCanvasDeck() {
	const wide = new TsPptx()
	wide.layout = 'LAYOUT_4x3'
	wide.addSlide().addText('other canvas', { x: 1, y: 1, w: 4, h: 1 })
	wide.addSlide().addText('and another', { x: 2, y: 2, w: 4, h: 1 })
	return Presentation.load(await wide.write({ outputType: 'uint8array' }))
}

/** Every `a:off` on a page, as `x,y` strings, so a rescale shows up as a list that moved. */
function offsetsOf(slide) {
	const xml = new TextDecoder().decode(slide.part.bytes)
	return [...xml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)].map((m) => `${m[1]},${m[2]}`)
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

	test('a source with a different slide size is rejected unless the request rescales', async () => {
		const target = await generatedDeck(false)
		const source = await otherCanvasDeck()
		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() => target.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])),
			'import/slide-size-mismatch',
			'a size difference is fatal without the rescale spelling'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'and the deck is untouched')
	})

	test("{ rescale } scales an imported page onto this deck's canvas", async () => {
		// The 4x3 source is 9144000 EMU wide against the destination's 12192000, so a
		// 'fit' rescale has to move every offset it finds on the page.
		const target = await generatedDeck(false)
		const source = await otherCanvasDeck()
		const before = offsetsOf(source.slides[0])
		target.importSlides([{ source, sourceIndex: 0, outputIndex: 0, rescale: 'fit' }])
		// Through a save: the rescale edits the part's DOM, and `part.bytes` is the
		// original until the package is serialized.
		const reopened = await Presentation.load(await target.save())
		const after = offsetsOf(reopened.slides[0])
		assertEqual(after.length, before.length, 'the same shapes came across')
		assert(
			after.some((off, i) => off !== before[i]),
			`expected the geometry to move; got ${JSON.stringify(after)} against ${JSON.stringify(before)}`
		)
		assertNoDanglingRels(reopened.opc)

		// And without the option the same import is refused, so the movement above is
		// the rescale rather than something the copy would have done anyway.
		const plain = await generatedDeck(false)
		assertEqual(
			catchCode(() => plain.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])),
			'import/slide-size-mismatch',
			'the same batch without { rescale } is the rejected one'
		)
	})

	test('requests from one source must agree on rescale, and disagreement changes no byte', async () => {
		// A batch import is `'copy'` themed, so a rescale moves the shared imported
		// layout and master too: rescaling one of a source's pages and not another
		// would leave the second aligned against a master that had shifted under it.
		const target = await generatedDeck(false)
		const source = await otherCanvasDeck()
		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() =>
				target.importSlides([
					{ source, sourceIndex: 0, outputIndex: 0, rescale: 'fit' },
					{ source, sourceIndex: 1, outputIndex: 1, rescale: 'stretch' },
				])
			),
			'import/rescale-conflict',
			'two rescale modes for one source is a conflict, not a pick'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'and the deck is untouched')
		// `true` is the documented spelling of `'fit'`, so the two agree.
		assertEqual(
			catchCode(() =>
				target.importSlides([
					{ source, sourceIndex: 0, outputIndex: 0, rescale: true },
					{ source, sourceIndex: 1, outputIndex: 1, rescale: 'fit' },
				])
			),
			null,
			'`true` and `fit` are one answer'
		)
	})

	test("{ embedFonts } carries a source deck's faces once, however many of its pages come over", async () => {
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 0, embedFonts: true },
			// The second request does not ask; the carry is a whole-deck operation, so
			// asking once is asking for the source's whole list, exactly once.
			{ source, sourceIndex: 0, outputIndex: 1 },
		])

		const zip = await JSZip.loadAsync(await target.save())
		const fontParts = Object.keys(zip.files).filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n))
		assertEqual(fontParts.length, 2, `the source's two faces carried once each (got ${JSON.stringify(fontParts)})`)
		const ct = await zip.file('[Content_Types].xml').async('string')
		assertEqual((ct.match(/x-fontdata/g) || []).length, 1, 'content type registered once (Default only)')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(/<p:font typeface="Silkscreen"/.test(pres), `embeddedFontLst merged; got ${pres}`)
		assertEqual((pres.match(/<p:embeddedFont>/g) || []).length, 1, 'one entry for the one typeface')
		assertNoDanglingRels((await Presentation.load(await target.save())).opc)
	})

	test('no embedFonts request leaves the deck without fonts, as before', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		target.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])
		const zip = await JSZip.loadAsync(await target.save())
		assert(!Object.keys(zip.files).some((n) => /fntdata/.test(n)), 'no font parts without embedFonts')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(!/embeddedFontLst/.test(pres), 'no embeddedFontLst without embedFonts')
	})

	test('a font carry that could not complete is refused with the deck byte-identical', async () => {
		// The carry runs after the pages are copied, so without a dry run of its own a
		// missing binary would throw with parts already added and no way back.
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		const presRels = source.opc.relationshipsFor(source.presentationPart.partName)
		const fontPart = presRels.byType(FONT_REL).map((rel) => presRels.resolveTarget(rel.id))[0]
		assert(fontPart, 'the fixture embeds at least one face')
		source.opc.removePart(fontPart)
		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() => target.importSlides([{ source, sourceIndex: 0, outputIndex: 0, embedFonts: true }])),
			'package/part-missing',
			'the font dry run rejects the batch'
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

	test('speaker notes travel per request, and only where asked', async () => {
		// The default is `importSlide`'s: the page copy drops the notesSlide rel and
		// the imported page arrives without notes. One request opting in must not
		// carry the other's, which is the point of the flag being per request.
		const target = await openFixture('textbox') // 2 slides, no notesMaster of its own
		const before = target.slides.length
		const withNotes = await openFixture('notes-slide-image')
		const withoutNotes = await openFixture('notes-slide-image')

		target.importSlides([
			{ source: withNotes, sourceIndex: 0, outputIndex: before, importNotes: true },
			{ source: withoutNotes, sourceIndex: 0, outputIndex: before + 1 },
		])

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		assertEqual(
			reopened.slides[before].notesText,
			'Speaker notes so PowerPoint emits the notes slide.',
			'the opted-in page kept the source notes'
		)
		assertEqual(reopened.slides[before + 1].notesText, null, 'the request that did not ask still gets no notes')
		assertEqual(notesMasters(reopened).length, 1, 'carrying notes into a deck with none installs exactly one master')
	})

	test('a destination that has a notes master keeps it, and the source master is not copied', async () => {
		// p:notesMasterIdLst is 0..1, so the destination's notes styling wins -- the
		// same rule importSlide({ importNotes: true }) and appendSlides follow, which
		// is what lets the three be mixed on one deck.
		const target = await openFixture('read-stress') // already carries a notesMaster
		const [ownMaster] = notesMasters(target)
		assert(ownMaster !== undefined, 'the fixture registers a notesMaster to defend')
		const masterParts = (pres) => [...pres.opc.parts.keys()].filter((name) => name.includes('/notesMasters/')).length
		const mastersBefore = masterParts(target)
		const source = await openFixture('notes-slide-image')
		const at = target.slides.length

		target.importSlides([{ source, sourceIndex: 0, outputIndex: at, importNotes: true }])

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		assertEqual(JSON.stringify(notesMasters(reopened)), JSON.stringify([ownMaster]), 'the deck kept its own master')
		assertEqual(masterParts(reopened), mastersBefore, 'and no second notesMaster part came across')
		assertEqual(
			notesMasterOf(reopened, reopened.slides[at].notesSlide.partName),
			ownMaster,
			'the carried notes bind to the destination master'
		)
	})

	test('a page named twice with notes gets notes -- and the parts under them -- of its own', async () => {
		// A notes slide is a part its page owns, so two copies of one page may not
		// resolve to one notes part, and neither may the parts those notes own.
		const target = await openFixture('textbox')
		const before = target.slides.length
		const source = await sourceWithOwnedNotesPart()

		target.importSlides([
			{ source, sourceIndex: 0, outputIndex: before, importNotes: true },
			{ source, sourceIndex: 0, outputIndex: before + 1, importNotes: true },
		])

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		const first = reopened.slides[before]
		const second = reopened.slides[before + 1]
		assert(first.notesSlide !== null && second.notesSlide !== null, 'both copies came across with notes')
		assert(first.notesSlide.partName !== second.notesSlide.partName, 'each copy has a notes part of its own')
		assertEqual(first.notesText, second.notesText, 'and both say what the source page said')
		const firstTags = tagTargets(reopened.opc, first.notesSlide.partName)
		const secondTags = tagTargets(reopened.opc, second.notesSlide.partName)
		assertEqual(firstTags.length, 1, 'the owned part under the notes came across')
		assertEqual(secondTags.length, 1, 'for the second copy too')
		assert(firstTags[0] !== secondTags[0], 'and the second copy took its own rather than sharing the first')
	})

	test('a batch that could not carry the notes is refused with the deck byte-identical', async () => {
		// The dry run is what makes a rejected batch a no-op, and carryNotes runs after
		// the copy -- so the notes graph has to be part of the dry run or the guarantee
		// stops at the notes rel. Snapping the source's notesMaster is the cheapest way
		// to reach that: the destination has none, so it would be copied.
		const target = await openFixture('textbox')
		const broken = await openFixture('notes-slide-image')
		broken.opc.removePart(notesMasterOf(broken, broken.slides[0].notesSlide.partName))

		const beforeBytes = await target.save()
		assertEqual(
			catchCode(() => target.importSlides([{ source: broken, sourceIndex: 0, outputIndex: 0, importNotes: true }])),
			'package/part-missing',
			'a part only the notes copy would have reached is missing from the source'
		)
		assert(bytesEqual(beforeBytes, await target.save()), 'the refused batch changed no byte')

		// Without the opt-in the notes graph is not the batch's business at all, so the
		// very same damaged source imports cleanly.
		target.importSlides([{ source: broken, sourceIndex: 0, outputIndex: 0 }])
		assertEqual(target.slides.length, 3, 'the same source imports fine when its notes are not asked for')
		assertEqual(target.slides[0].notesText, null, 'and the page arrives without notes, as ever')
	})

	test('a destination master spares the source master the dry run would otherwise reject', async () => {
		// The mirror of the case above: with a notesMaster of its own the deck never
		// reads the source's, so a dry run that walked it regardless would reject a
		// batch the copy would have completed.
		const target = await openFixture('read-stress') // has a notesMaster
		const broken = await openFixture('notes-slide-image')
		broken.opc.removePart(notesMasterOf(broken, broken.slides[0].notesSlide.partName))
		const at = target.slides.length

		target.importSlides([{ source: broken, sourceIndex: 0, outputIndex: at, importNotes: true }])

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		assertEqual(
			reopened.slides[at].notesText,
			'Speaker notes so PowerPoint emits the notes slide.',
			'the notes bound to the destination master'
		)
	})

	test.skipIf(!validatorInstalled)('a batch that carried notes stays schema-valid', async () => {
		const target = await openFixture('textbox')
		const source = await openFixture('notes-slide-image')
		target.importSlides([
			{ source, sourceIndex: 0, outputIndex: 0, importNotes: true },
			{ source, sourceIndex: 0, outputIndex: 1, importNotes: true },
		])
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
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
