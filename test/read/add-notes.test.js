// Authoring speaker notes onto a loaded deck (`Slide.addNotes`).
//
// Contract under test: the read model can give a slide notes it does not have. That is
// the third way a notes slide comes into being — the other two only *move* an existing
// part (`importSlide({ importNotes: true })` across decks, `appendSlides` from the
// generator) — and it is the only one reachable for a slide whose notes part is simply
// absent, the state an import without `importNotes` leaves behind.
//
// Two things are worth more than the round trip. First, the single-notesMaster rule: a
// presentation holds at most one, so authoring notes onto several slides of a deck that
// had none must install exactly one master and bind every notes part to it. Second, the
// part `addNotes` builds must stay the part the *write* path builds — it reuses the
// generator's frame (`makeXmlNotesSlideSkeleton`) precisely so the two cannot drift, and
// the equivalence test below is what keeps that true.

import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { openFixture } from './corpus.js'
import { assertNoDanglingRels, resolveSingle } from './opc.js'

const validatorInstalled = await validatorAvailable()

const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const NOTES_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

// `empty.pptx` is a real PowerPoint deck with one slide, no notes part, and no notes
// master — the only fixture that exercises the branch which has to *install* one. A
// generator deck cannot: the write path ships a notesMaster and a notes part for every
// slide whether or not `addNotes` was called, so it always takes the reuse branch.

/** A one-slide generator deck, loaded through the read model. */
async function authoredDeck(build) {
	const pptx = new TsPptx()
	const slide = pptx.addSlide()
	slide.addText('body', { x: 1, y: 1, w: 4, h: 1 })
	if (build) build(slide)
	return Presentation.load(new Uint8Array(await pptx.toBytes()))
}

/** notesMaster partnames registered in presentation.xml's p:notesMasterIdLst. */
function registeredNotesMasters(opc) {
	const rootRels = opc.relationshipsFor('/')
	const officeDoc = [...rootRels].find((rel) => rel.type === OFFICE_DOCUMENT_REL)
	const presName = rootRels.resolveTarget(officeDoc.id)
	const root = opc.part(presName).dom.documentElement
	const rels = opc.relationshipsFor(presName)
	const out = []
	for (let n = root.firstChild; n; n = n.nextSibling) {
		if (n.nodeType !== 1 || n.localName !== 'notesMasterIdLst') continue
		for (let e = n.firstChild; e; e = e.nextSibling) {
			if (e.nodeType !== 1 || e.localName !== 'notesMasterId') continue
			const relId = e.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
			out.push(rels.resolveTarget(relId))
		}
	}
	return out
}

/** The body placeholder's `p:txBody` XML of a slide's notes part. */
function notesBodyXml(pres, slideIndex) {
	const xml = new TextDecoder().decode(pres.slides[slideIndex].notesSlide.part.bytes)
	// The three placeholders in document order are sldImg (no txBody), body, sldNum.
	const bodies = xml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/g)
	return bodies ? bodies[0] : null
}

describe('Slide.addNotes on a loaded deck', () => {
	test('gives a slide with no notes part one, and it survives a round trip', async () => {
		// An imported slide with `importNotes` off is the real-world no-notes slide.
		const source = await authoredDeck()
		const target = await authoredDeck()
		const imported = target.importSlide(source, 0, { theme: 'copy' })
		assertEqual(imported.notesSlide, null, 'imported slide starts with no notes part')

		const notes = imported.addNotes('what changed, and why')
		assertEqual(notes.text, 'what changed, and why', 'addNotes returns the modeled notes slide')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides[1].notesText, 'what changed, and why', 'notes survive save → reload')
		assertNoDanglingRels(reopened.opc)
	})

	test('splits on \\n into paragraphs, matching the write-side addNotes', async () => {
		const written = await authoredDeck((slide) => slide.addNotes('line one\nline two'))
		const read = await authoredDeck()
		read.slides[0].addNotes('line one\nline two')
		const reopened = await Presentation.load(await read.save())

		assertEqual(reopened.slides[0].notesText, 'line one\nline two', 'flattened text matches')
		assertEqual(
			reopened.slides[0].notesTextFrame.paragraphs.length,
			written.slides[0].notesTextFrame.paragraphs.length,
			'same paragraph count as the write path'
		)
		assertEqual(reopened.slides[0].notesTextFrame.paragraphs.length, 2, 'two paragraphs')
	})

	// The anti-drift guard the module doc promises. The read model reserializes every
	// part it authors through `read/oxml/dom.ts`, which self-closes empty elements,
	// where the write path spells `a:rPr` open/close — so the comparison normalizes
	// that spelling and nothing else. Any *other* difference is real drift.
	test('authors the same notes body the write path does, modulo empty-element spelling', async () => {
		const written = await authoredDeck((slide) => slide.addNotes('line one\nline two'))
		const read = await authoredDeck()
		read.slides[0].addNotes('line one\nline two')
		const reopened = await Presentation.load(await read.save())

		const collapseEmptyTags = (xml) => xml.replace(/<([a-z0-9:]+)([^>]*?)><\/\1>/gi, '<$1$2/>')
		assertEqual(
			collapseEmptyTags(notesBodyXml(reopened, 0)),
			collapseEmptyTags(notesBodyXml(written, 0)),
			'read-authored notes body matches the generator’s'
		)
	})

	test('wires notesMaster as rId1 and the slide back-rel as rId2, as the write path reserves', async () => {
		const source = await authoredDeck()
		const target = await authoredDeck()
		const imported = target.importSlide(source, 0, { theme: 'copy' })
		imported.addNotes('note')

		const reopened = await Presentation.load(await target.save())
		const slideName = reopened.slides[1].partName
		const notesName = resolveSingle(reopened.opc, slideName, NOTES_SLIDE_REL)
		assert(notesName, 'slide gained a notesSlide rel')

		const rels = [...reopened.opc.relationshipsFor(notesName)]
		const byId = Object.fromEntries(rels.map((rel) => [rel.id, rel.type]))
		assertEqual(byId.rId1, NOTES_MASTER_REL, 'rId1 is the notes master')
		assertEqual(byId.rId2, SLIDE_REL, 'rId2 is the slide back-reference')
		assertEqual(
			reopened.opc.relationshipsFor(notesName).resolveTarget('rId2'),
			slideName,
			'the back-reference points at the annotated slide, not the source'
		)
	})

	test('annotates a deck that has no notes part and no notes master at all', async () => {
		const deck = await openFixture('empty')
		assertEqual(deck.slides[0].notesSlide, null, 'fixture slide has no notes part')
		assertEqual(registeredNotesMasters(deck.opc).length, 0, 'and the deck has no notes master')

		deck.slides[0].addNotes('notes on a deck that never had any')

		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides[0].notesText, 'notes on a deck that never had any')
		assertEqual(registeredNotesMasters(reopened.opc).length, 1, 'a notes master was installed')
		assertNoDanglingRels(reopened.opc)
	})

	test('installs exactly one notes master however many slides gain notes', async () => {
		const source = await authoredDeck()
		const target = await openFixture('empty')
		// Two imported slides, neither carrying notes, both annotated afterwards.
		// `rescale` only reconciles the fixture's 16:9 canvas with the generator's 4:3.
		const a = target.importSlide(source, 0, { theme: 'copy', rescale: 'fit' })
		const b = target.importSlide(source, 0, { theme: 'copy', rescale: 'fit' })
		assertEqual(registeredNotesMasters(target.opc).length, 0, 'no notes master before')
		a.addNotes('first')
		b.addNotes('second')

		const reopened = await Presentation.load(await target.save())
		const masters = registeredNotesMasters(reopened.opc)
		assertEqual(masters.length, 1, 'exactly one notesMaster registered')

		const notesA = resolveSingle(reopened.opc, reopened.slides[a.index].partName, NOTES_SLIDE_REL)
		const notesB = resolveSingle(reopened.opc, reopened.slides[b.index].partName, NOTES_SLIDE_REL)
		assert(notesA !== notesB, 'each slide got its own notes part')
		assertEqual(
			resolveSingle(reopened.opc, notesA, NOTES_MASTER_REL),
			resolveSingle(reopened.opc, notesB, NOTES_MASTER_REL),
			'both notes parts bind to the same master'
		)
		assertEqual(resolveSingle(reopened.opc, notesA, NOTES_MASTER_REL), masters[0], 'and it is the registered one')
		assertEqual(reopened.slides[a.index].notesText, 'first')
		assertEqual(reopened.slides[b.index].notesText, 'second')
		assertNoDanglingRels(reopened.opc)
	})

	test('reuses a notes master the deck already has', async () => {
		// A deck whose slide 0 has notes already carries a notesMaster.
		const deck = await authoredDeck((slide) => slide.addNotes('existing'))
		const before = registeredNotesMasters(deck.opc)
		assertEqual(before.length, 1, 'fixture deck has one notesMaster')

		const source = await authoredDeck()
		const imported = deck.importSlide(source, 0, { theme: 'copy' })
		imported.addNotes('added later')

		const reopened = await Presentation.load(await deck.save())
		const after = registeredNotesMasters(reopened.opc)
		assertEqual(after.length, 1, 'still exactly one notesMaster')
		assertEqual(after[0], before[0], 'and it is the deck’s own, not a new one')
	})

	test('the installed notes master binds to its own theme part, not the slide master’s', async () => {
		const target = await openFixture('empty')
		target.slides[0].addNotes('note')

		const reopened = await Presentation.load(await target.save())
		const master = registeredNotesMasters(reopened.opc)[0]
		const notesTheme = resolveSingle(reopened.opc, master, THEME_REL)
		assert(notesTheme, 'the notes master resolves a theme')

		// Every slide master's theme must be a different part: sharing one would make
		// the package depend on a part two masters claim.
		for (const slideMaster of reopened.masters()) {
			const slideTheme = resolveSingle(reopened.opc, slideMaster.partName, THEME_REL)
			assert(slideTheme !== notesTheme, `notes theme ${notesTheme} is not slide master theme ${slideTheme}`)
		}
	})

	test('replaces the body of a slide that already has notes, keeping the same part', async () => {
		const deck = await authoredDeck((slide) => slide.addNotes('original'))
		const partName = deck.slides[0].notesSlide.part.partName
		deck.slides[0].addNotes('replaced\nover two lines')

		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides[0].notesSlide.part.partName, partName, 'no second notes part was created')
		assertEqual(reopened.slides[0].notesText, 'replaced\nover two lines', 'body was replaced')
		assertEqual(reopened.slides[0].notesTextFrame.paragraphs.length, 2, 'and re-split into paragraphs')
		// The other two placeholders are untouched by a body rewrite.
		assert(reopened.slides[0].notesSlide.slideImage, 'sldImg placeholder survives')
		assert(reopened.slides[0].notesSlide.slideNumber, 'sldNum placeholder survives')
	})

	test('empty text is a note with one empty paragraph, not a missing part', async () => {
		const source = await authoredDeck()
		const target = await authoredDeck()
		const imported = target.importSlide(source, 0, { theme: 'copy' })
		imported.addNotes('')

		const reopened = await Presentation.load(await target.save())
		assert(reopened.slides[1].notesSlide, 'the part exists')
		assertEqual(reopened.slides[1].notesText, '', 'and reads back as empty, not null')
	})

	test('escapes XML metacharacters in the note', async () => {
		const source = await authoredDeck()
		const target = await authoredDeck()
		const imported = target.importSlide(source, 0, { theme: 'copy' })
		const raw = 'a < b & c > d "quoted"'
		imported.addNotes(raw)

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides[1].notesText, raw, 'metacharacters round-trip as text, not markup')
	})

	test('a notes part whose target is missing is reported, not silently replaced', async () => {
		const deck = await authoredDeck((slide) => slide.addNotes('original'))
		const partName = deck.slides[0].notesSlide.part.partName
		deck.opc.removePart(partName)

		let code = null
		try {
			deck.slides[0].addNotes('replacement')
		} catch (err) {
			code = err.code
		}
		assertEqual(code, 'package/part-missing', 'a dangling notes rel surfaces rather than being papered over')
	})

	test.skipIf(!validatorInstalled)('the resulting package is schema-valid', async () => {
		// The install branch, on a real PowerPoint deck: a notes part, a notes master,
		// and a cloned theme all newly authored into a package that had none of them.
		const target = await openFixture('empty')
		target.slides[0].addNotes('validated notes\nsecond paragraph')
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
