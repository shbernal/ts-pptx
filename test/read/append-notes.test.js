// Speaker notes across the append path.
//
// Contract under test: a generator slide authored with `addNotes` gains a real
// `notesSlide` part when spliced in by `Presentation.appendSlides` — wired back to the
// slide it annotates, bound to a notes master, and with its hyperlink rels preserved.
//
// The notes-master policy is the interesting half. A notes slide must bind to a notes
// master, and a template commonly has none (a deck authored without notes carries no
// `notesMaster` part), so the generator ships one. But a destination that *does* have one
// keeps it — the same "destination styling wins" rule `importNotes` follows.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await validatorAvailable()

const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const NOTES_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', name)
}

/** Author `build` onto a template shell, append, save, and hand back the bytes + re-read deck. */
async function appendOnto(fixture, build) {
	const bytes = await readFile(fixturePath(fixture))
	const deck = await Presentation.fromTemplate(bytes)
	const size = deck.slideSize
	const pptx = new TsPptx()
	pptx.defineLayout({ name: 'MATCH', width: size.widthEmu / 914400, height: size.heightEmu / 914400 })
	pptx.layout = 'MATCH'
	build(pptx)
	await deck.appendSlides(pptx, { layout: deck.layouts()[0] })
	const out = await deck.save()
	return { out, reread: await Presentation.load(out) }
}

function relsOf(opc, partName) {
	return [...opc.relationshipsFor(partName)]
}

describe('appendSlides carries speaker notes', () => {
	test('a notes-bearing slide gains a notesSlide part wired back to it', async () => {
		const { reread } = await appendOnto('placeholder-inherit.pptx', (pptx) => {
			const slide = pptx.addSlide()
			slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
			slide.addNotes('These are the notes')
		})

		const slide = reread.slides[0]
		assertEqual(slide.notesText, 'These are the notes', 'notes text round-trips')

		// The slide -> notesSlide rel resolves to a real part...
		const slideRels = reread.opc.relationshipsFor(slide.partName)
		const notesRel = [...slideRels].find((r) => r.type === NOTES_SLIDE_REL)
		assert(notesRel, 'the appended slide has a notesSlide relationship')
		const notesPartName = slideRels.resolveTarget(notesRel.id)
		assert(reread.opc.part(notesPartName), `notesSlide part ${notesPartName} exists`)

		// ...and the notesSlide points back at the slide it annotates, not at a stale target.
		const notesRels = reread.opc.relationshipsFor(notesPartName)
		const backRel = [...notesRels].find((r) => r.type === SLIDE_REL)
		assert(backRel, 'the notesSlide has a back-relationship to its slide')
		assertEqual(notesRels.resolveTarget(backRel.id), slide.partName, 'the back-rel targets the appended slide')
	})

	test('a slide without notes gains no notesSlide part', async () => {
		const { reread } = await appendOnto('placeholder-inherit.pptx', (pptx) => {
			pptx.addSlide().addText('Body', { x: 1, y: 1, w: 4, h: 1 })
		})
		assertEqual(reread.slides[0].notesText, null, 'no notes were invented')
		const notesParts = [...reread.opc.parts.keys()].filter((n) => /notesSlides\/notesSlide\d+\.xml$/.test(n))
		assertEqual(notesParts.length, 0, 'no notesSlide part was added')
	})

	test('installs a notes master when the template has none, and registers it once', async () => {
		const templateBytes = await readFile(fixturePath('placeholder-inherit.pptx'))
		const templateParts = [...(await Presentation.load(templateBytes)).opc.parts.keys()]
		assert(
			!templateParts.some((n) => /notesMasters\/notesMaster\d+\.xml$/.test(n)),
			'precondition: this template carries no notesMaster'
		)

		const { reread } = await appendOnto('placeholder-inherit.pptx', (pptx) => {
			const a = pptx.addSlide()
			a.addText('A', { x: 1, y: 1, w: 4, h: 1 })
			a.addNotes('notes A')
			const b = pptx.addSlide()
			b.addText('B', { x: 1, y: 1, w: 4, h: 1 })
			b.addNotes('notes B')
		})

		const masters = [...reread.opc.parts.keys()].filter((n) => /notesMasters\/notesMaster\d+\.xml$/.test(n))
		assertEqual(masters.length, 1, 'exactly one notesMaster is installed for two notes slides')

		// CT_NotesMasterIdList holds a single p:notesMasterId; two notes slides must not add two.
		const presRels = relsOf(reread.opc, reread.presentationPart.partName)
		assertEqual(presRels.filter((r) => r.type === NOTES_MASTER_REL).length, 1, 'one notesMaster rel')

		// The notesMaster's own .rels must resolve a theme, or the part dangles.
		const masterRels = relsOf(reread.opc, masters[0])
		const themeRel = masterRels.find((r) => r.type.endsWith('/theme'))
		assert(themeRel, 'the installed notesMaster references a theme')
		const themeTarget = reread.opc.relationshipsFor(masters[0]).resolveTarget(themeRel.id)
		assert(reread.opc.part(themeTarget), `notesMaster theme part ${themeTarget} exists`)

		// Both slides' notes survive and stay distinct.
		assertEqual(reread.slides[0].notesText, 'notes A', 'first slide notes')
		assertEqual(reread.slides[1].notesText, 'notes B', 'second slide notes')
	})

	test("reuses the template's own notes master instead of installing a second", async () => {
		const templateBytes = await readFile(fixturePath('mixed.pptx'))
		const before = [...(await Presentation.load(templateBytes)).opc.parts.keys()].filter((n) =>
			/notesMasters\/notesMaster\d+\.xml$/.test(n)
		)
		assertEqual(before.length, 1, 'precondition: this template already carries a notesMaster')

		const { reread } = await appendOnto('mixed.pptx', (pptx) => {
			const slide = pptx.addSlide()
			slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
			slide.addNotes('destination master wins')
		})

		const after = [...reread.opc.parts.keys()].filter((n) => /notesMasters\/notesMaster\d+\.xml$/.test(n))
		assertEqual(after.length, 1, 'no second notesMaster was installed')
		assertEqual(after[0], before[0], "the template's own notesMaster partname is reused")
		assertEqual(reread.slides[0].notesText, 'destination master wins', 'notes still round-trip')
	})

	test('notes hyperlink rels are preserved alongside the reserved notesMaster/slide rels', async () => {
		const { reread } = await appendOnto('placeholder-inherit.pptx', (pptx) => {
			const slide = pptx.addSlide()
			slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
			slide.addNotes([{ text: 'see ' }, { text: 'docs', options: { hyperlink: { url: 'https://example.com/' } } }])
		})

		const slide = reread.slides[0]
		const slideRels = reread.opc.relationshipsFor(slide.partName)
		const notesPartName = slideRels.resolveTarget([...slideRels].find((r) => r.type === NOTES_SLIDE_REL).id)
		const notesRels = relsOf(reread.opc, notesPartName)

		const hyperlinks = notesRels.filter((r) => r.type.endsWith('/hyperlink'))
		assertEqual(hyperlinks.length, 1, 'the notes hyperlink rel survived')
		assertEqual(hyperlinks[0].target, 'https://example.com/', 'hyperlink target')
		// Reserved ids must not have been overwritten by the hyperlink.
		assert(
			notesRels.some((r) => r.id === 'rId1' && r.type === NOTES_MASTER_REL),
			'rId1 is the notesMaster'
		)
		assert(
			notesRels.some((r) => r.id === 'rId2' && r.type === SLIDE_REL),
			'rId2 is the slide'
		)
	})

	test('every notes part is content-type registered', async () => {
		const { out } = await appendOnto('placeholder-inherit.pptx', (pptx) => {
			const slide = pptx.addSlide()
			slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
			slide.addNotes('registered?')
		})
		const zip = await JSZip.loadAsync(out)
		const contentTypes = await zip.file('[Content_Types].xml').async('string')
		for (const [part, type] of [
			['/ppt/notesSlides/notesSlide1.xml', 'notesSlide+xml'],
			['/ppt/notesMasters/notesMaster1.xml', 'notesMaster+xml'],
		]) {
			assert(
				contentTypes.includes(`PartName="${part}"`),
				`[Content_Types].xml declares an Override for ${part}; got:\n${contentTypes}`
			)
			assert(contentTypes.includes(type), `[Content_Types].xml carries the ${type} content type`)
		}
	})

	test.skipIf(!validatorInstalled)('the resulting package is schema-valid', async () => {
		const { out } = await appendOnto('placeholder-inherit.pptx', (pptx) => {
			const slide = pptx.addSlide()
			slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
			slide.addNotes([{ text: 'plain ' }, { text: 'link', options: { hyperlink: { url: 'https://example.com/' } } }])
		})
		const errors = await validateBuf(out)
		assertEqual(errors.length, 0, `expected a schema-clean package, got:\n${JSON.stringify(errors, null, 2)}`)
	})
})
