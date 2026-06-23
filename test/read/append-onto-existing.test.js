// Append-onto-existing tests for `pptxgenjs/read` (dn-append-onto-existing-deck).
//
// Contract under test: Presentation.appendSlides(source, { layout }) authors
// slides on a generator (PptxGenJS), serializes them via source.extractSlides(),
// and splices them into a loaded deck bound to an existing layout — keeping the
// deck's masters/layouts/theme (and every other untouched part) byte-identical,
// changing only presentation.xml, its .rels, [Content_Types].xml, and the new
// slide/media parts. Survives a save → reopen round-trip, resolves its layout to
// the *existing* layout (no new chrome), and stays schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'

// 1×1 transparent PNG.
const PNG_1PX =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

function bytesEqual(a, b) {
	return a && b && a.length === b.length && a.every((value, index) => value === b[index])
}

async function rejects(fn) {
	try {
		await fn()
		return false
	} catch {
		return true
	}
}

function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const matches = [...rels].filter((rel) => rel.type === type)
	if (matches.length === 0) return null
	return rels.resolveTarget(matches[0].id)
}

/** A generator deck sized to LAYOUT_WIDE (12192000×6858000), matching theme-colors / image. */
function wideGenerator() {
	const pptx = new PptxGenJS()
	pptx.layout = 'LAYOUT_WIDE'
	return pptx
}

describe('Presentation.appendSlides', () => {
	test('appends a generated slide bound to an existing layout, keeping chrome byte-identical', async () => {
		const originalBytes = await readFile(fixturePath('theme-colors'))
		const before = await partBodies(originalBytes)
		const pres = await Presentation.load(originalBytes)
		const beforeSlideCount = pres.slides.length

		const target = pres.layouts().find((l) => l.name === 'Blank')
		assert(target, 'theme-colors has a "Blank" layout to bind to')

		const pptx = wideGenerator()
		const slide = pptx.addSlide()
		slide.addText('hello append', { x: 1, y: 1, w: 6, h: 1, color: 'FF0000' })
		slide.addImage({ data: PNG_1PX, x: 1, y: 3, w: 1, h: 1 })

		const added = await pres.appendSlides(pptx, { layout: 'Blank' })
		assertEqual(added.length, 1, 'one slide was appended')

		const out = await pres.save()
		const after = await partBodies(out)

		// Byte-stability: every original part is byte-identical except the three the
		// append legitimately touches.
		const expectedChanged = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, bytes] of before) {
			assert(after.has(name), `part ${name} survives the append`)
			if (expectedChanged.has(name)) continue
			assert(bytesEqual(bytes, after.get(name)), `part ${name} is byte-identical after the append`)
		}

		// The only new parts are the slide, its .rels, and its media.
		const newParts = [...after.keys()].filter((name) => !before.has(name))
		for (const name of newParts) {
			assert(/^ppt\/(slides|media)\//.test(name), `new part ${name} is a slide or media part`)
		}
		assert(
			newParts.some((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)),
			'a new slide part was added'
		)
		assert(
			newParts.some((n) => /^ppt\/media\/image\d+\.png$/.test(n)),
			'a new media part was added'
		)

		// Round-trip: the new slide is present, last, and carries the authored text.
		const reopened = await Presentation.load(out)
		assertEqual(reopened.slides.length, beforeSlideCount + 1, 'the deck gained exactly one slide')
		const zipPath = added[0].partName.slice(1)
		const body = new TextDecoder().decode(after.get(zipPath))
		assert(body.includes('hello append'), 'the appended slide carries the authored text')

		// Layout binding resolves to the EXISTING layout — no new chrome.
		const newSlide = reopened.slides[reopened.slides.length - 1]
		assertEqual(
			resolveSingle(reopened.opc, newSlide.partName, SLIDE_LAYOUT_REL),
			target.partName,
			'the appended slide binds to the existing layout part'
		)
		assert(before.has(target.partName.slice(1)), 'the bound layout existed in the original deck')

		// Media resolves; no notes were generated.
		const image = resolveSingle(reopened.opc, newSlide.partName, IMAGE_REL)
		assert(image && reopened.opc.part(image), `the appended slide's image rel resolves (${image})`)
		assertEqual(resolveSingle(reopened.opc, newSlide.partName, NOTES_SLIDE_REL), null, 'no notes slide was generated')
	})

	test('accepts a LayoutHandle and inserts at a chosen position', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('image'))) // LAYOUT_WIDE, 2 slides
		const handle = pres.layouts().find((l) => l.name === 'Leer') // German "Blank"
		assert(handle, 'image deck has a "Leer" layout')

		const pptx = wideGenerator()
		pptx.addSlide().addText('inserted first', { x: 1, y: 1, w: 6, h: 1 })

		const [added] = await pres.appendSlides(pptx, { layout: handle, at: 0 })
		const reopened = await Presentation.load(await pres.save())
		assertEqual(reopened.slides.length, 3, 'the slide was added')
		assertEqual(reopened.slides[0].slideId, added.slideId, 'the appended slide landed first (at: 0)')
	})

	test.skipIf(!validatorInstalled)('the appended deck stays schema-valid', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		const slide = pptx.addSlide()
		slide.addText('valid', { x: 1, y: 1, w: 6, h: 1, color: '0000FF' })
		slide.addImage({ data: PNG_1PX, x: 1, y: 3, w: 1, h: 1 })
		await pres.appendSlides(pptx, { layout: 'Blank' })

		const errors = await validateBuf(Buffer.from(await pres.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test('rejects an unknown layout name', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 })
		assert(
			await rejects(() => pres.appendSlides(pptx, { layout: 'Nonexistent Layout' })),
			'an unknown layout name throws'
		)
	})

	test('rejects a slide-size mismatch', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors'))) // LAYOUT_WIDE
		const pptx = new PptxGenJS() // default LAYOUT_16x9 — narrower
		pptx.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 })
		assert(await rejects(() => pres.appendSlides(pptx, { layout: 'Blank' })), 'a mismatched slide size throws')
	})

	test('rejects an appended slide containing a chart (v1 limitation)', async () => {
		const pres = await Presentation.load(await readFile(fixturePath('theme-colors')))
		const pptx = wideGenerator()
		pptx
			.addSlide()
			.addChart(pptx.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], { x: 1, y: 1, w: 6, h: 3 })
		assert(
			await rejects(() => pres.appendSlides(pptx, { layout: 'Blank' })),
			'a chart slide throws the v1-unsupported error'
		)
	})
})
