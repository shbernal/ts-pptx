// Author-on-template tests for `pptxgenjs/read` (dn-import-template-masters).
//
// Contract under test: Presentation.fromTemplate(input) opens a PowerPoint
// template (.pptx or .potx) as an empty deck shell — its slide masters, layouts,
// and theme stay byte-identical, any sample slides are stripped, and a .potx main
// part's template content type is normalized to the editable presentation type so
// the saved package opens as a normal deck. The shell is then authored onto with
// appendSlides() and saved, reusing the template's chrome verbatim.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'
import JSZip from 'jszip'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const PRESENTATION_MAIN_CT = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
const TEMPLATE_MAIN_CT = 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', name)
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

function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const matches = [...rels].filter((rel) => rel.type === type)
	if (matches.length === 0) return null
	return rels.resolveTarget(matches[0].id)
}

/** True for parts that are shared deck chrome (master/layout/theme). */
function isChromePart(name) {
	return /^ppt\/(slideMasters|slideLayouts|theme)\//.test(name)
}

function wideGenerator() {
	const pptx = new PptxGenJS()
	pptx.layout = 'LAYOUT_WIDE'
	return pptx
}

describe('Presentation.fromTemplate', () => {
	test('strips sample slides to a shell while keeping the layout gallery', async () => {
		const bytes = await readFile(fixturePath('multi-theme.pptx'))
		const plain = await Presentation.load(bytes)
		assert(plain.slides.length > 0, 'the fixture carries sample slides to strip')
		const galleryBefore = plain.layouts().map((l) => l.name)

		const deck = await Presentation.fromTemplate(bytes)
		assertEqual(deck.slides.length, 0, 'all sample slides are stripped')
		assertEqual(
			deck
				.layouts()
				.map((l) => l.name)
				.join('|'),
			galleryBefore.join('|'),
			'the layout gallery is preserved unchanged'
		)
	})

	test('keeps masters/layouts/theme byte-identical, touching only presentation.xml, its rels, and content types', async () => {
		const bytes = await readFile(fixturePath('multi-theme.pptx'))
		const before = await partBodies(bytes)

		const deck = await Presentation.fromTemplate(bytes)
		const after = await partBodies(await deck.save())

		// Every chrome part that survives is byte-identical.
		for (const [name, body] of before) {
			if (!isChromePart(name)) continue
			assert(after.has(name), `chrome part ${name} survives`)
			assert(bytesEqual(body, after.get(name)), `chrome part ${name} is byte-identical`)
		}

		// The only parts that change are presentation.xml, its .rels, and
		// [Content_Types].xml; slide parts are dropped; no chrome part is added.
		const changed = [...after.keys()].filter(
			(name) => !before.has(name) || !bytesEqual(before.get(name), after.get(name))
		)
		const allowed = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const name of changed)
			assert(allowed.has(name), `only chrome-neutral parts change; unexpected change: ${name}`)
		const removed = [...before.keys()].filter((name) => !after.has(name))
		for (const name of removed)
			assert(/^ppt\/slides\//.test(name), `only slide parts are removed; unexpected removal: ${name}`)
	})

	test('authors generated slides onto the template, binding to an existing layout (no new chrome)', async () => {
		const bytes = await readFile(fixturePath('multi-theme.pptx'))
		const before = await partBodies(bytes)

		const deck = await Presentation.fromTemplate(bytes)
		const target = deck.layouts().find((l) => l.name === 'Title and Content')
		assert(target, 'the template exposes a "Title and Content" layout to bind to')

		const pptx = wideGenerator()
		pptx.addSlide().addText('on template', { x: 1, y: 1, w: 6, h: 1, color: '0000FF' })
		const added = await deck.appendSlides(pptx, { layout: 'Title and Content' })
		assertEqual(added.length, 1, 'one slide was authored onto the shell')

		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides.length, 1, 'the saved deck has exactly the authored slide')

		const slide = reopened.slides[0]
		assertEqual(
			resolveSingle(reopened.opc, slide.partName, SLIDE_LAYOUT_REL),
			target.partName,
			'the authored slide binds to the template layout part'
		)
		assert(before.has(target.partName.slice(1)), 'the bound layout existed in the original template (no new chrome)')
	})

	test('normalizes a .potx main part to the editable presentation content type', async () => {
		const bytes = await readFile(fixturePath('template.potx'))
		const before = await partBodies(bytes)

		const deck = await Presentation.fromTemplate(bytes)
		const mainPartName = deck.presentationPart.partName
		assertEqual(
			deck.opc.contentTypes.contentTypeFor(mainPartName),
			PRESENTATION_MAIN_CT,
			'the .potx main part is flipped to the editable presentation content type'
		)
		assertEqual(deck.slides.length, 0, 'the template carries no sample slides (strip is a no-op)')

		// Chrome stays byte-identical through the content-type flip.
		const after = await partBodies(await deck.save())
		for (const [name, body] of before) {
			if (!isChromePart(name)) continue
			assert(bytesEqual(body, after.get(name)), `chrome part ${name} is byte-identical after the flip`)
		}
		const savedCt = new TextDecoder().decode(after.get('[Content_Types].xml'))
		assert(savedCt.includes(PRESENTATION_MAIN_CT), 'the saved package declares the editable main content type')
		assert(!savedCt.includes(TEMPLATE_MAIN_CT), 'the saved package no longer declares the template main content type')
	})

	test('keepTemplateContentType: true preserves the .potx template content type', async () => {
		const bytes = await readFile(fixturePath('template.potx'))
		const deck = await Presentation.fromTemplate(bytes, { keepTemplateContentType: true })
		assertEqual(
			deck.opc.contentTypes.contentTypeFor(deck.presentationPart.partName),
			TEMPLATE_MAIN_CT,
			'the template content type is preserved when requested'
		)
	})

	test.skipIf(!validatorInstalled)('the saved .potx-derived deck is schema-valid after authoring', async () => {
		const deck = await Presentation.fromTemplate(await readFile(fixturePath('template.potx')))
		const pptx = wideGenerator()
		pptx.addSlide().addText('valid on template', { x: 1, y: 1, w: 6, h: 1 })
		await deck.appendSlides(pptx, { layout: 'Title and Content' })

		const errors = await validateBuf(Buffer.from(await deck.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
