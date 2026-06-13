// Phase 4 slide-cloning tests for `pptxgenjs/read`.
//
// Contract under test: Presentation.cloneSlide(index) appends an independent
// duplicate (its own slide part + copied .rels), wires a presentation→slide
// relationship and a p:sldId entry, survives a save → reopen round-trip, leaves
// parts it does not touch byte-identical, and keeps the package schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
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

function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

describe('Presentation.cloneSlide', () => {
	test('appends an independent duplicate that reloads with the source content', async () => {
		const presentation = await open('textbox')
		const beforeCount = presentation.slides.length
		const sourceText = presentation.slides[0].shapes.find((s) => s.hasTextFrame).text
		const clone = presentation.cloneSlide(0)
		assertEqual(presentation.slides.length, beforeCount + 1, 'a slide was appended in-memory')
		assertEqual(clone.index, beforeCount, 'clone is the last slide')

		const reopened = await Presentation.load(await presentation.save())
		assertEqual(reopened.slides.length, beforeCount + 1, 'slide count grew after reload')
		const last = reopened.slides[reopened.slides.length - 1]
		const lastText = last.shapes.find((s) => s.hasTextFrame)?.text
		assertEqual(lastText, sourceText, 'clone carries the source slide text')
		// Slide ids are unique.
		const ids = reopened.slides.map((s) => s.slideId)
		assertEqual(new Set(ids).size, ids.length, 'slide ids are unique')
	})

	test('clone is independent of the source (editing one does not affect the other)', async () => {
		const presentation = await open('textbox')
		const clone = presentation.cloneSlide(0)
		clone.shapes.find((s) => s.hasTextFrame).textFrame.paragraphs[0].runs[0].text = 'CLONE ONLY'

		const reopened = await Presentation.load(await presentation.save())
		const sourceRun = reopened.slides[0].shapes.find((s) => s.hasTextFrame).textFrame.paragraphs[0].runs[0].text
		const cloneRun = reopened.slides[reopened.slides.length - 1].shapes.find((s) => s.hasTextFrame).textFrame
			.paragraphs[0].runs[0].text
		assertEqual(cloneRun, 'CLONE ONLY', 'edit landed on the clone')
		assert(sourceRun !== 'CLONE ONLY', 'source slide is untouched by the clone edit')
	})

	test('only the presentation part + its rels change; the clone parts are added', async () => {
		const input = await readFile(fixturePath('textbox'))
		const presentation = await Presentation.load(input)
		presentation.cloneSlide(0)
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())

		// The presentation part + its rels change; [Content_Types].xml gains an
		// Override for the new slide part (the xml Default maps to application/xml).
		const allowedToChange = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, body] of inputBodies) {
			if (allowedToChange.has(name)) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
		const added = [...outputBodies.keys()].filter((name) => !inputBodies.has(name))
		assert(added.includes('ppt/slides/slide3.xml'), `new slide part added: ${JSON.stringify(added)}`)
		assert(added.includes('ppt/slides/_rels/slide3.xml.rels'), 'new slide rels added')
	})

	test('rejects an out-of-range index', async () => {
		const presentation = await open('textbox')
		assert(
			throws(() => presentation.cloneSlide(99)),
			'cloning a missing slide throws'
		)
	})

	test.skipIf(!validatorInstalled)('a deck with a cloned slide stays schema-valid', async () => {
		const presentation = await open('textbox')
		presentation.cloneSlide(0)
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
