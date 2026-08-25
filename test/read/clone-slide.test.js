// Phase 4 slide-cloning tests for `ts-pptx/read`.
//
// Contract under test: Presentation.cloneSlide(index) appends an independent
// duplicate (its own slide part + copied .rels), wires a presentation→slide
// relationship and a p:sldId entry, survives a save → reopen round-trip, leaves
// parts it does not touch byte-identical, and keeps the package schema-valid.

import { readFile } from 'node:fs/promises'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { throws, assert, assertEqual, partBodies, assertUnchangedExcept } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'

const validatorInstalled = await validatorAvailable()

describe('Presentation.cloneSlide', () => {
	test('appends an independent duplicate that reloads with the source content', async () => {
		const presentation = await openFixture('textbox')
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
		const presentation = await openFixture('textbox')
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
		assertUnchangedExcept(inputBodies, outputBodies, [
			'ppt/presentation.xml',
			'ppt/_rels/presentation.xml.rels',
			'[Content_Types].xml',
		])
		const added = [...outputBodies.keys()].filter((name) => !inputBodies.has(name))
		assert(added.includes('ppt/slides/slide3.xml'), `new slide part added: ${JSON.stringify(added)}`)
		assert(added.includes('ppt/slides/_rels/slide3.xml.rels'), 'new slide rels added')
	})

	test('a slide imported this session clones with its relationships', async () => {
		// The clone used to copy the source's `.rels` *part bytes*, and a page brought
		// in by `importSlide` holds its relationships in memory until the deck is
		// saved: there was no such part yet, so the clone came out with no
		// relationships at all — its `r:id`s resolved to nothing and the chart on it
		// had no chart part behind it.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 7) // the fixture's chart page
		const clone = target.cloneSlide(imported.index)

		const reopened = await Presentation.load(await target.save())
		const rels = reopened.opc.relationshipsFor(clone.partName)
		const types = [...rels].map((rel) => rel.type.split('/').pop()).sort()
		assertEqual(JSON.stringify(types), JSON.stringify(['chart', 'slideLayout']), 'the clone kept both relationships')
		const chartRel = [...rels].find((rel) => rel.type.endsWith('/chart'))
		assert(
			reopened.opc.part(rels.resolveTarget(chartRel.id)) !== undefined,
			'and the chart relationship resolves to a part that is in the package'
		)
	})

	test('rejects an out-of-range index', async () => {
		const presentation = await openFixture('textbox')
		assert(
			throws(() => presentation.cloneSlide(99)),
			'cloning a missing slide throws'
		)
	})

	test.skipIf(!validatorInstalled)('a deck with a cloned slide stays schema-valid', async () => {
		const presentation = await openFixture('textbox')
		presentation.cloneSlide(0)
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
