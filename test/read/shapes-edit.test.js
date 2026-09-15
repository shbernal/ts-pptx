// Structural-edit tests: adding and removing shapes via the slide DOM.
//
// Contract under test: Slide.addTextBox(...) appends a schema-valid p:sp with a
// slide-unique drawing id; Shape.delete() detaches a shape. Both mutate only
// the owning slide part, survive a save → reopen round-trip, leave untouched
// parts byte-identical, and keep the package schema-valid.

import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import TsPptx, { ShapeType } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { throws, bytesEqual, assert, assertEqual, partBodies, assertUnchangedExcept } from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'

describe('Slide.addTextBox', () => {
	test('appends a text box that reloads with its text and geometry', async () => {
		const presentation = await openFixture('empty')
		const before = presentation.slides[0].shapes.length
		const box = presentation.slides[0].addTextBox({
			text: 'Hello',
			left: 914400,
			top: 457200,
			width: 1828800,
			height: 685800,
			name: 'MyBox',
		})
		assertEqual(box.shapeType, 'autoShape', 'returns an AutoShape')
		assertEqual(box.text, 'Hello', 'returned box reports its text')

		const reopened = await Presentation.load(await presentation.save())
		const shapes = reopened.slides[0].shapes
		assertEqual(shapes.length, before + 1, 'shape count grew by one')
		const reloaded = shapes.find((shape) => shape.name === 'MyBox')
		assert(reloaded, 'added box reloads by name')
		assertEqual(reloaded.text, 'Hello', 'text reloads')
		assertEqual(reloaded.left, 914400, 'left reloads')
		assertEqual(reloaded.width, 1828800, 'width reloads')
	})

	test('allocates a drawing id unique within the slide', async () => {
		const presentation = await openFixture('textbox')
		const slide = presentation.slides[0]
		const existingIds = new Set(slide.shapes.map((shape) => shape.id))
		const box = slide.addTextBox({ text: 'x', left: 0, top: 0, width: 100000, height: 100000 })
		assert(typeof box.id === 'number', 'new box has a numeric id')
		assert(!existingIds.has(box.id), `new id ${box.id} is not reused`)
	})

	test('rejects non-positive or non-finite geometry', async () => {
		const slide = (await openFixture('empty')).slides[0]
		const base = { left: 0, top: 0, width: 100000, height: 100000 }
		assert(
			throws(() => slide.addTextBox({ ...base, width: 0 })),
			'zero width throws'
		)
		assert(
			throws(() => slide.addTextBox({ ...base, height: -1 })),
			'negative height throws'
		)
		assert(
			throws(() => slide.addTextBox({ ...base, left: Number.NaN })),
			'NaN left throws'
		)
	})

	test('adding a shape leaves every other part byte-identical', async () => {
		const input = await readFile(fixturePath('empty'))
		const presentation = await Presentation.load(input)
		presentation.slides[0].addTextBox({ text: 'x', left: 0, top: 0, width: 100000, height: 100000 })
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())
		const dirty = 'ppt/slides/slide1.xml'
		assert(!bytesEqual(inputBodies.get(dirty), outputBodies.get(dirty)), 'edited slide differs')
		assertUnchangedExcept(inputBodies, outputBodies, [dirty])
	})
})

describe('Shape.delete', () => {
	test('removes a shape and the removal survives a reload', async () => {
		const presentation = await openFixture('textbox')
		const slide = presentation.slides[0]
		const before = slide.shapes.length
		const target = slide.shapes.find((shape) => shape.name === 'replaceText')
		assert(target, 'precondition: replaceText shape exists')
		target.delete()

		const reopened = await Presentation.load(await presentation.save())
		const shapes = reopened.slides[0].shapes
		assertEqual(shapes.length, before - 1, 'shape count shrank by one')
		assert(!shapes.some((shape) => shape.name === 'replaceText'), 'deleted shape is gone')
	})

	// A build left naming a shape that is not on the slide makes PowerPoint refuse the deck
	// (0x80070570), which is what deleting an animated shape used to save.
	test('removes the build animations that target the deleted shape', async () => {
		const presentation = await openFixture('slide-animation-rich')
		const slide = presentation.slides[0]
		assertEqual(slide.animationSpids().join(','), '2,3,4,5', 'precondition: four animated shapes')
		const target = slide.shapeByIdDeep(3)
		assert(target, 'precondition: shape 3 exists')
		target.delete()

		const reopened = (await Presentation.load(await presentation.save())).slides[0]
		assertEqual(reopened.animationSpids().join(','), '2,4,5', 'no build names the deleted shape')
		assertEqual(reopened.shapes.length, 3, 'the other animated shapes stay')
	})

	// PowerPoint keeps a binding to a missing id through its own save, and a shape later given
	// that id would inherit the connector.
	test('unbinds a connector from the deleted shape and keeps its other end', async () => {
		const pptx = new TsPptx()
		const s = pptx.addSlide()
		s.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'A' })
		s.addShape(ShapeType.rect, { x: 6, y: 3, w: 2, h: 1, objectName: 'B' })
		s.addConnector({ x1: 3, y1: 1.5, x2: 6, y2: 3.5, startShape: 'A', startShapeIdx: 3, endShape: 'B', endShapeIdx: 1 })
		const presentation = await Presentation.load(await pptx.write({ outputType: 'uint8array' }))
		const slideXmlOf = async (/** @type {Uint8Array} */ bytes) =>
			(await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml').async('string')
		assert(/<a:stCxn\b/.test(await slideXmlOf(await presentation.save())), 'precondition: the start is bound')

		presentation.slides[0].shapeByName('A')?.delete()
		const xml = await slideXmlOf(await presentation.save())
		assert(!/<a:stCxn\b/.test(xml), 'the binding to the deleted shape is gone')
		assert(/<a:endCxn id="\d+" idx="1"\/>/.test(xml), 'the binding to the shape that stays is kept')
		assert(/<p:cxnSp>/.test(xml), 'the connector stays')
	})

	test('deleting a group removes the animations of the shapes inside it', async () => {
		const pptx = new TsPptx()
		const s = pptx.addSlide()
		s.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, objectName: 'Inside' } }], { objectName: 'Outer' })
		s.addShape(ShapeType.rect, { x: 6, y: 3, w: 2, h: 1, objectName: 'Outside' })
		s.addAnimation({ preset: 'fadeIn', objectName: 'Inside' })
		s.addAnimation({ preset: 'fadeIn', objectName: 'Outside' })
		const presentation = await Presentation.load(await pptx.write({ outputType: 'uint8array' }))
		const slide = presentation.slides[0]
		const outsideId = slide.shapeByName('Outside')?.id
		assertEqual(slide.animationSpids().length, 2, 'precondition: one build inside the group, one outside')

		slide.shapeByName('Outer')?.delete()
		const reopened = (await Presentation.load(await presentation.save())).slides[0]
		assertEqual(reopened.animationSpids().join(','), String(outsideId), 'only the build outside the group is left')
	})
})

describe('schema validity of structural edits', () => {
	test.skipIf(!validatorInstalled)('add + delete stays schema-valid', async () => {
		const presentation = await openFixture('textbox')
		const slide = presentation.slides[0]
		slide.addTextBox({ text: 'Added', left: 914400, top: 914400, width: 1828800, height: 685800 })
		slide.shapes.find((shape) => shape.name === 'replaceText')?.delete()
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
