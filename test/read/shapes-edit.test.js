// Phase 4 structural-edit tests: adding and removing shapes via the slide DOM.
//
// Contract under test: Slide.addTextBox(...) appends a schema-valid p:sp with a
// slide-unique drawing id; Shape.delete() detaches a shape. Both mutate only
// the owning slide part, survive a save → reopen round-trip, leave untouched
// parts byte-identical, and keep the package schema-valid.

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
	return a.length === b.length && a.every((value, index) => value === b[index])
}

function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

describe('Slide.addTextBox', () => {
	test('appends a text box that reloads with its text and geometry', async () => {
		const presentation = await open('empty')
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
		const presentation = await open('textbox')
		const slide = presentation.slides[0]
		const existingIds = new Set(slide.shapes.map((shape) => shape.id))
		const box = slide.addTextBox({ text: 'x', left: 0, top: 0, width: 100000, height: 100000 })
		assert(typeof box.id === 'number', 'new box has a numeric id')
		assert(!existingIds.has(box.id), `new id ${box.id} is not reused`)
	})

	test('rejects non-positive or non-finite geometry', async () => {
		const slide = (await open('empty')).slides[0]
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
		for (const [name, body] of inputBodies) {
			if (name === dirty) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
	})
})

describe('Shape.delete', () => {
	test('removes a shape and the removal survives a reload', async () => {
		const presentation = await open('textbox')
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
})

describe('schema validity of structural edits', () => {
	test.skipIf(!validatorInstalled)('add + delete stays schema-valid', async () => {
		const presentation = await open('textbox')
		const slide = presentation.slides[0]
		slide.addTextBox({ text: 'Added', left: 914400, top: 914400, width: 1828800, height: 685800 })
		slide.shapes.find((shape) => shape.name === 'replaceText')?.delete()
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
