// Phase 3 edit-vertical-slice tests for `pptxgenjs/read` (src/read/api/).
//
// Contract under test: mutating a Run (text + font props) or a Shape's geometry
// through the read model mutates the live DOM, marks only the owning slide part
// dirty, survives a save → reopen round-trip, leaves untouched parts
// byte-identical, and keeps the package schema-valid.

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

/** Open a fixture, mutate it via `edit`, then reopen the saved bytes. */
async function editAndReopen(name, edit) {
	const presentation = await open(name)
	await edit(presentation)
	const saved = await presentation.save()
	return { presentation, saved, reopened: await Presentation.load(saved) }
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

function replaceTextShape(presentation) {
	return presentation.slides[0].shapes.find((shape) => shape.name === 'replaceText')
}

describe('Run text editing', () => {
	test('run.text mutates the a:t node and survives a reload', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			replaceTextShape(presentation).textFrame.paragraphs[0].runs[0].text = 'CHANGED'
		})
		assertEqual(replaceTextShape(reopened).textFrame.paragraphs[0].runs[0].text, 'CHANGED', 'edited run text reloads')
	})

	test('whitespace-significant text gets xml:space="preserve"', async () => {
		const { saved } = await editAndReopen('textbox', (presentation) => {
			replaceTextShape(presentation).textFrame.paragraphs[0].runs[0].text = '  spaced  '
		})
		const slideXml = new TextDecoder().decode((await partBodies(saved)).get('ppt/slides/slide1.xml'))
		assert(slideXml.includes('xml:space="preserve"'), 'preserve attr present')
		assert(slideXml.includes('>  spaced  </a:t>'), 'whitespace-significant text written verbatim')
	})

	test('editing one slide leaves every other part byte-identical', async () => {
		const input = await readFile(fixturePath('textbox'))
		const presentation = await Presentation.load(input)
		replaceTextShape(presentation).textFrame.paragraphs[0].runs[0].text = 'CHANGED'
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())
		const dirty = 'ppt/slides/slide1.xml'
		assert(!bytesEqual(inputBodies.get(dirty), outputBodies.get(dirty)), 'dirty slide body should differ')
		for (const [name, body] of inputBodies) {
			if (name === dirty) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
	})
})

describe('Run font properties', () => {
	test('sets size, bold, font, and explicit colour; clears the prior scheme colour', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			const run = replaceTextShape(presentation).textFrame.paragraphs[0].runs[0]
			run.fontSizePt = 32
			run.bold = true
			run.fontName = 'Georgia'
			run.color = 'FF0000' // run[0] starts with a schemeClr fill; this must replace it
		})
		const run = replaceTextShape(reopened).textFrame.paragraphs[0].runs[0]
		assertEqual(run.fontSizePt, 32, 'font size reloads (3200 → 32pt)')
		assertEqual(run.bold, true, 'bold reloads')
		assertEqual(run.fontName, 'Georgia', 'font name reloads')
		assertEqual(run.color, 'FF0000', 'explicit srgb colour reloads')
		assertEqual(run.schemeColor, null, 'scheme colour was replaced by the srgb fill')
		assertEqual(run.italic, true, 'untouched italic flag is preserved')
	})

	test('setting a boolean prop to null removes it (back to inherited)', async () => {
		// run[0] is italic; clearing it should drop the @i attribute, not set i="0".
		const { saved, reopened } = await editAndReopen('textbox', (presentation) => {
			replaceTextShape(presentation).textFrame.paragraphs[0].runs[0].italic = null
		})
		assertEqual(replaceTextShape(reopened).textFrame.paragraphs[0].runs[0].italic, null, 'italic now inherited')
		const slideXml = new TextDecoder().decode((await partBodies(saved)).get('ppt/slides/slide1.xml'))
		assert(!/<a:rPr[^>]*\bi="0"/.test(slideXml), 'must not emit i="0"; the attribute is removed')
	})

	test('creates an a:rPr when a plain run gains a property', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			const plain = replaceTextShape(presentation).textFrame.paragraphs[0].runs.find((run) => run.text === ' is test')
			assertEqual(plain.bold, null, 'precondition: plain run has no rPr/@b')
			plain.bold = true
		})
		const plain = replaceTextShape(reopened).textFrame.paragraphs[0].runs.find((run) => run.text === ' is test')
		assert(plain, 'the " is test" run still exists')
		assertEqual(plain.bold, true, 'bold persisted via a freshly created rPr')
	})

	test('schemeColor setter replaces an explicit srgb fill', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			const run = replaceTextShape(presentation).textFrame.paragraphs[0].runs[0]
			run.schemeColor = 'accent4'
		})
		const run = replaceTextShape(reopened).textFrame.paragraphs[0].runs[0]
		assertEqual(run.schemeColor, 'accent4', 'scheme colour reloads')
		assertEqual(run.color, null, 'no explicit srgb colour remains')
	})

	test('rejects a non-positive font size and a malformed colour', async () => {
		const run = replaceTextShape(await open('textbox')).textFrame.paragraphs[0].runs[0]
		assert(
			throws(() => (run.fontSizePt = 0)),
			'fontSizePt = 0 should throw'
		)
		assert(
			throws(() => (run.fontSizePt = Number.NaN)),
			'fontSizePt = NaN should throw'
		)
		assert(
			throws(() => (run.color = 'nothex')),
			'malformed colour should throw'
		)
	})
})

describe('Shape geometry editing', () => {
	test('sets left/top/width/height on an auto shape', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			shape.left = 914400
			shape.top = 457200
			shape.width = 1828800
			shape.height = 685800
		})
		const shape = replaceTextShape(reopened)
		assertEqual(shape.left, 914400, 'left reloads')
		assertEqual(shape.top, 457200, 'top reloads')
		assertEqual(shape.width, 1828800, 'width reloads')
		assertEqual(shape.height, 685800, 'height reloads')
	})

	test('rounds fractional EMU and rejects NaN / negative extents', async () => {
		const shape = replaceTextShape(await open('textbox'))
		shape.left = 100.6
		assertEqual(shape.left, 101, 'fractional EMU is rounded')
		assert(
			throws(() => (shape.width = Number.NaN)),
			'NaN width should throw'
		)
		assert(
			throws(() => (shape.width = -10)),
			'negative width should throw'
		)
	})

	test('sets geometry on a graphic frame (p:xfrm)', async () => {
		const { reopened } = await editAndReopen('table', (presentation) => {
			const frame = presentation.slides
				.flatMap((slide) => slide.shapes)
				.find((shape) => shape.shapeType === 'graphicFrame')
			frame.left = 1000000
			frame.width = 5000000
		})
		const frame = reopened.slides.flatMap((slide) => slide.shapes).find((shape) => shape.shapeType === 'graphicFrame')
		assertEqual(frame.left, 1000000, 'graphic frame left reloads')
		assertEqual(frame.width, 5000000, 'graphic frame width reloads')
	})
})

describe('Slide.hidden editing', () => {
	test('hiding a slide writes show="0" and survives a reload', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			assertEqual(presentation.slides[0].hidden, false, 'slide starts shown')
			presentation.slides[0].hidden = true
		})
		assertEqual(reopened.slides[0].hidden, true, 'hidden state reloads')
	})

	test('showing a hidden slide removes the attribute (canonical shown form)', async () => {
		const { reopened, saved } = await editAndReopen('hidden', (presentation) => {
			assertEqual(presentation.slides[1].hidden, true, 'slide 2 starts hidden')
			presentation.slides[1].hidden = false
		})
		assertEqual(reopened.slides[1].hidden, false, 'shown state reloads')
		const bodies = await partBodies(saved)
		const slide2 = new TextDecoder().decode(bodies.get('ppt/slides/slide2.xml'))
		assert(!slide2.includes('show='), `@show should be absent when shown; got: ${slide2.slice(0, 200)}`)
	})

	test('toggling hidden marks only the owning slide part dirty', async () => {
		const presentation = await open('textbox')
		const inputBodies = await partBodies(await presentation.save())
		presentation.slides[1].hidden = true
		const outputBodies = await partBodies(await presentation.save())
		const dirty = 'ppt/slides/slide2.xml'
		assert(!bytesEqual(inputBodies.get(dirty), outputBodies.get(dirty)), 'hidden slide body should differ')
		for (const [name, bytes] of inputBodies) {
			if (name === dirty) continue
			assert(bytesEqual(bytes, outputBodies.get(name)), `untouched part ${name} should be byte-identical`)
		}
	})
})

describe('schema validity of edited packages', () => {
	test.skipIf(!validatorInstalled)('a text + font + geometry edit stays schema-valid', async () => {
		const { saved } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			shape.left = 914400
			shape.width = 1828800
			const run = shape.textFrame.paragraphs[0].runs[0]
			run.text = 'Edited'
			run.fontSizePt = 28
			run.bold = true
			run.color = '1A2B3C'
		})
		const errors = await validateBuf(Buffer.from(saved))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}
