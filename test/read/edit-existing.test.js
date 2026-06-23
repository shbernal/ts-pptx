// Acceptance test for the `dn-edit-existing-content` scenario: open a real
// PowerPoint-authored deck, *target* an existing shape or placeholder, swap its
// text and/or image, and save — while preserving sibling run formatting and
// leaving every untouched part byte-identical.
//
// This exercises the ergonomic surface that makes that scenario turnkey:
//   - Slide.shapeByName / shapeById / placeholder(type, idx?)  (addressing)
//   - Shape.text / TextFrame.text setters                       (whole-frame swap)
//   - Run.text setter                                           (sibling-preserving swap)
//   - Picture.setImage                                          (media swap)
// over the existing lossless OPC save (untouched parts pass through byte-for-byte).

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

// A 1×1 transparent PNG.
const PNG_1X1 = new Uint8Array(
	Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64')
)

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

describe('shape addressing', () => {
	test('shapeByName and shapeById resolve the same shape', async () => {
		const slide = (await open('textbox')).slides[0]
		const byName = slide.shapeByName('replaceText')
		assert(byName, 'shapeByName finds the textbox')
		const byId = slide.shapeById(byName.id)
		assert(byId, 'shapeById finds it by drawing id')
		assertEqual(byId.name, 'replaceText', 'shapeById returns the same shape')
		assertEqual(slide.shapeByName('does-not-exist'), undefined, 'missing name yields undefined')
	})

	test('placeholder(type, idx?) targets master/layout placeholders', async () => {
		const slide = (await open('mixed')).slides[0]
		const title = slide.placeholder('ctrTitle')
		assert(title, 'finds the centre-title placeholder')
		assertEqual(title.name, 'Titre 1', 'ctrTitle is the expected shape')
		assertEqual(title.placeholder.type, 'ctrTitle', 'exposes its placeholder type')

		const subtitle = slide.placeholder('subTitle', '1')
		assert(subtitle, 'finds the subtitle placeholder by type + idx')
		assertEqual(subtitle.name, 'Sous-titre 2', 'subTitle idx=1 is the expected shape')
		assertEqual(slide.placeholder('subTitle', '9'), undefined, 'wrong idx yields undefined')
	})

	test('non-placeholder shapes report a null placeholder', async () => {
		const slide = (await open('textbox')).slides[0]
		assertEqual(slide.shapeByName('replaceText').placeholder, null, 'a plain text box is not a placeholder')
	})
})

describe('whole-text-frame text swap', () => {
	test('Shape.text collapses to one run, preserving the first run formatting, and reloads', async () => {
		const presentation = await open('textbox')
		const shape = presentation.slides[0].shapeByName('replaceText')
		shape.text = 'BRAND NEW TEXT'

		const reopened = await Presentation.load(await presentation.save())
		const frame = reopened.slides[0].shapeByName('replaceText').textFrame
		assertEqual(frame.text, 'BRAND NEW TEXT', 'whole-frame text reloads')
		assertEqual(frame.paragraphs.length, 1, 'collapsed to a single paragraph')
		const run = frame.paragraphs[0].runs[0]
		assertEqual(frame.paragraphs[0].runs.length, 1, 'collapsed to a single run')
		// The first original run was italic, 20pt — that formatting carries over.
		assertEqual(run.italic, true, 'first run italic preserved')
		assertEqual(run.fontSizePt, 20, 'first run size preserved')
	})

	test('TextFrame.text behaves identically to Shape.text', async () => {
		const presentation = await open('textbox')
		presentation.slides[0].shapeByName('replaceText').textFrame.text = 'VIA FRAME'
		const reopened = await Presentation.load(await presentation.save())
		assertEqual(reopened.slides[0].shapeByName('replaceText').text, 'VIA FRAME', 'textFrame.text reloads')
	})

	test('placeholder text can be replaced in place', async () => {
		const presentation = await open('mixed')
		presentation.slides[0].placeholder('ctrTitle').text = 'Replaced Title'
		const reopened = await Presentation.load(await presentation.save())
		assertEqual(reopened.slides[0].placeholder('ctrTitle').text, 'Replaced Title', 'placeholder title reloads')
	})

	test('Shape.text throws on a shape with no text frame', async () => {
		const picture = (await open('image')).slides[0].shapeByName('Grafik 5')
		assert(picture, 'fixture has the picture')
		assert(
			throws(() => {
				picture.text = 'nope'
			}),
			'setting text on a picture throws'
		)
	})
})

describe('targeted run edit preserves sibling run formatting', () => {
	test('replacing one run leaves its siblings (italic/bold) untouched', async () => {
		const presentation = await open('textbox')
		const frame = presentation.slides[0].shapeByName('replaceText').textFrame
		// Para 2 holds the "{{replace}}" run (16pt) among differently-formatted siblings.
		const para = frame.paragraphs[2]
		const target = para.runs.find((run) => run.text === '{{replace}}')
		assert(target, 'found the {{replace}} run')
		target.text = 'VALUE'

		const reopened = await Presentation.load(await presentation.save())
		const reframe = reopened.slides[0].shapeByName('replaceText').textFrame
		assert(reframe.text.includes('VALUE'), 'the targeted run text changed')
		// Sibling runs in paragraph 0 keep their original character formatting.
		const p0 = reframe.paragraphs[0]
		assertEqual(p0.runs[0].italic, true, 'the italic "This" run is preserved')
		assertEqual(p0.runs.find((run) => run.text === 'content').bold, true, 'the bold "content" run is preserved')
	})
})

describe('acceptance: target a shape, swap text + image, untouched parts byte-stable', () => {
	test('image.pptx — replace a text box and swap a picture, leaving other parts byte-identical', async () => {
		const input = await readFile(fixturePath('image'))
		const presentation = await Presentation.load(input)
		const slide = presentation.slides[0]

		slide.shapeByName('Textfeld 1').text = 'swapped caption'
		const picture = slide.shapeByName('Grafik 5')
		assertEqual(picture.shapeType, 'picture', 'Grafik 5 is the picture')
		const oldPartName = picture.imagePartName
		picture.setImage(PNG_1X1, { contentType: 'image/png' })

		const saved = await presentation.save()
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(saved)

		// Only the edited slide, its rels, and the content-types map may change.
		const allowedToChange = new Set([
			'ppt/slides/slide1.xml',
			'ppt/slides/_rels/slide1.xml.rels',
			'[Content_Types].xml',
		])
		for (const [name, body] of inputBodies) {
			if (allowedToChange.has(name)) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be byte-identical after the edit`)
		}
		// The original media part is preserved byte-identical (copy-on-write).
		assert(inputBodies.has(oldPartName.slice(1)), 'original media part still present in input')
		assert(
			bytesEqual(inputBodies.get(oldPartName.slice(1)), outputBodies.get(oldPartName.slice(1))),
			'old media untouched'
		)
		const newMedia = [...outputBodies.keys()].filter((name) => name.startsWith('ppt/media/') && !inputBodies.has(name))
		assertEqual(newMedia.length, 1, `exactly one new media part, got ${JSON.stringify(newMedia)}`)

		// Re-read the saved deck and confirm the edits took.
		const reopened = await Presentation.load(saved)
		assertEqual(reopened.slides[0].shapeByName('Textfeld 1').text, 'swapped caption', 'caption edit reloads')
		const newPartName = reopened.slides[0].shapeByName('Grafik 5').imagePartName
		assert(newPartName !== oldPartName, 'picture points at a new media part')
		assert(bytesEqual(reopened.opc.part(newPartName).bytes, PNG_1X1), 'new media holds the supplied bytes')
	})

	test.skipIf(!validatorInstalled)('the edited deck stays schema-valid', async () => {
		const presentation = await open('image')
		const slide = presentation.slides[0]
		slide.shapeByName('Textfeld 1').text = 'swapped caption'
		slide.shapeByName('Grafik 5').setImage(PNG_1X1, { contentType: 'image/png' })
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
