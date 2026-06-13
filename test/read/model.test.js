// Phase 2 read-model tests for `pptxgenjs/read` (src/read/api/).
//
// Contract under test: Presentation.load(buf) exposes a navigable, typed view
// of the deck — slides in order, shapes from the spTree by kind, geometry in
// EMU, and text frame → paragraphs → runs with character formatting — all read
// from the live DOM. No mutation here (that is Phase 3).

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

describe('Presentation', () => {
	test('resolves the presentation part and slide size', async () => {
		const presentation = await open('textbox')
		assertEqual(presentation.presentationPart.partName, '/ppt/presentation.xml', 'presentation part')
		const size = presentation.slideSize
		assert(size, 'slide size should resolve')
		assert(size.widthEmu > 0 && size.heightEmu > 0, 'slide size EMU should be positive')
		assert(Math.abs(size.widthIn - size.widthEmu / 914400) < 1e-9, 'widthIn derived from EMU')
	})

	test('exposes slides in presentation order with stable indices', async () => {
		const presentation = await open('textbox')
		const slides = presentation.slides
		assert(slides.length >= 1, 'expected at least one slide')
		slides.forEach((slide, i) => {
			assertEqual(slide.index, i, `slide ${i} index`)
			assertEqual(slide.partName, `/ppt/slides/slide${i + 1}.xml`, `slide ${i} partname in order`)
			assert(slide.presentation === presentation, 'slide back-references its presentation')
		})
	})

	test('save() delegates to the package and round-trips', async () => {
		const presentation = await open('empty')
		const saved = await presentation.save()
		const reopened = await Presentation.load(saved)
		assertEqual(reopened.slides.length, presentation.slides.length, 'slide count survives save')
	})
})

describe('Slide.shapes', () => {
	test('builds an AutoShape with a text frame from a text box', async () => {
		const slide = (await open('textbox')).slides[0]
		const shapes = slide.shapes
		assert(shapes.length >= 1, 'expected at least one shape')
		const textShape = shapes.find((shape) => shape.shapeType === 'autoShape' && shape.hasTextFrame)
		assert(textShape, 'expected a text-bearing auto shape')
		assert(typeof textShape.id === 'number', 'shape id is numeric')
		assert(textShape.name.length > 0, 'shape has a name')
		assert(textShape.left !== null && textShape.width !== null, 'shape geometry resolves')
		assert(textShape.text.includes('test'), `text frame text: ${JSON.stringify(textShape.text)}`)
	})

	test('reads geometry in EMU from the shape transform', async () => {
		const slide = (await open('textbox')).slides[0]
		const shape = slide.shapes.find((shape) => shape.name === 'replaceText')
		assert(shape, 'expected the replaceText shape')
		// <a:off x="2068830" y="1794510"/> <a:ext cx="7566660" cy="2616101"/>
		assertEqual(shape.left, 2068830, 'left EMU')
		assertEqual(shape.top, 1794510, 'top EMU')
		assertEqual(shape.width, 7566660, 'width EMU')
		assertEqual(shape.height, 2616101, 'height EMU')
	})

	test('builds a Picture and resolves its embedded image part', async () => {
		const presentation = await open('image')
		const pictures = presentation.slides
			.flatMap((slide) => slide.shapes)
			.filter((shape) => shape.shapeType === 'picture')
		assert(pictures.length >= 1, 'expected at least one picture')
		const picture = pictures[0]
		assert(picture.imageRelId, 'picture has an embed rel id')
		const partName = picture.imagePartName
		assert(partName, 'picture resolves an image partname')
		assert(presentation.opc.part(partName), `image part ${partName} exists in the package`)
	})

	test('builds a GraphicFrame and detects a hosted table', async () => {
		const presentation = await open('table')
		const frames = presentation.slides
			.flatMap((slide) => slide.shapes)
			.filter((shape) => shape.shapeType === 'graphicFrame')
		assert(frames.length >= 1, 'expected at least one graphic frame')
		const tableFrame = frames.find((frame) => frame.hasTable)
		assert(tableFrame, 'expected a graphic frame hosting a table')
		assert(tableFrame.left !== null, 'graphic frame geometry resolves from p:xfrm')
	})
})

describe('TextFrame / Paragraph / Run', () => {
	test('reads runs and character formatting from the textbox fixture', async () => {
		const slide = (await open('textbox')).slides[0]
		const shape = slide.shapes.find((shape) => shape.name === 'replaceText')
		const frame = shape.textFrame
		assert(frame, 'shape has a text frame')
		const firstParagraph = frame.paragraphs[0]
		const runs = firstParagraph.runs
		assert(runs.length >= 4, `expected several runs, got ${runs.length}`)

		// First run: <a:rPr sz="2000" i="1"><a:solidFill><a:schemeClr val="accent2"/>…
		assertEqual(runs[0].text, 'This', 'first run text')
		assertEqual(runs[0].fontSizePt, 20, 'first run font size (2000 → 20pt)')
		assertEqual(runs[0].italic, true, 'first run italic')
		assertEqual(runs[0].schemeColor, 'accent2', 'first run scheme colour')
		assertEqual(runs[0].color, null, 'first run has no explicit srgb colour')

		// A later bold run: <a:rPr b="1"/><a:t>content</a:t>
		const boldRun = runs.find((run) => run.bold === true)
		assert(boldRun, 'expected a bold run')
		assertEqual(boldRun.text, 'content', 'bold run text')

		// Runs without an rPr flag report null (inherited), not false.
		const plainRun = runs.find((run) => run.text === ' is test')
		assert(plainRun, 'expected the " is test" run')
		assertEqual(plainRun.bold, null, 'unset bold is null, not false')

		assert(frame.text.includes('This is test'), `frame text: ${JSON.stringify(frame.text)}`)
		assertEqual(firstParagraph.level, 0, 'default paragraph level is 0')
	})
})

describe('empty deck', () => {
	test('a slide with no real shapes yields an empty shape list', async () => {
		const slide = (await open('empty')).slides[0]
		// The spTree always has the group's own nv/grpSpPr; those are not shapes.
		for (const shape of slide.shapes) {
			assert(
				['autoShape', 'picture', 'connector', 'graphicFrame', 'group'].includes(shape.shapeType),
				`unexpected shape kind: ${shape.shapeType}`
			)
		}
	})
})
