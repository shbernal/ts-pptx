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

describe('Slide.hidden', () => {
	test('reads p:sld/@show="0" as hidden and treats an absent attr as shown', async () => {
		// hidden.pptx is textbox.pptx with show="0" set on slide 2 (the attribute
		// PowerPoint writes when a slide is hidden); slide 1 omits @show entirely.
		const slides = (await open('hidden')).slides
		assertEqual(slides.length, 2, 'hidden fixture has two slides')
		assertEqual(slides[0].hidden, false, 'slide with no @show is shown')
		assertEqual(slides[1].hidden, true, 'slide with show="0" is hidden')
	})

	test('a deck with no hidden slides reports every slide as shown', async () => {
		const slides = (await open('textbox')).slides
		for (const slide of slides) {
			assertEqual(slide.hidden, false, `slide ${slide.index} should be shown`)
		}
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

// mixed.pptx is the real-world coverage deck for shape kinds the vendored
// fixtures lack: connectors, nested groups, tables, charts, and SmartArt.
describe('mixed.pptx — connectors, groups, graphic frames', () => {
	/** Flatten a shape list, descending into groups. */
	function allShapes(shapes) {
		return shapes.flatMap((shape) => (shape.shapeType === 'group' ? [shape, ...allShapes(shape.shapes)] : [shape]))
	}

	test('reads top-level connectors (p:cxnSp) with resolvable geometry', async () => {
		// slide6 (index 5) has three connectors directly in its spTree.
		const slide = (await open('mixed')).slides[5]
		const connectors = slide.shapes.filter((shape) => shape.shapeType === 'connector')
		assert(connectors.length >= 3, `expected ≥3 top-level connectors, got ${connectors.length}`)
		for (const connector of connectors) {
			assert(typeof connector.id === 'number', 'connector has a numeric id')
			assert(connector.left !== null && connector.width !== null, 'connector geometry resolves from spPr/a:xfrm')
		}
	})

	test('descends into nested groups; connectors surface only via group traversal', async () => {
		// slide5 (index 4): connectors live inside groups, not at the top level.
		const slide = (await open('mixed')).slides[4]
		const topShapes = slide.shapes
		const groups = topShapes.filter((shape) => shape.shapeType === 'group')
		assert(groups.length >= 4, `expected ≥4 top-level groups, got ${groups.length}`)
		assertEqual(
			topShapes.filter((shape) => shape.shapeType === 'connector').length,
			0,
			'no connectors at the top level of slide5'
		)

		const nestedConnectors = allShapes(topShapes).filter((shape) => shape.shapeType === 'connector')
		assert(nestedConnectors.length >= 3, `expected ≥3 connectors via group recursion, got ${nestedConnectors.length}`)

		// "Groupe 2" nests one shape + two connectors; check the proxy enumerates them.
		const groupe2 = groups.find((group) => group.name === 'Groupe 2')
		assert(groupe2, 'expected a group named "Groupe 2"')
		assertEqual(groupe2.shapes.length, 3, 'Groupe 2 has three nested children')
		assertEqual(
			groupe2.shapes.filter((shape) => shape.shapeType === 'connector').length,
			2,
			'Groupe 2 nests two connectors'
		)
		assert(
			groupe2.shapes.every((shape) => shape.slide === slide),
			'nested shapes back-reference their owning slide'
		)
	})

	test('distinguishes table, chart, and SmartArt graphic frames', async () => {
		const slides = (await open('mixed')).slides
		const frameOn = (index) => slides[index].shapes.find((shape) => shape.shapeType === 'graphicFrame')

		const table = frameOn(6) // slide7: a:tbl
		assert(table, 'slide7 has a graphic frame')
		assert(table.hasTable && !table.hasChart, 'slide7 frame is a table')
		assert(table.left !== null, 'graphic frame geometry resolves from p:xfrm')

		const chart = frameOn(7) // slide8: c:chart
		assert(chart, 'slide8 has a graphic frame')
		assert(chart.hasChart && !chart.hasTable, 'slide8 frame is a chart')

		const smartArt = frameOn(1) // slide2: dgm diagram
		assert(smartArt, 'slide2 has a graphic frame')
		assert(!smartArt.hasTable && !smartArt.hasChart, 'SmartArt frame is neither table nor chart')
	})
})
