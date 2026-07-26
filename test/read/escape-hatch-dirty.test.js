// Dirty-obligation tests for the `element_` escape hatch (src/read/api/*.ts).
//
// Contract under test: `element_` hands out the live DOM node, and a mutation
// through it is a **no-op on save** unless the caller marks the owning part
// dirty. That is deliberate — it is what keeps untouched parts byte-identical
// for read-only consumers — but it is the one failure mode where a silent
// wrong answer is possible, so it is specified here rather than left to
// surprise. The companion assertion is that every class exposing `element_`
// also exposes a public `markDirty()` that makes the edit stick.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

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
	return a.length === b.length && a.every((value, index) => value === b[index])
}

/**
 * Load `fixture`, run `mutate(presentation)` (which is expected to reach the DOM
 * only through `element_`), save, and report both the saved bytes and whether
 * every part body came back byte-identical to the input.
 */
async function saveAfter(fixture, mutate) {
	const input = await readFile(fixturePath(fixture))
	const presentation = await Presentation.load(input)
	mutate(presentation)
	const saved = await presentation.save()
	const inputBodies = await partBodies(input)
	const outputBodies = await partBodies(saved)
	const changed = []
	for (const [name, body] of inputBodies) {
		if (!bytesEqual(body, outputBodies.get(name))) changed.push(name)
	}
	return { saved, changed }
}

const SLIDE1 = 'ppt/slides/slide1.xml'

/** The `textbox` fixture's first shape with a text frame. */
function textShape(presentation) {
	return presentation.slides[0].shapes.find((shape) => shape.textFrame)
}

describe('element_ without markDirty()', () => {
	test('a shape rotation written through element_ is dropped on save', async () => {
		const { saved, changed } = await saveAfter('textbox', (presentation) => {
			textShape(presentation).element_.setAttribute('rot', '5400000')
		})
		assertEqual(changed.length, 0, 'no part should be reserialized without markDirty()')
		const reopened = await Presentation.load(saved)
		assertEqual(textShape(reopened).element_.getAttribute('rot'), null, 'the unmarked edit does not reach the output')
	})

	test('a slide attribute written through element_ is dropped on save', async () => {
		const { changed } = await saveAfter('textbox', (presentation) => {
			presentation.slides[0].element_.setAttribute('show', '0')
		})
		assertEqual(changed.length, 0, 'no part should be reserialized without markDirty()')
	})
})

describe('element_ with markDirty()', () => {
	test('Shape.markDirty() makes an element_ edit survive the round-trip', async () => {
		const { saved, changed } = await saveAfter('textbox', (presentation) => {
			const shape = textShape(presentation)
			shape.element_.setAttribute('rot', '5400000')
			shape.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		assertEqual(textShape(reopened).element_.getAttribute('rot'), '5400000', 'the marked edit reaches the output')
	})

	test('Slide.markDirty() makes an element_ edit survive the round-trip', async () => {
		const { saved, changed } = await saveAfter('textbox', (presentation) => {
			const slide = presentation.slides[0]
			slide.element_.setAttribute('show', '0')
			slide.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		assertEqual(reopened.slides[0].hidden, true, 'the marked edit reaches the output')
	})

	test('TextFrame/Paragraph/Run each expose a markDirty() reaching the same part', async () => {
		const { saved, changed } = await saveAfter('textbox', (presentation) => {
			const frame = textShape(presentation).textFrame
			const paragraph = frame.paragraphs[0]
			const run = paragraph.runs[0]
			// Reach the run's `a:t` through the hatch, then mark from the *run* —
			// the innermost rung of the ladder — and assert it reaches the slide part.
			const t = run.element_.getElementsByTagName('a:t')[0]
			t.textContent = 'HATCHED'
			run.markDirty()
			assert(typeof paragraph.markDirty === 'function', 'Paragraph exposes markDirty()')
			assert(typeof frame.markDirty === 'function', 'TextFrame exposes markDirty()')
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		assertEqual(
			textShape(reopened).textFrame.paragraphs[0].runs[0].text,
			'HATCHED',
			'the marked edit reaches the output'
		)
	})

	test('Table/TableRow/TableCell each expose a markDirty() reaching the same part', async () => {
		const { saved, changed } = await saveAfter('table', (presentation) => {
			const table = presentation.slides
				.flatMap((slide) => slide.shapes)
				.find((shape) => shape.shapeType === 'graphicFrame' && shape.table).table
			const row = table.rows[0]
			const cell = row.cells[0]
			assert(typeof table.markDirty === 'function', 'Table exposes markDirty()')
			assert(typeof row.markDirty === 'function', 'TableRow exposes markDirty()')
			cell.element_.setAttribute('marL', '91440')
			cell.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		const cell = reopened.slides
			.flatMap((slide) => slide.shapes)
			.find((shape) => shape.shapeType === 'graphicFrame' && shape.table).table.rows[0].cells[0]
		assertEqual(cell.element_.getAttribute('marL'), '91440', 'the marked edit reaches the output')
	})

	test('Chart.markDirty() reserializes the chart part, not the slide', async () => {
		const { changed } = await saveAfter('bar-chart-data-labels', (presentation) => {
			const chart = presentation.slides
				.flatMap((slide) => slide.shapes)
				.find((shape) => shape.shapeType === 'graphicFrame' && shape.chart).chart
			chart.element_.setAttribute('xmlns:tspptx', 'urn:test')
			chart.markDirty()
		})
		assertEqual(changed.length, 1, 'exactly one part should be reserialized')
		assert(changed[0].startsWith('ppt/charts/chart'), `expected a chart part, got ${changed[0]}`)
	})
})
