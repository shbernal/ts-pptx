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
import { describe, test } from 'vitest'
import { ChartType } from '../../dist/node.js'
import { Presentation, isGraphicFrame } from '../../dist/read.js'
import { authorRead } from './authored.js'
import { bytesEqual, assert, assertEqual, partBodies } from '../helpers.js'
import { fixturePath } from './corpus.js'

/**
 * Load `input`, run `mutate(presentation)` (which is expected to reach the DOM
 * only through `element_`), save, and report both the saved bytes and which part
 * bodies came back differing from the input.
 */
async function saveAfterBytes(input, mutate) {
	const presentation = await Presentation.load(input)
	await mutate(presentation)
	const saved = await presentation.save()
	const inputBodies = await partBodies(input)
	const outputBodies = await partBodies(saved)
	const changed = []
	for (const [name, body] of inputBodies) {
		const after = outputBodies.get(name)
		if (!after || !bytesEqual(body, after)) changed.push(name)
	}
	return { saved, changed }
}

/** {@link saveAfterBytes} against a `fixtures/<name>.pptx` on disk. */
async function saveAfter(fixture, mutate) {
	return saveAfterBytes(await readFile(fixturePath(fixture)), mutate)
}

/**
 * {@link saveAfterBytes} against a deck authored in memory with the write API —
 * for the read models no checked-in fixture exercises (chartEx, master/layout
 * chrome, notes placeholders).
 */
async function saveAfterAuthored(build, mutate) {
	const { buf } = await authorRead(build)
	return saveAfterBytes(buf, mutate)
}

const SLIDE1 = 'ppt/slides/slide1.xml'

/** The first `graphicFrame` on any slide whose `pick` yields a model, or null. */
function firstFramed(presentation, pick) {
	for (const shape of presentation.slides.flatMap((slide) => slide.shapes)) {
		if (shape.shapeType !== 'graphicFrame') continue
		const model = pick(shape)
		if (model) return model
	}
	return null
}

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
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		assertEqual(
			textShape(reopened).textFrame.paragraphs[0].runs[0].text,
			'HATCHED',
			'the marked edit reaches the output'
		)
	})

	// The outer rungs of the same ladder. Asserting `typeof x.markDirty === 'function'`
	// would prove only that a method exists — these call it and assert the edit lands,
	// which is the property the contract actually promises.
	test('TextFrame.markDirty() reaches the slide part from the outer rung', async () => {
		const { saved, changed } = await saveAfter('textbox', (presentation) => {
			const frame = textShape(presentation).textFrame
			frame.element_.getElementsByTagName('a:t')[0].textContent = 'FROM-FRAME'
			frame.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		// The fixture's frame holds several runs; the hatch edited the first `a:t`,
		// so assert that run rather than the flattened whole-frame text.
		assertEqual(
			textShape(reopened).textFrame.paragraphs[0].runs[0].text,
			'FROM-FRAME',
			'the marked edit reaches the output'
		)
	})

	test('Paragraph.markDirty() reaches the slide part from the middle rung', async () => {
		const { saved, changed } = await saveAfter('textbox', (presentation) => {
			const paragraph = textShape(presentation).textFrame.paragraphs[0]
			paragraph.element_.getElementsByTagName('a:t')[0].textContent = 'FROM-PARA'
			paragraph.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		assertEqual(
			textShape(reopened).textFrame.paragraphs[0].runs[0].text,
			'FROM-PARA',
			'the marked edit reaches the output'
		)
	})

	test('Table/TableRow/TableCell each expose a markDirty() reaching the same part', async () => {
		const { saved, changed } = await saveAfter('table', (presentation) => {
			const table = firstFramed(presentation, (shape) => shape.table)
			const cell = table.rows[0].cells[0]
			cell.element_.setAttribute('marL', '91440')
			cell.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		const cell = reopened.slides
			.flatMap((slide) => slide.shapes)
			.filter(isGraphicFrame)
			.find((shape) => shape.table).table.rows[0].cells[0]
		assertEqual(cell.element_.getAttribute('marL'), '91440', 'the marked edit reaches the output')
	})

	test('Table.markDirty() reaches the slide part from the outer rung', async () => {
		const { saved, changed } = await saveAfter('table', (presentation) => {
			const table = firstFramed(presentation, (shape) => shape.table)
			table.element_.getElementsByTagName('a:tc')[0].setAttribute('marR', '45720')
			table.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		const cell = firstFramed(reopened, (shape) => shape.table).rows[0].cells[0]
		assertEqual(cell.element_.getAttribute('marR'), '45720', 'the marked edit reaches the output')
	})

	test('TableRow.markDirty() reaches the slide part from the middle rung', async () => {
		const { saved, changed } = await saveAfter('table', (presentation) => {
			const row = firstFramed(presentation, (shape) => shape.table).rows[0]
			row.element_.setAttribute('h', '742950')
			row.markDirty()
		})
		assertEqual(changed.join(), SLIDE1, 'only the owning slide part should be reserialized')
		const reopened = await Presentation.load(saved)
		const row = firstFramed(reopened, (shape) => shape.table).rows[0]
		assertEqual(row.element_.getAttribute('h'), '742950', 'the marked edit reaches the output')
	})

	test('Chart.markDirty() reserializes the chart part, not the slide', async () => {
		const { changed } = await saveAfter('bar-chart-data-labels', (presentation) => {
			const chart = firstFramed(presentation, (shape) => shape.chart)
			chart.element_.setAttribute('xmlns:tspptx', 'urn:test')
			chart.markDirty()
		})
		assertEqual(changed.length, 1, 'exactly one part should be reserialized')
		assert(changed[0].startsWith('ppt/charts/chart'), `expected a chart part, got ${changed[0]}`)
	})

	// ChartAxis / ChartSeries hang off the chart part, not the slide: their
	// markDirty() has to walk to the chart part they were constructed with (the
	// second constructor argument), so mark from each and assert the same
	// single-part reserialization the whole-chart case gets.
	test('ChartAxis.markDirty() reserializes the owning chart part', async () => {
		const { saved, changed } = await saveAfter('bar-chart-data-labels', (presentation) => {
			const axis = firstFramed(presentation, (shape) => shape.chart).axes[0]
			axis.element_.getElementsByTagName('c:axId')[0].setAttribute('val', '191919')
			axis.markDirty()
		})
		assertEqual(changed.length, 1, 'exactly one part should be reserialized')
		assert(changed[0].startsWith('ppt/charts/chart'), `expected a chart part, got ${changed[0]}`)
		const reopened = await Presentation.load(saved)
		assertEqual(firstFramed(reopened, (shape) => shape.chart).axes[0].id, 191919, 'the marked edit reaches the output')
	})

	test('ChartSeries.markDirty() reserializes the owning chart part', async () => {
		const { saved, changed } = await saveAfter('bar-chart-data-labels', (presentation) => {
			const series = firstFramed(presentation, (shape) => shape.chart).series[0]
			series.element_.getElementsByTagName('c:idx')[0].setAttribute('val', '7')
			series.markDirty()
		})
		assertEqual(changed.length, 1, 'exactly one part should be reserialized')
		assert(changed[0].startsWith('ppt/charts/chart'), `expected a chart part, got ${changed[0]}`)
		const reopened = await Presentation.load(saved)
		assertEqual(firstFramed(reopened, (shape) => shape.chart).series[0].index, 7, 'the marked edit reaches the output')
	})
})

// chartEx has no checked-in fixture — the writer emits every `cx:` piece, so these
// author a waterfall in memory and round-trip the writer's own bytes (the same
// harness test/read/chartex-read.test.js uses for the getters).
function authorWaterfall(pres) {
	pres.addSlide().addChart([{ name: 'Cash Flow', labels: ['Start', 'Q1', 'End'], values: [100, 40, 190] }], {
		type: ChartType.waterfall,
		x: 1,
		y: 1,
		w: 6,
		h: 4,
	})
}

/** The chartEx part a `changed` list is expected to hold exactly one of. */
function assertOneChartPart(changed) {
	assertEqual(changed.length, 1, `exactly one part should be reserialized, got ${changed.join() || '(none)'}`)
	assert(changed[0].startsWith('ppt/charts/'), `expected a chart part, got ${changed[0]}`)
}

describe('chartEx element_ / markDirty()', () => {
	test('a chartEx edit through element_ is dropped without markDirty()', async () => {
		const { changed } = await saveAfterAuthored(authorWaterfall, (presentation) => {
			const cx = firstFramed(presentation, (shape) => shape.chartEx)
			cx.element_.setAttribute('xmlns:tspptx', 'urn:test')
		})
		assertEqual(changed.length, 0, 'no part should be reserialized without markDirty()')
	})

	test('ChartEx.markDirty() reserializes the chartEx part, not the slide', async () => {
		const { changed } = await saveAfterAuthored(authorWaterfall, (presentation) => {
			const cx = firstFramed(presentation, (shape) => shape.chartEx)
			cx.element_.setAttribute('xmlns:tspptx', 'urn:test')
			cx.markDirty()
		})
		assertOneChartPart(changed)
	})

	test('ChartExSeries.markDirty() reserializes the owning chartEx part', async () => {
		const { saved, changed } = await saveAfterAuthored(authorWaterfall, (presentation) => {
			const series = firstFramed(presentation, (shape) => shape.chartEx).series[0]
			assertEqual(series.layoutId, 'waterfall', 'the series layout token reads back before the edit')
			series.element_.setAttribute('uniqueId', '{ESCAPE-HATCH}')
			series.markDirty()
		})
		assertOneChartPart(changed)
		const reopened = await Presentation.load(saved)
		assertEqual(
			firstFramed(reopened, (shape) => shape.chartEx).series[0].uniqueId,
			'{ESCAPE-HATCH}',
			'the marked edit reaches the output'
		)
	})

	test('ChartExAxis.markDirty() reserializes the owning chartEx part', async () => {
		const { saved, changed } = await saveAfterAuthored(authorWaterfall, (presentation) => {
			const axis = firstFramed(presentation, (shape) => shape.chartEx).axes[0]
			// A waterfall's category axis carries no cx:valScaling, so the value-scale
			// getters read null rather than throwing — pin that alongside the hatch.
			assertEqual(axis.min, null, 'a cat axis has no value-scale minimum')
			assertEqual(axis.max, null, 'a cat axis has no value-scale maximum')
			assertEqual(typeof axis.tickLabels, 'boolean', 'tickLabels is a presence flag')
			axis.element_.setAttribute('id', '42')
			axis.markDirty()
		})
		assertOneChartPart(changed)
		const reopened = await Presentation.load(saved)
		assertEqual(firstFramed(reopened, (shape) => shape.chartEx).axes[0].id, 42, 'the marked edit reaches the output')
	})
})

describe('shared-chrome and notes element_ / markDirty()', () => {
	/** A deck with a defined master (so master/layout placeholders exist) and notes. */
	function authorChrome(pres) {
		pres.theme = { headFontFace: 'Georgia' }
		pres.defineSlideMaster({ title: 'BRANDED', slideNumber: { x: 0.5, y: 7.0 } })
		pres.addSlide({ masterTitle: 'BRANDED' }).addNotes('a note')
	}

	test('Placeholder.markDirty() reserializes the owning master part', async () => {
		const { saved, changed } = await saveAfterAuthored(authorChrome, (presentation) => {
			const sldNum = presentation.slides[0].master.placeholders.find((ph) => ph.type === 'sldNum')
			assert(sldNum, 'the master has a slide-number placeholder')
			assert(sldNum.idx !== undefined, 'the placeholder exposes its p:ph idx')
			assert(typeof sldNum.id === 'number' || sldNum.id === null, 'the placeholder exposes its drawing id')
			assert(sldNum.textFrame !== undefined, 'the placeholder exposes its text frame')
			sldNum.element_.getElementsByTagName('p:cNvPr')[0].setAttribute('name', 'HATCHED-PH')
			sldNum.markDirty()
		})
		assertEqual(changed.join(), 'ppt/slideMasters/slideMaster1.xml', 'only the owning master part is reserialized')
		const reopened = await Presentation.load(saved)
		const sldNum = reopened.slides[0].master.placeholders.find((ph) => ph.type === 'sldNum')
		assertEqual(sldNum.name, 'HATCHED-PH', 'the marked edit reaches the output')
	})

	test('Theme.markDirty() reserializes the theme part', async () => {
		const { saved, changed } = await saveAfterAuthored(authorChrome, (presentation) => {
			const theme = presentation.slides[0].theme
			assert(theme, 'the slide resolves its theme')
			theme.element_.setAttribute('name', 'Hatched Theme')
			theme.markDirty()
		})
		assertEqual(changed.length, 1, `exactly one part should be reserialized, got ${changed.join()}`)
		assert(changed[0].startsWith('ppt/theme/'), `expected a theme part, got ${changed[0]}`)
		const reopened = await Presentation.load(saved)
		assertEqual(reopened.slides[0].theme.name, 'Hatched Theme', 'the marked edit reaches the output')
	})

	test('NotesPlaceholder.markDirty() reserializes the owning notes-slide part', async () => {
		const { saved, changed } = await saveAfterAuthored(authorChrome, (presentation) => {
			const body = presentation.slides[0].notesSlide.body
			assert(body, 'the notes slide has a body placeholder')
			assert(typeof body.name === 'string', 'the notes placeholder exposes its shape name')
			assert(typeof body.id === 'number' || body.id === null, 'the notes placeholder exposes its drawing id')
			body.element_.getElementsByTagName('a:t')[0].textContent = 'HATCHED NOTE'
			body.markDirty()
		})
		assertEqual(changed.length, 1, `exactly one part should be reserialized, got ${changed.join()}`)
		assert(changed[0].startsWith('ppt/notesSlides/'), `expected a notes-slide part, got ${changed[0]}`)
		const reopened = await Presentation.load(saved)
		assertEqual(reopened.slides[0].notesSlide.body.text, 'HATCHED NOTE', 'the marked edit reaches the output')
	})

	test('ResolvedTableStyle.markDirty() reserializes tableStyles.xml', async () => {
		const { saved, changed } = await saveAfter('table-styles', (presentation) => {
			const style = firstFramed(presentation, (shape) => shape.table).resolvedStyle
			assert(style, 'the table resolves its style entry')
			style.element_.setAttribute('styleName', 'Hatched Style')
			style.markDirty()
		})
		assertEqual(changed.join(), 'ppt/tableStyles.xml', 'only the style part is reserialized, not the slide')
		const reopened = await Presentation.load(saved)
		assertEqual(
			firstFramed(reopened, (shape) => shape.table).resolvedStyle.name,
			'Hatched Style',
			'the marked edit reaches the output'
		)
	})
})
