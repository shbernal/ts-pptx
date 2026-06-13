// Phase 4 chart-read tests for `pptxgenjs/read` (src/read/api/chart.ts).
//
// Contract under test: a chart graphic frame resolves its chart part via the
// slide relationships and exposes the chart type, title, series, and cached
// category/value data read from the live DOM. Read-only: opening and reading a
// chart dirties nothing, so save() stays byte-identical.

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

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

/** The first chart on any slide of the fixture. */
function firstChart(presentation) {
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) {
			if (shape.shapeType === 'graphicFrame' && shape.chart) return shape.chart
		}
	}
	return null
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

describe('Chart read model', () => {
	test('resolves the chart part and reads type + title', async () => {
		const chart = firstChart(await open('mixed'))
		assert(chart, 'mixed.pptx has a chart')
		assert(chart.partName.startsWith('/ppt/charts/'), `chart partname: ${chart.partName}`)
		assertEqual(chart.chartType, 'line', 'chart type')
		assert(chart.chartTypes.includes('line'), 'chartTypes lists line')
		assertEqual(chart.title, 'Breakeven Point', 'chart title (rich text concatenated)')
	})

	test('reads series names, indices, and cached values', async () => {
		const chart = firstChart(await open('mixed'))
		const series = chart.series
		assertEqual(series.length, 2, 'two series')
		assertEqual(series[0].name, 'Costs', 'first series name')
		assertEqual(series[1].name, 'Revenue', 'second series name')
		assertEqual(series[0].index, 0, 'first series index')

		const values = series[0].values
		assertEqual(values.length, 16, 'first series has 16 cached values')
		assertEqual(values[0], 360000, 'first cached value')
		assertEqual(values[15], 435000, 'last cached value')
		assert(
			values.every((v) => typeof v === 'number'),
			'all cached values parsed as numbers'
		)
	})

	test('reads category labels from the first series cache', async () => {
		const chart = firstChart(await open('mixed'))
		const categories = chart.categories
		assertEqual(categories.length, 16, '16 categories')
		assertEqual(categories[0], '3200', 'first category as written')
		assertEqual(chart.categories.length, chart.series[0].values.length, 'categories align with values')
	})

	test('a non-chart graphic frame has a null chart', async () => {
		const slides = (await open('mixed')).slides
		const tableFrame = slides
			.flatMap((s) => s.shapes)
			.find((shape) => shape.shapeType === 'graphicFrame' && shape.hasTable)
		assert(tableFrame, 'mixed.pptx has a table graphic frame')
		assertEqual(tableFrame.chart, null, 'table frame has no chart')
	})

	test('reading a chart dirties nothing — save stays byte-identical', async () => {
		const input = await readFile(fixturePath('mixed'))
		const presentation = await Presentation.load(input)
		const chart = firstChart(presentation)
		// Touch every read accessor.
		void [chart.chartType, chart.title, chart.categories, chart.series.map((s) => [s.name, s.index, s.values])]

		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())
		for (const [name, body] of inputBodies) {
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be byte-identical after a read-only open`)
		}
	})
})
