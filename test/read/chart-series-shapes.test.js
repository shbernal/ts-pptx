// The chart caches that are not a `c:val` beside a flat `c:cat`: a scatter's and a bubble's
// `c:xVal` / `c:yVal` / `c:bubbleSize`, categories on more than one level, and the data labels a
// pie keeps on its series. The read model used to return nothing for all four.
//
// Each shape is read twice. `chart-series-shapes.pptx` is the ground truth, authored in desktop
// PowerPoint by `fixtures/authoring/author-chart-series-shapes.ps1`; the written cases check that
// this library's writer and reader agree on the same shapes.

import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { ChartType } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual, captureDiagnostics } from '../helpers.js'
import { authorRead } from './authored.js'
import { openFixture, readFixture } from './corpus.js'

/** The chart in the graphic frame named `name`, on any slide. */
function chartNamed(presentation, name) {
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) {
			if (shape.shapeType === 'graphicFrame' && shape.name === name && shape.chart) return shape.chart
		}
	}
	throw new Error(`no chart named ${name}`)
}

/** The first chart of a one-chart deck written with `addChart(data, options)`. */
async function writtenChart(data, options) {
	const { presentation } = await authorRead((pres) => {
		pres.addSlide().addChart(data, { x: 1, y: 1, w: 6, h: 4, ...options })
	})
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) if (shape.shapeType === 'graphicFrame' && shape.chart) return shape.chart
	}
	throw new Error('the written deck has no chart')
}

/** The `chart/point-*` codes reading `read()` raised. */
async function pointWarnings(read) {
	const { result, codes } = await captureDiagnostics(read)
	return { result, codes: codes.filter((code) => code.startsWith('chart/point-')) }
}

describe('Scatter and bubble series read their X, Y and size caches', () => {
	test('a PowerPoint scatter series reads c:xVal and c:yVal, and no c:val', async () => {
		const [series] = chartNamed(await openFixture('chart-series-shapes'), 'scatter-chart').series
		assertEqual(series.name, 'Y', 'the series name')
		expect(series.xValues).toEqual([1, 2, 3, 4])
		expect(series.yValues).toEqual([2.5, 4, 3.5, 6])
		expect(series.values).toEqual([])
		expect(series.bubbleSizes).toEqual([])
		assertEqual(series.xLabels, null, 'numeric X values are not labels')
	})

	test('a PowerPoint scatter against a column holding text reads X labels, and no X values', async () => {
		// One text cell ("2026e") makes PowerPoint cache the whole X column as strings, the years with
		// it, and plot the points at X = 1..4: the slide renders pixel-identical to slide 1's scatter,
		// which has the same Y values against X = 1..4. So no label is a coordinate, not even "2023".
		const [series] = chartNamed(await openFixture('chart-series-shapes'), 'scatter-text-x-chart').series
		expect(series.xLabels).toEqual(['2023', '2024', '2025', '2026e'])
		expect(series.xValues).toEqual([null, null, null, null])
		expect(series.yValues).toEqual([2.5, 4, 3.5, 6])
	})

	test('X labels read from a string literal, and from the leaf level of a multi-level cache', async () => {
		// Neither is PowerPoint output. `c:xVal` takes every data-source form `c:cat` does, so both
		// are the fixture's string cache respelled.
		const buf = await readFixture('chart-series-shapes')
		const labelsAfter = async (respell) => {
			const zip = await JSZip.loadAsync(buf)
			const part = 'ppt/charts/chart5.xml'
			const xml = await zip.file(part).async('string')
			const cache =
				/<c:xVal><c:strRef><c:f>[^<]*<\/c:f><c:strCache><c:ptCount val="4"\/>([\s\S]*?)<\/c:strCache><\/c:strRef><\/c:xVal>/
			assert(cache.test(xml), 'slide 5 caches its X labels as a string reference')
			zip.file(
				part,
				xml.replace(cache, (_match, points) => `<c:xVal>${respell(points)}</c:xVal>`)
			)
			const presentation = await Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
			const [series] = chartNamed(presentation, 'scatter-text-x-chart').series
			return { labels: series.xLabels, values: series.xValues }
		}

		const literal = await labelsAfter((points) => `<c:strLit><c:ptCount val="4"/>${points}</c:strLit>`)
		expect(literal).toEqual({ labels: ['2023', '2024', '2025', '2026e'], values: [null, null, null, null] })

		const outer = '<c:lvl><c:pt idx="0"><c:v>Actual</c:v></c:pt><c:pt idx="3"><c:v>Estimate</c:v></c:pt></c:lvl>'
		const levels = await labelsAfter(
			(points) =>
				`<c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="4"/><c:lvl>${points}</c:lvl>${outer}</c:multiLvlStrCache></c:multiLvlStrRef>`
		)
		expect(levels).toEqual({ labels: ['2023', '2024', '2025', '2026e'], values: [null, null, null, null] })
	})

	test('a PowerPoint bubble series reads its sizes beside X and Y', async () => {
		const [series] = chartNamed(await openFixture('chart-series-shapes'), 'bubble-chart').series
		expect(series.xValues).toEqual([1, 2, 3, 4])
		expect(series.yValues).toEqual([2, 4, 3, 6])
		expect(series.bubbleSizes).toEqual([5, 10, 7, 12])
	})

	test('a PowerPoint bubble against a column holding text reads X labels, and no X values', async () => {
		// The bubble takes the same `c:xVal` as a scatter, and renders pixel-identical to slide 2's
		// bubble, which has the same Y values and sizes against X = 1..4.
		const [series] = chartNamed(await openFixture('chart-series-shapes'), 'bubble-text-x-chart').series
		expect(series.xLabels).toEqual(['2023', '2024', '2025', '2026e'])
		expect(series.xValues).toEqual([null, null, null, null])
		expect(series.yValues).toEqual([2, 4, 3, 6])
		expect(series.bubbleSizes).toEqual([5, 10, 7, 12])
	})

	test('a PowerPoint scatter or bubble with no X column reads no X values and no X labels', async () => {
		// Neither series has a `c:xVal`. Each renders pixel-identical to the same chart against X = 1..4.
		const presentation = await openFixture('chart-series-shapes')
		const [scatter] = chartNamed(presentation, 'scatter-no-x-chart').series
		const [bubble] = chartNamed(presentation, 'bubble-no-x-chart').series
		expect([scatter.xValues, scatter.xLabels, scatter.yValues]).toEqual([[], null, [2.5, 4, 3.5, 6]])
		expect([bubble.xValues, bubble.xLabels, bubble.yValues, bubble.bubbleSizes]).toEqual([
			[],
			null,
			[2, 4, 3, 6],
			[5, 10, 7, 12],
		])
	})

	test('a written scatter reads every series against the X row it was given', async () => {
		const chart = await writtenChart(
			[
				{ name: 'X', values: [1, 2, 3] },
				{ name: 'Y1', values: [4, 5, 6] },
				{ name: 'Y2', values: [7, 8, 9] },
			],
			{ type: ChartType.scatter }
		)
		expect(chart.series.map((series) => series.name)).toEqual(['Y1', 'Y2'])
		expect(chart.series.map((series) => series.xValues)).toEqual([
			[1, 2, 3],
			[1, 2, 3],
		])
		expect(chart.series.map((series) => series.yValues)).toEqual([
			[4, 5, 6],
			[7, 8, 9],
		])
	})

	test('a written bubble reads its sizes, and its series carry no label block of their own', async () => {
		const chart = await writtenChart(
			[
				{ name: 'X', values: [1, 2, 3] },
				{ name: 'Y', values: [4, 5, 6], sizes: [7, 8, 9] },
			],
			{ type: ChartType.bubble }
		)
		const [series] = chart.series
		expect(series.yValues).toEqual([4, 5, 6])
		expect(series.bubbleSizes).toEqual([7, 8, 9])
		assertEqual(series.dataLabels, null, 'the bubble writer puts its labels on the group')
	})
})

describe('Multi-level categories read level by level', () => {
	test('a PowerPoint two-level axis reads leaf first, with the outer level sparse, and warns about nothing', async () => {
		const { result, codes } = await pointWarnings(async () => {
			const chart = chartNamed(await openFixture('chart-series-shapes'), 'multilevel-bar-chart')
			return { levels: chart.categoryLevels, categories: chart.categories, values: chart.series[0].values }
		})
		expect(result.levels).toEqual([
			['Q1', 'Q2', 'Q1', 'Q2'],
			['North', null, 'South', null],
		])
		expect(result.categories).toEqual(['Q1', 'Q2', 'Q1', 'Q2'])
		expect(result.values).toEqual([10, 14, 8, 12])
		// The cache's `c:ptCount` counts categories. Checked against the sparse outer level it
		// would disagree on every multi-level chart PowerPoint writes.
		expect(codes).toEqual([])
	})

	test('a single-level axis reads as one level', async () => {
		const chart = chartNamed(await openFixture('chart-series-shapes'), 'pie-chart')
		expect(chart.categoryLevels).toEqual([['North', 'South', 'East', 'West']])
		expect(chart.categories).toEqual(['North', 'South', 'East', 'West'])
	})

	test('written multi-level labels read back in the order they were given', async () => {
		const labels = [
			['Q1', 'Q2', 'Q1', 'Q2'],
			['North', '', 'South', ''],
		]
		const { result, codes } = await pointWarnings(async () => {
			const chart = await writtenChart([{ name: 'Revenue', labels, values: [10, 14, 8, 12] }], { type: ChartType.bar })
			return { levels: chart.categoryLevels, categories: chart.categories }
		})
		expect(result.levels).toEqual(labels)
		expect(result.categories).toEqual(labels[0])
		expect(codes).toEqual([])
	})
})

describe('Data labels a series carries itself', () => {
	test("a PowerPoint pie's label flags are on the series, and its group block is all off", async () => {
		const chart = chartNamed(await openFixture('chart-series-shapes'), 'pie-chart')
		expect(chart.series[0].dataLabels).toMatchObject({
			showCategoryName: true,
			showPercent: true,
			showValue: false,
			showSeriesName: false,
			showLegendKey: false,
			position: null,
			numberFormat: null,
		})
		const group = chart.dataLabels
		assert(group, 'PowerPoint also writes a group block for the pie')
		assertEqual(group.showCategoryName, false, 'which shows no category name')
		assertEqual(group.showPercent, false, 'and no percentage')
	})

	test('a written pie reads its label flags from the series', async () => {
		const chart = await writtenChart([{ name: 'Share', labels: ['A', 'B', 'C'], values: [50, 30, 20] }], {
			type: ChartType.pie,
			showPercent: true,
			showLabel: true,
		})
		expect(chart.series[0].dataLabels).toMatchObject({ showPercent: true, showCategoryName: true, showValue: false })
		assertEqual(chart.dataLabels, null, 'the pie writer emits no group block')
	})
})
