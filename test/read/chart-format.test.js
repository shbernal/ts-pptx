// Write→read fidelity for the chart formatting reads added to
// src/read/api/chart.ts: axes (min/max/title/number-format/gridlines), the
// legend (position/overlay), the aggregate data-label block, and per-series
// fill/line. Each is a BLIND SPOT the writer already emits — the axis/legend/
// dataLabel/chartColors chart options — so it is proven by authoring a deck with
// the write API, reading it back through the deep model, and asserting the
// extracted values match what was written. The writer's bytes are the fixture.

import { describe, test } from 'vitest'
import { authorRead, firstChart, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** A bar chart carrying axis / legend / data-label / series-colour formatting. */
function formattedBar(pres) {
	const data = [
		{ name: 'Costs', labels: ['A', 'B', 'C'], values: [10, 20, 30] },
		{ name: 'Revenue', labels: ['A', 'B', 'C'], values: [40, 50, 60] },
	]
	pres.addSlide().addChart(data, {
		type: pres.ChartType.bar,
		x: 1,
		y: 1,
		w: 8,
		h: 4,
		chartColors: ['FF0000', '00FF00'],
		// Value axis
		valAxisMinVal: 0,
		valAxisMaxVal: 100,
		valAxisMajorUnit: 20,
		showValAxisTitle: true,
		valAxisTitle: 'Revenue',
		valAxisLabelFormatCode: '0.0%',
		valGridLine: { style: 'solid', size: 1, color: '888888' },
		// Category axis
		showCatAxisTitle: true,
		catAxisTitle: 'Quarter',
		// Legend
		showLegend: true,
		legendPos: 'b',
		// Data labels (clustered bar accepts ctr/inBase/inEnd)
		showValue: true,
		dataLabelPosition: 'inEnd',
		dataLabelFormatCode: '0%',
	})
}

/** A line chart whose series carry an explicit stroke (width / dash / colour). */
function formattedLine(pres) {
	const data = [{ name: 'Trend', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
	pres.addSlide().addChart(data, {
		type: pres.ChartType.line,
		x: 1,
		y: 1,
		w: 8,
		h: 4,
		chartColors: ['0000FF'],
		lineSize: 3,
		lineDash: 'dash',
	})
}

describe('Chart axes — c:catAx / c:valAx', () => {
	test('the value axis reads min/max/major-unit, title, gridlines, and number format', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		assert(chart, 'authored bar chart is read back')

		const valAxis = chart.valueAxis
		assert(valAxis, 'chart exposes a value axis')
		assertEqual(valAxis.kind, 'val', 'value axis kind')
		assertEqual(valAxis.min, 0, 'authored scale minimum')
		assertEqual(valAxis.max, 100, 'authored scale maximum')
		assertEqual(valAxis.majorUnit, 20, 'authored major unit')
		assertEqual(valAxis.title, 'Revenue', 'authored value-axis title')
		assertEqual(valAxis.majorGridlines, true, 'valGridLine authored → major gridlines present')
		assert(valAxis.numberFormat, 'value axis carries a number format')
		assertEqual(valAxis.numberFormat.formatCode, '0.0%', 'authored value-axis format code')
		assertEqual(valAxis.numberFormat.sourceLinked, false, 'value-axis format is not source-linked')
		assertEqual(valAxis.hidden, false, 'value axis is not hidden')
		assert(valAxis.position, `value axis has a position (${valAxis.position})`)
	})

	test('the value axis exposes id, orientation, and tick/scale accessors', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		const valAxis = chart.valueAxis
		assert(valAxis, 'chart exposes a value axis')

		// `c:axId` is always emitted, so the id resolves to a concrete number.
		assertEqual(typeof valAxis.id, 'number', 'value axis carries a numeric c:axId')
		// Deterministic-null accessors for this deck: a linear, auto-unit axis with
		// no minor gridlines authored.
		assertEqual(valAxis.logBase, null, 'linear scale → no logBase')
		assertEqual(valAxis.minorGridlines, false, 'no minor gridlines authored')
		assertEqual(valAxis.minorUnit, null, 'auto minor unit → null')
		// The writer's exact tick/orientation emission is not pinned here; assert the
		// accessor contract (a string when present, else null).
		for (const [name, value] of [
			['orientation', valAxis.orientation],
			['majorTickMark', valAxis.majorTickMark],
			['minorTickMark', valAxis.minorTickMark],
			['tickLabelPosition', valAxis.tickLabelPosition],
		]) {
			assert(value === null || typeof value === 'string', `${name} is a string or null (${value})`)
		}
	})

	test('the category axis reads its title and its (default) number format', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		const catAxis = chart.categoryAxis
		assert(catAxis, 'chart exposes a category axis')
		assertEqual(catAxis.kind, 'cat', 'category axis kind')
		assertEqual(catAxis.title, 'Quarter', 'authored category-axis title')
		assertEqual(catAxis.numberFormat.formatCode, 'General', 'category axis defaults to General')
		assertEqual(catAxis.numberFormat.sourceLinked, true, 'category-axis format is source-linked')
	})

	test('axes lists both plot axes in document order', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		assertEqual(chart.axes.map((a) => a.kind).join(','), 'cat,val', 'category axis precedes value axis')
	})
})

describe('Chart legend — c:legend', () => {
	test('a shown legend reads its position and overlay flag', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		assert(chart.legend, 'chart exposes a legend')
		assertEqual(chart.legend.position, 'b', 'authored legend position')
		assertEqual(chart.legend.overlay, false, 'writer emits overlay=0')
	})

	test('a chart authored without a legend reports null', async () => {
		const chart = firstChart(
			(
				await authorRead((pres) => {
					pres.addSlide().addChart([{ name: 'S', labels: ['A'], values: [1] }], {
						type: pres.ChartType.bar,
						x: 1,
						y: 1,
						w: 6,
						showLegend: false,
					})
				})
			).presentation
		)
		assertEqual(chart.legend, null, 'no c:legend → null')
	})
})

describe('Chart data labels — c:dLbls', () => {
	test('the aggregate data-label block reads its show flags, position, and format', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		const labels = chart.dataLabels
		assert(labels, 'chart exposes an aggregate data-label block')
		assertEqual(labels.showValue, true, 'showValue authored true')
		assertEqual(labels.showSeriesName, false, 'showSerName defaults false')
		assertEqual(labels.showCategoryName, false, 'showCatName defaults false')
		assertEqual(labels.showPercent, false, 'showPercent defaults false')
		assertEqual(labels.position, 'inEnd', 'authored data-label position')
		assertEqual(labels.numberFormat.formatCode, '0%', 'authored data-label format code')
	})
})

describe('Chart series appearance — c:ser/c:spPr', () => {
	test('bar series read their authored fill colours', async () => {
		const chart = firstChart((await authorRead(formattedBar)).presentation)
		const series = chart.series
		assertEqual(series.length, 2, 'two series')
		assert(series[0].fill, 'first series has a fill')
		assertEqual(series[0].fill.color, 'FF0000', 'first series colour from chartColors[0]')
		assertEqual(series[0].fill.noFill, false, 'a coloured series is not noFill')
		assertEqual(series[1].fill.color, '00FF00', 'second series colour from chartColors[1]')
		assertEqual(series[0].line, null, 'bar series carry no stroke by default')
	})

	test('line series read their authored stroke width / dash / colour', async () => {
		const chart = firstChart((await authorRead(formattedLine)).presentation)
		const series = chart.series[0]
		assert(series.line, 'line series has a stroke')
		assertEqual(series.line.widthPt, 3, 'authored line width in points')
		assertEqual(series.line.dash, 'dash', 'authored line dash')
		assertEqual(series.line.color, '0000FF', 'authored line colour')
		assertEqual(series.line.noFill, false, 'a drawn line is not noFill')
	})
})

describe('Chart formatting — schema validity', () => {
	test.skipIf(!validatorInstalled)('the authored formatted decks are schema-valid', async () => {
		assertEqual((await schemaErrors((await authorRead(formattedBar)).buf)).length, 0, 'formatted bar deck validates')
		assertEqual((await schemaErrors((await authorRead(formattedLine)).buf)).length, 0, 'formatted line deck validates')
	})
})
