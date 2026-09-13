// The arms of the script converter's chart mapper that no fixture reaches.
//
// Each one is a *loss* the converter declares — an unwritable plot type, a combo flattened to
// one type, a cached blank that has no spelling in `values: number[]` — and a declared loss is
// exactly what the round-trip verifier excludes from comparison. So an arm that stops firing,
// or fires with the wrong construct, is invisible to the round trip by construction: it is the
// mechanism that decides what the round trip is allowed to ignore.
//
// The decks are authored with the write API and then edited in the chart part, because the
// write path cannot produce most of these: it has no unwritable plot type and no blank point.

import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { ChartType } from '../../dist/node.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'
import { readFixture } from './corpus.js'

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]

/** A scatter's data: the X row, then two Y series. */
const XY = [
	{ name: 'X', values: [1, 2, 3] },
	{ name: 'Y1', values: [4, 5, 6] },
	{ name: 'Y2', values: [7, 8, 9] },
]

/** The first chart call in the IR. */
function chartCall(ir) {
	const call = ir.slides.flatMap((slide) => slide.calls).find((one) => one.method === 'addChart')
	assert(call, 'the IR has a chart call')
	return call
}

/** The chart call for the graphic frame named `objectName`. */
function chartNamed(ir, objectName) {
	const call = ir.slides
		.flatMap((slide) => slide.calls)
		.find((one) => one.method === 'addChart' && one.args[1].objectName === objectName)
	assert(call, `the IR has a chart call for ${objectName}`)
	return call
}

/** Apply `rewrite` to every chart part of `buf`, reload, and convert. */
async function irWithChartXml(buf, rewrite) {
	const zip = await JSZip.loadAsync(buf)
	for (const name of Object.keys(zip.files)) {
		if (!/^ppt\/charts\/chart\d+\.xml$/.test(name)) continue
		zip.file(name, rewrite(await zip.file(name).async('string')))
	}
	const reopened = await Presentation.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
	return readModelToIr(reopened)
}

/** Every note construct the IR recorded. */
const constructs = (ir) => ir.fidelity.map((note) => note.construct)

/** A one-chart deck of `type`. */
function chartDeck(type, options = {}) {
	return authorRead((pres) => {
		pres.addSlide().addChart(SERIES, { type, x: 1, y: 1, w: 6, h: 4, ...options })
	})
}

describe('the chart mapper declares what it cannot carry', () => {
	test('a plot type with no write-API counterpart drops the chart, and says so', async () => {
		// `c:ofPieChart` (bar-of-pie / pie-of-pie) is a real ECMA-376 plot this library cannot
		// author. Dropping it silently would leave a slide with a hole in it.
		const { buf } = await chartDeck(ChartType.bar)
		const ir = await irWithChartXml(buf, (xml) => xml.replaceAll('c:barChart', 'c:ofPieChart'))
		assert(constructs(ir).includes('chart.type'), 'the unwritable type is noted; got ' + JSON.stringify(constructs(ir)))
		assertEqual(
			ir.slides[0].calls.filter((call) => call.method === 'addChart').length,
			0,
			'and no chart call is emitted'
		)
	})

	test('a chart whose caches hold no series is dropped once, not twice', async () => {
		// The `chart.data` note sits BELOW the drop guard deliberately: recorded unconditionally,
		// a chart with no cached series emitted two notes -- one saying it was dropped and one
		// saying it was rebuilt -- and the second maps to `['*']`, so it excused every difference
		// on that frame. The widest exclusion in the table, applied to the case it least fits.
		const { buf } = await chartDeck(ChartType.bar)
		const ir = await irWithChartXml(buf, (xml) => xml.replace(/<c:ser>[\s\S]*<\/c:ser>/, ''))
		const seen = constructs(ir)
		assert(seen.includes('chart.data'), 'the empty chart is noted; got ' + JSON.stringify(seen))
		assert(!seen.includes('chart.rebuilt'), 'and not also noted as rebuilt: ' + JSON.stringify(seen))
	})

	test('a combo chart is flattened to one type, and says which types it had', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addChart(
				[
					{ type: ChartType.bar, data: SERIES, options: {} },
					{
						type: ChartType.line,
						data: [{ name: 'S2', labels: ['A', 'B', 'C'], values: [4, 5, 6] }],
						options: { secondaryValAxis: true, secondaryCatAxis: true },
					},
				],
				{ x: 1, y: 1, w: 8, h: 4 }
			)
		})
		const ir = readModelToIr(await Presentation.load(buf))
		assert(constructs(ir).includes('chart.combo'), 'the flattening is noted; got ' + JSON.stringify(constructs(ir)))
		const call = ir.slides[0].calls.find((c) => c.method === 'addChart')
		assert(call, 'and one chart call is still emitted')
	})

	test('a blank cached point becomes 0, and the dip is declared', async () => {
		// `OptsChartData.values` is `number[]` and has no spelling for a gap, so the blank becomes
		// a zero -- which on a line chart draws a dip to the axis where the source showed a break.
		// That is a real difference in the picture, so it is declared rather than quietly made.
		const { buf } = await chartDeck(ChartType.line)
		const ir = await irWithChartXml(buf, (xml) => xml.replace('<c:pt idx="1"><c:v>2</c:v></c:pt>', ''))
		assert(constructs(ir).includes('chart.blanks'), 'the blank is noted; got ' + JSON.stringify(constructs(ir)))
		const call = ir.slides[0].calls.find((c) => c.method === 'addChart')
		const values = call.args[0][0].values
		assertEqual(values[1], 0, 'and the gap reads as a zero: ' + JSON.stringify(values))
	})

	test('a chart whose every series caches no values is dropped, like one with no series', async () => {
		// Only an empty series *list* was caught. Series that each cache nothing produced an `addChart`
		// of empty value arrays: a blank frame where the source showed no chart.
		const { buf } = await chartDeck(ChartType.bar)
		let removed = 0
		const ir = await irWithChartXml(buf, (xml) =>
			xml.replace(/<c:val>[\s\S]*?<\/c:val>/g, () => {
				removed++
				return ''
			})
		)
		assert(removed > 0, 'the value caches are removed')
		assert(constructs(ir).includes('chart.data'), 'the empty chart is noted; got ' + JSON.stringify(constructs(ir)))
		assertEqual(
			ir.slides[0].calls.filter((call) => call.method === 'addChart').length,
			0,
			'and no chart call is emitted'
		)
	})

	test('scatter series with X values of their own are rebuilt against the first, and say so', async () => {
		// `addChart` plots every series against one X row, and a chart part gives each series a
		// `c:xVal` of its own. When they differ, the second series moves.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addChart(XY, { type: ChartType.scatter, x: 1, y: 1, w: 6, h: 4 })
		})
		// The last X point 1 is the second series' own: its `c:xVal` comes after the first series,
		// and no Y value is 1.
		const ir = await irWithChartXml(buf, (xml) => {
			const find = '<c:pt idx="0"><c:v>1</c:v></c:pt>'
			const at = xml.lastIndexOf(find)
			assert(at > xml.indexOf(find), 'the second series caches its own X row')
			return xml.slice(0, at) + '<c:pt idx="0"><c:v>10</c:v></c:pt>' + xml.slice(at + find.length)
		})
		assert(constructs(ir).includes('chart.xValues'), 'the move is noted; got ' + JSON.stringify(constructs(ir)))
		expect(chartCall(ir).args[0][0]).toEqual({ values: [1, 2, 3] })
	})
})

describe('chart data keeps the shape addChart takes', () => {
	test('a scatter keeps its X values as the leading X row', async () => {
		// The read model had no accessor for `c:xVal` or `c:yVal`, so a scatter's series read back
		// empty and the chart was dropped.
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(XY, { type: ChartType.scatter, x: 1, y: 1, w: 6, h: 4 })
		})
		const ir = readModelToIr(presentation)
		expect(chartCall(ir).args[0]).toEqual([
			{ values: [1, 2, 3] },
			{ name: 'Y1', values: [4, 5, 6] },
			{ name: 'Y2', values: [7, 8, 9] },
		])
		assert(!constructs(ir).includes('chart.data'), 'and nothing is dropped; got ' + JSON.stringify(constructs(ir)))
		assert(
			!constructs(ir).includes('chart.xLabels'),
			'and numbers are not labels; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a PowerPoint scatter against text X labels is rebuilt at the positions it is plotted at, and says so', async () => {
		// PowerPoint caches the X column as strings and draws the points at X = 1..4, which is also
		// what an X row of 1..4 draws. Read as X values, every label was a blank that became a 0,
		// under a note about gaps in a line chart.
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'scatter-text-x-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ name: 'Y', values: [2.5, 4, 3.5, 6] },
		])
		const seen = constructs(ir)
		assert(seen.includes('chart.xLabels'), 'the lost labels are noted; got ' + JSON.stringify(seen))
		assert(!seen.includes('chart.blanks'), 'and no point is called a blank; got ' + JSON.stringify(seen))
	})

	test('a PowerPoint bubble keeps its sizes', async () => {
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'bubble-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ values: [2, 4, 3, 6], sizes: [5, 10, 7, 12] },
		])
	})

	test('PowerPoint multi-level categories become labels leaf first, the outer level filled in with blanks', async () => {
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'multilevel-bar-chart').args[0][0].labels).toEqual([
			['Q1', 'Q2', 'Q1', 'Q2'],
			['North', '', 'South', ''],
		])
	})
})

/** The options argument of slide 1's chart call. */
const chartOptions = (ir) => ir.slides[0].calls.find((call) => call.method === 'addChart').args[1]

/** Rewrite every `<qname val="from"/>` in the chart part to `to`, asserting at least one was there. */
function flipFlag(qname, from, to) {
	return (xml) => {
		let flipped = 0
		const out = xml.replaceAll(`<${qname} val="${from}"/>`, () => {
			flipped++
			return `<${qname} val="${to}"/>`
		})
		assert(flipped > 0, `the chart part carries <${qname} val="${from}"/>`)
		return out
	}
}

describe('chart options are spelled the way ChartOpts spells them', () => {
	// `showCatName` and `showLegendKey` were emitted from the read model's flags, and neither is a
	// `ChartOpts` key: the writer ignored both, and nothing noted it.
	test("a pie's category-name labels are `showLabel`, read from the series where PowerPoint keeps them", async () => {
		// PowerPoint writes a pie's label flags into `c:ser/c:dLbls` and an all-off block for the
		// group. Read from the group block, the fixture's labelled pie was rebuilt with no labels.
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		const options = chartNamed(ir, 'pie-chart').args[1]
		assertEqual(options.showLabel, true, 'the slice names are asked for the way the writer reads them')
		assert(!('showCatName' in options), 'and not under a key the writer ignores')
		assertEqual(options.showPercent, true, 'the percentages are kept')
		assertEqual(options.showValue, false, 'and the values stay off')
		// No `c:dLblPos` is best fit, and the writer spells an unset pie position `ctr`.
		assertEqual(options.dataLabelPosition, 'bestFit', 'the position is written out')
	})

	test('label flags with no ChartOpts spelling are dropped and noted', async () => {
		const { buf } = await chartDeck(ChartType.bar, { showValue: true })
		const ir = await irWithChartXml(buf, (xml) =>
			flipFlag('c:showCatName', '0', '1')(flipFlag('c:showLegendKey', '0', '1')(xml))
		)
		const options = chartOptions(ir)
		for (const key of ['showLegendKey', 'showCatName', 'showLabel']) {
			assert(!(key in options), `a bar chart emits no ${key}; got ${JSON.stringify(options)}`)
		}
		assertEqual(
			constructs(ir).filter((construct) => construct === 'chart.labels').length,
			2,
			'the legend key and the category name are each noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a 3-D line chart is rebuilt flat, and says so', async () => {
		const { buf } = await chartDeck(ChartType.line)
		const ir = await irWithChartXml(buf, (xml) => xml.replaceAll('c:lineChart>', 'c:line3DChart>'))
		assertEqual(chartOptions(ir).type, 'line', 'it is rebuilt as a line chart')
		assert(
			constructs(ir).includes('chart.type3D'),
			'and the lost depth is noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a 2-D surface stays flat', async () => {
		// The write side's `surface` draws the 3-D one unless `surface3D` is false.
		const flat = await chartDeck(ChartType.surface, { surface3D: false })
		assertEqual(chartOptions(readModelToIr(flat.presentation)).surface3D, false, 'a contour says so')
		const tilted = await chartDeck(ChartType.surface)
		assert(!('surface3D' in chartOptions(readModelToIr(tilted.presentation))), 'and a 3-D surface keeps the default')
	})
})
