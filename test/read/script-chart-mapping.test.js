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

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { ChartType } from '../../dist/node.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]

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
})
