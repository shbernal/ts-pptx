import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, listEntries, assert, assertEqual } from '../../helpers.js'

// Chart part filenames must be a pure function of deck structure, not of process
// history. They used to be drawn from a never-reset module global (`_chartCounter`),
// so two identical, independent presentations built in one process emitted different
// chart part filenames (`chart1.xml`/`chart2.xml` vs `chart3.xml`/`chart4.xml`) —
// same input, different bytes. `exportPresentation` now assigns them from a
// per-presentation counter at write time. See backlog fork-chart-counter-nondeterminism.

const chartParts = (zip) =>
	listEntries(zip)
		.filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
		.sort()
const embeddingParts = (zip) =>
	listEntries(zip)
		.filter((f) => /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(f))
		.sort()

const buildTwoChartDeck = (p) => {
	const data1 = [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }]
	const data2 = [{ name: 'S2', labels: ['A', 'B'], values: [3, 4] }]
	p.addSlide().addChart(data1, { type: ChartType.bar, x: 1, y: 1, w: 4, h: 3 })
	p.addSlide().addChart(data2, { type: ChartType.line, x: 1, y: 1, w: 4, h: 3 })
}

defineRegressionSuite('Chart part-name determinism', 'backlog fork-chart-counter-nondeterminism', [
	{
		name: 'two identical decks built in one process emit identical chart part filenames',
		fn: async () => {
			// Built sequentially in the same process — the previous module-global counter
			// carried its value from the first build into the second.
			const { zip: zipA } = await build(buildTwoChartDeck)
			const { zip: zipB } = await build(buildTwoChartDeck)

			const partsA = chartParts(zipA)
			const partsB = chartParts(zipB)

			assertEqual(partsA.length, 2, 'first deck has two chart parts')
			// Numbering is per-presentation from 1, so both decks agree — and both start at chart1.
			assert(
				JSON.stringify(partsA) === JSON.stringify(['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml']),
				'chart parts numbered per-presentation from 1; got ' + JSON.stringify(partsA)
			)
			assert(
				JSON.stringify(partsA) === JSON.stringify(partsB),
				'identical decks must emit identical chart parts; got ' +
					JSON.stringify(partsA) +
					' vs ' +
					JSON.stringify(partsB)
			)

			// Embedded workbooks share the same globalId, so they must line up too.
			const embA = embeddingParts(zipA)
			assert(
				JSON.stringify(embA) ===
					JSON.stringify([
						'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',
						'ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx',
					]),
				'embedded workbooks numbered per-presentation from 1; got ' + JSON.stringify(embA)
			)
			assert(
				JSON.stringify(embA) === JSON.stringify(embeddingParts(zipB)),
				'identical decks must emit identical embedded workbook parts'
			)
		},
	},
	{
		name: 'writing the same presentation twice is byte-identical across chart parts',
		fn: async () => {
			// A single presentation exported twice must renumber to the same names each time
			// (the write-time pass must not accumulate state across exports).
			const pres = (await build(buildTwoChartDeck)).pres
			const { zip: first } = { zip: await reExport(pres) }
			const { zip: second } = { zip: await reExport(pres) }
			assert(
				JSON.stringify(chartParts(first)) === JSON.stringify(chartParts(second)),
				'same presentation exported twice must emit identical chart parts'
			)
			assertEqual(chartParts(first).length, 2, 'two chart parts on re-export')
		},
	},
])

async function reExport(pres) {
	const JSZip = (await import('jszip')).default
	const buf = /** @type {Uint8Array} */ (await pres.stream())
	return JSZip.loadAsync(buf)
}
