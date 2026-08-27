// The bound on a chart point cache — `src/read/oxml/point-cache.ts`, shared by the classic
// `c:` reader and the 2016 `cx:` one.
//
// Reading a deck is the one part of this library whose input comes from somewhere other than
// the calling program, and a point cache carries two numbers a deck can set freely. Both were
// obeyed: `c:ptCount/@val` sized a `new Array(n).fill(null)` directly, and a single high `@idx`
// widened it past whatever the count said. `@ptCount` is `xsd:unsignedInt`, so `4294967295` is
// schema-valid, and an allocation that size is not a slow path — V8 answers `FATAL ERROR:
// invalid table size` and the host process dies with no exception to catch. Reproduced end to
// end before the fix: `Presentation.load()` returned normally and the first `chart.series`
// access killed the run, which is why every case here asserts a value rather than merely that
// nothing threw. A crash would have taken the whole vitest worker, not failed one case.
//
// Each case authors a real chart with the write API and edits ONE attribute value in the bytes
// it produced, so everything around the hostile number is genuine writer output rather than
// hand-typed XML. Both chart families get a case, because the two decodes are one function now
// and a regression in either would otherwise only show up in whichever family a test picked.
//
// The reads happen INSIDE the diagnostic capture on purpose: the read model is lazy, so a cache
// is not decoded — and cannot warn — until something asks it for its points.

import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { ChartType } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { describe, test } from 'vitest'
import { authorRead } from './authored.js'
import { assert, assertEqual, captureDiagnostics } from '../helpers.js'

/** Author a chart, rewrite its chart part with `edit`, load the result back and hand it to `read`. */
async function authorEditRead(chartType, edit, read) {
	const { buf } = await authorRead((pres) => {
		pres.addSlide().addChart([{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [10, 20, 30, 40] }], {
			type: chartType,
			x: 1,
			y: 1,
			w: 6,
			h: 4,
		})
	})
	const zip = await JSZip.loadAsync(buf)
	const partName = Object.keys(zip.files).find((n) => /^ppt\/charts\/chart(Ex)?\d+\.xml$/.test(n))
	assert(partName, `expected a chart part; got ${Object.keys(zip.files).filter((n) => n.includes('charts'))}`)
	zip.file(partName, edit(await zip.files[partName].async('string')))
	return read(await Presentation.load(await zip.generateAsync({ type: 'uint8array' })))
}

/** The first classic chart on any slide. */
function firstChartOf(presentation) {
	for (const slide of presentation.slides) for (const shape of slide.shapes) if (shape.chart) return shape.chart
	return null
}

/** The first chartEx chart on any slide. */
function firstChartExOf(presentation) {
	for (const slide of presentation.slides) for (const shape of slide.shapes) if (shape.chartEx) return shape.chartEx
	return null
}

/** Replace the last occurrence of `find` in `xml` with `replacement`. */
function replaceLast(xml, find, replacement) {
	const at = xml.lastIndexOf(find)
	assert(at !== -1, `expected the authored chart part to contain ${find}`)
	return xml.slice(0, at) + replacement + xml.slice(at + find.length)
}

describe('Chart point caches are sized by the points that are there', () => {
	test('a declared c:ptCount of 4294967295 reads the four real points and warns', async () => {
		const { result, codes } = await captureDiagnostics(() =>
			authorEditRead(
				ChartType.bar,
				(xml) => xml.replace(/<c:ptCount val="4"\/>/g, '<c:ptCount val="4294967295"/>'),
				(presentation) => {
					const chart = firstChartOf(presentation)
					assert(chart, 'the edited deck still reads as a chart')
					return { values: chart.series[0].values.length, categories: chart.categories.length }
				}
			)
		)
		assertEqual(result.values, 4, 'the four points that are actually in the cache')
		assertEqual(result.categories, 4, 'and the four categories alongside them')
		assert(codes.includes('chart/point-count-mismatch'), `expected the mismatch warning; got ${codes}`)
	})

	test('a c:pt indexed past the worksheet bound is dropped rather than sizing the array', async () => {
		const { result, codes } = await captureDiagnostics(() =>
			authorEditRead(
				ChartType.bar,
				(xml) => replaceLast(xml, '<c:pt idx="3">', '<c:pt idx="900000000">'),
				(presentation) => firstChartOf(presentation).series[0].values.length
			)
		)
		// A worksheet has 1,048,576 rows, so a point at index 900,000,000 cannot describe data
		// any producer charted. It is dropped, and the array ends at the last real point.
		assertEqual(result, 3, 'the points below the bound')
		assert(codes.includes('chart/point-index-out-of-range'), `expected the out-of-range warning; got ${codes}`)
	})

	test('chartEx caches take the same bound from the same helper', async () => {
		const { result, codes } = await captureDiagnostics(() =>
			authorEditRead(
				ChartType.waterfall,
				(xml) => xml.replace(/ptCount="4"/g, 'ptCount="4294967295"'),
				(presentation) => {
					const chartEx = firstChartExOf(presentation)
					assert(chartEx, 'the edited deck still reads as a chartEx chart')
					return chartEx.categories.length
				}
			)
		)
		assertEqual(result, 4, 'the four points that are actually in the level')
		assert(codes.includes('chart/point-count-mismatch'), `expected the mismatch warning; got ${codes}`)
	})

	test('a cache whose count agrees with its points is read silently', async () => {
		const { result, codes } = await captureDiagnostics(() =>
			authorEditRead(
				ChartType.bar,
				(xml) => xml,
				(presentation) => firstChartOf(presentation).series[0].values.length
			)
		)
		assertEqual(result, 4, 'the untouched cache reads its four points')
		assertEqual(
			codes.filter((c) => c.startsWith('chart/point-')).length,
			0,
			`an ordinary chart warns about nothing; got ${codes}`
		)
	})
})

describe('The bound holds through the public read entry point', () => {
	test('a hostile ptCount on a PowerPoint-authored deck survives load + series access', async () => {
		const zip = await JSZip.loadAsync(await readFile('test/read/fixtures/bar-chart-data-labels.pptx'))
		const part = 'ppt/charts/chart1.xml'
		const xml = await zip.files[part].async('string')
		zip.file(part, xml.replace(/<c:ptCount val="\d+"\/>/g, '<c:ptCount val="4294967295"/>'))
		const { result } = await captureDiagnostics(async () => {
			const presentation = await Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
			const chart = firstChartOf(presentation)
			assert(chart, 'the deck still reads as a chart')
			return chart.series[0].values.length
		})
		assert(result > 0 && result < 1000, `the series is sized by its real points, not the claim; got ${result}`)
	})
})
