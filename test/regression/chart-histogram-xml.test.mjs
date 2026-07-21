import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries, assert } from '../helpers.js'

// Histogram is the category-less chartEx (cx:) layout. Unlike every other chart, it is fed RAW
// OBSERVATIONS (a single `values` series, no `labels`) and PowerPoint bins them itself. Confirmed
// against a chart PowerPoint authored (`Shapes.AddChart2(-1, 118, …)`) and read back as ChartType
// 118, the pieces SPECIFIC to histogram are:
//   - the chart data has NO <cx:strDim> at all — just one <cx:numDim type="val"> whose values sit
//     in column A (there are no leading label columns), so the series-name <cx:tx> is $A$1;
//   - the series layoutId is "clusteredColumn" (histogram is a binned clustered column) with a
//     <cx:layoutPr><cx:binning> that carries the interval-closed side;
//   - it keeps a category (id 0, gapWidth 0) + value (id 1) axis pair.
// The shared chartEx wiring (content type, MS rel type, style/colors sidecars, mc:Fallback) is
// asserted in full by chart-waterfall-xml; here we pin the category-less deltas.

const OBS = [
	{
		name: 'Test Scores',
		values: [55, 62, 68, 71, 72, 74, 75, 77, 78, 78, 80, 81, 82, 83, 85, 86, 88, 90, 92, 95],
	},
]

function chartExPath(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chartEx\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartExN.xml entry')
	return path
}

async function buildHistogram(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(OBS, { type: 'histogram', x: 1, y: 1, w: 8, h: 4.5, ...extra })
	})
}

describe('histogram (category-less chartEx) chart', () => {
	test('emits a chartEx part with the mandatory sidecars, no classic chart part', async () => {
		const { zip } = await buildHistogram()
		expect(chartExPath(zip)).toBe('ppt/charts/chartEx1.xml')
		const contentTypes = await readEntry(zip, '[Content_Types].xml')
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>'
		)
		const entries = listEntries(zip)
		expect(entries).toContain('ppt/charts/style1.xml')
		expect(entries).toContain('ppt/charts/colors1.xml')
		expect(entries.some((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))).toBe(false)
	})

	test('slide Choice requires the base feature namespace cx1', async () => {
		const { zip } = await buildHistogram()
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).toContain(
			'<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
		)
		expect(slide).toContain('<mc:Fallback>')
	})

	test('chart data has no category dimension — one value dim of raw observations in column A', async () => {
		const { zip } = await buildHistogram()
		const xml = await readEntry(zip, chartExPath(zip))
		// No category/label dimension whatsoever.
		expect(xml).not.toContain('cx:strDim')
		// A single numeric dimension in column A (no leading label columns), 20 observations.
		expect(xml).toContain('<cx:numDim type="val"><cx:f>Sheet1!$A$2:$A$21</cx:f>')
		expect(xml).toContain('<cx:lvl ptCount="20" formatCode="General">')
		expect(xml).toContain('<cx:pt idx="0">55</cx:pt>')
		// Series name is the value column's header — column A, not B.
		expect(xml).toContain('<cx:f>Sheet1!$A$1</cx:f>')
	})

	test('series is a clusteredColumn with binning; the layout keeps a cat + val axis', async () => {
		const { zip } = await buildHistogram()
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('layoutId="clusteredColumn"')
		// Auto binning defaults to right-closed intervals, matching PowerPoint.
		expect(xml).toContain('<cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr>')
		expect(xml).toContain('<cx:axis id="0"><cx:catScaling gapWidth="0"/><cx:tickLabels/></cx:axis>')
		expect(xml).toContain('<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>')
	})

	test('binning.intervalClosed flips the interval side', async () => {
		const { zip } = await buildHistogram({ binning: { intervalClosed: 'l' } })
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('<cx:binning intervalClosed="l"/>')
	})
})
