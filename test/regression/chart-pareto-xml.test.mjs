import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries, assert } from '../helpers.js'

// Pareto is the first MULTI-SERIES chartEx (cx:) layout. Confirmed against a chart PowerPoint
// authored (`Shapes.AddChart2(-1, 122, …)`) and read back as ChartType 122, the pieces SPECIFIC to
// pareto are:
//   - the plotAreaRegion holds TWO <cx:series>. Series 0 is a `clusteredColumn` (the bars) bound to
//     value axis 1, with a <cx:layoutPr><cx:aggregation/> that tells PowerPoint to sum the values
//     per category, sort the bars descending, and derive the cumulative line. Series 1 is the
//     `paretoLine` itself: `ownerIdx="0"` (it derives its data from series 0, so it has NO <cx:tx>
//     and NO <cx:dataId>) bound to axis 2;
//   - THREE axes: category (id 0, gapWidth 0), primary value (id 1, gridlines), and a SECONDARY
//     value axis (id 2) scaled 0..1 with percentage units — the scale the cumulative line plots on;
//   - the DATA is a plain flat category chart (labels in col A + values in col B), so its <cx:strDim>
//     / <cx:numDim> shape is the same as a bar chart's — aggregation, not a special data layout, is
//     what makes it a pareto.
// The `<cx:axisId val="N"/>` attribute form is a deliberate schema-vs-PowerPoint divergence: the
// OpenXML SDK models cx:axisId as leaf TEXT, but PowerPoint refuses to open the text form (see the
// schema-case comment + COM smoke). The shared chartEx wiring (content type, MS rel type,
// style/colors sidecars, mc:Fallback) is asserted in full by chart-waterfall-xml; here we pin the
// multi-series + secondary-axis deltas.

const DATA = [
	{
		name: 'Defects',
		labels: ['Scratch', 'Dent', 'Crack', 'Smudge', 'Chip'],
		values: [45, 30, 15, 7, 3],
	},
]

function chartExPath(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chartEx\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartExN.xml entry')
	return path
}

async function buildPareto(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, { type: 'pareto', x: 1, y: 1, w: 8, h: 4.5, ...extra })
	})
}

describe('pareto (multi-series chartEx) chart', () => {
	test('emits a chartEx part with the mandatory sidecars, no classic chart part', async () => {
		const { zip } = await buildPareto()
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
		const { zip } = await buildPareto()
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).toContain(
			'<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
		)
		expect(slide).toContain('<mc:Fallback>')
	})

	test('chart data is a flat labeled category chart — cat in col A, val in col B', async () => {
		const { zip } = await buildPareto()
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$6</cx:f>')
		expect(xml).toContain('<cx:pt idx="0">Scratch</cx:pt>')
		expect(xml).toContain('<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$6</cx:f>')
		expect(xml).toContain('<cx:pt idx="0">45</cx:pt>')
		// Series name is the value column's header — column B.
		expect(xml).toContain('<cx:f>Sheet1!$B$1</cx:f>')
	})

	test('emits two series: a clusteredColumn (aggregation, axis 1) and a paretoLine (ownerIdx 0, axis 2)', async () => {
		const { zip } = await buildPareto()
		const xml = await readEntry(zip, chartExPath(zip))
		// Exactly two series.
		expect((xml.match(/<cx:series /g) || []).length).toBe(2)
		// Series 0: the bars. clusteredColumn + <cx:aggregation/> + binds axis 1.
		expect(xml).toContain('layoutId="clusteredColumn"')
		expect(xml).toContain('<cx:layoutPr><cx:aggregation/></cx:layoutPr>')
		// Series 1: the cumulative line. It derives from series 0 (ownerIdx 0), so it has no tx/dataId.
		expect(xml).toContain('<cx:series layoutId="paretoLine" ownerIdx="0"')
		const lineSeries = xml.slice(xml.indexOf('<cx:series layoutId="paretoLine"'))
		expect(lineSeries).not.toContain('<cx:tx>')
		expect(lineSeries).not.toContain('<cx:dataId')
		// Axis bindings: attribute form (PowerPoint-required; the SDK text form corrupts the deck).
		expect(xml).toContain('<cx:axisId val="1"/>')
		expect(xml).toContain('<cx:axisId val="2"/>')
	})

	test('emits three axes including the secondary 0..1 percentage axis', async () => {
		const { zip } = await buildPareto()
		const xml = await readEntry(zip, chartExPath(zip))
		expect((xml.match(/<cx:axis /g) || []).length).toBe(3)
		// Category axis: abutting bars (gapWidth 0), like a histogram.
		expect(xml).toContain('<cx:axis id="0"><cx:catScaling gapWidth="0"/><cx:tickLabels/></cx:axis>')
		// Primary value axis: the bar scale, with gridlines.
		expect(xml).toContain('<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>')
		// Secondary value axis: the cumulative-line scale, 0..1 as a percentage, no gridlines.
		expect(xml).toContain(
			'<cx:axis id="2"><cx:valScaling max="1" min="0"/><cx:units unit="percentage"/><cx:tickLabels/></cx:axis>'
		)
	})
})
