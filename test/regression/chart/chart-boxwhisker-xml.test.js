import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries } from '../../helpers.js'
import { chartExPath } from './chart-parts.js'

// Box-and-whisker (`boxWhisker`) is a chartEx (cx:) layout. Confirmed against a chart PowerPoint
// authored (`Shapes.AddChart2(-1, 121, …)`) and read back as ChartType 121, the pieces SPECIFIC to
// box-and-whisker are:
//   - the series carries `layoutId="boxWhisker"` and a <cx:layoutPr> holding a <cx:visibility> toggle
//     set (meanLine / meanMarker / nonoutliers / outliers, each a 0|1) followed by a <cx:statistics>
//     with the quartileMethod choice — in that document order (what PowerPoint itself emits);
//   - it binds NO <cx:axisId> (a single default value axis is implicit, unlike pareto's two);
//   - two axes: category (id 0, gapWidth 1 — the wider default gap PowerPoint uses to space the
//     boxes) + value (id 1, gridlines);
//   - the DATA is a plain flat labeled category chart (labels in col A + values in col B): each value
//     series is summarized into a box, so the <cx:strDim>/<cx:numDim> shape is the same as a bar
//     chart's — the statistics config, not a special data layout, is what makes it a box plot.
// Unlike pareto, box-and-whisker has no schema-vs-PowerPoint divergence — it validates cleanly. The
// shared chartEx wiring (content type, MS rel type, style/colors sidecars, mc:Fallback) is asserted
// in full by chart-waterfall-xml; here we pin the boxWhisker series/axis/statistics deltas.

const DATA = [
	{
		name: 'Measurements',
		labels: ['Line A', 'Line A', 'Line A', 'Line B', 'Line B', 'Line B'],
		values: [12, 15, 9, 22, 18, 25],
	},
]

async function buildBox(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, { type: 'boxWhisker', x: 1, y: 1, w: 8, h: 4.5, ...extra })
	})
}

describe('box-and-whisker (chartEx) chart', () => {
	test('emits a chartEx part with the mandatory sidecars, no classic chart part', async () => {
		const { zip } = await buildBox()
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
		const { zip } = await buildBox()
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).toContain(
			'<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
		)
		expect(slide).toContain('<mc:Fallback>')
	})

	test('chart data is a flat labeled category chart — cat in col A, val in col B', async () => {
		const { zip } = await buildBox()
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$7</cx:f>')
		expect(xml).toContain('<cx:pt idx="0">Line A</cx:pt>')
		expect(xml).toContain('<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$7</cx:f>')
		expect(xml).toContain('<cx:pt idx="0">12</cx:pt>')
		// Series name is the value column's header — column B.
		expect(xml).toContain('<cx:f>Sheet1!$B$1</cx:f>')
	})

	test('single boxWhisker series with default statistics/visibility and no axisId binding', async () => {
		const { zip } = await buildBox()
		const xml = await readEntry(zip, chartExPath(zip))
		// Exactly one series.
		expect((xml.match(/<cx:series /g) || []).length).toBe(1)
		expect(xml).toContain('layoutId="boxWhisker"')
		// PowerPoint defaults: mean marker + outliers shown; mean line + full non-outlier scatter off.
		expect(xml).toContain(
			'<cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr>'
		)
		// A boxWhisker series binds no explicit axis (unlike pareto's dual axisId).
		expect(xml).not.toContain('<cx:axisId')
	})

	test('statistics opt overrides the visibility flags and quartile method', async () => {
		const { zip } = await buildBox({
			statistics: {
				quartileMethod: 'inclusive',
				meanLine: true,
				meanMarker: false,
				outliers: false,
				nonoutliers: true,
			},
		})
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain(
			'<cx:layoutPr><cx:visibility meanLine="1" meanMarker="0" nonoutliers="1" outliers="0"/><cx:statistics quartileMethod="inclusive"/></cx:layoutPr>'
		)
	})

	test('emits two axes: category (gapWidth 1) + value (gridlines)', async () => {
		const { zip } = await buildBox()
		const xml = await readEntry(zip, chartExPath(zip))
		expect((xml.match(/<cx:axis /g) || []).length).toBe(2)
		expect(xml).toContain('<cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>')
		expect(xml).toContain('<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>')
	})
})
