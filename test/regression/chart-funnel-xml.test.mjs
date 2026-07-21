import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries, assert } from '../helpers.js'

// Funnel is the second chartEx (cx:) chart type, landing on the subsystem waterfall introduced.
// This pins the parts that are SPECIFIC to funnel (and were confirmed against a chart PowerPoint
// itself authored, `Shapes.AddChart2(-1, 123, …)`):
//   - it declares a DIFFERENT feature namespace on `<mc:Choice>` than waterfall — `cx2`
//     (`…/2015/10/21/chartex`), not waterfall's `cx1` (`…/2015/9/8/chartex`);
//   - its plot area carries a SINGLE category axis (PowerPoint numbers it id="1") with no value
//     axis and no gridlines — the funnel bars run off one category scale.
// The shared chartEx wiring (chartex+xml content type, MS rel type, style/colors sidecars) is
// asserted in full by chart-waterfall-xml; here we re-check the load-bearing pieces plus the deltas.

const DATA = [
	{
		name: 'Sales Funnel',
		labels: ['Leads', 'Qualified', 'Proposals', 'Negotiation', 'Won'],
		values: [5000, 4000, 3000, 1000, 250],
	},
]

function chartExPath(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chartEx\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartExN.xml entry')
	return path
}

async function buildFunnel(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, {
			type: 'funnel',
			x: 1,
			y: 1,
			w: 8,
			h: 4.5,
			showValue: true,
			...extra,
		})
	})
}

describe('funnel (chartEx) chart', () => {
	test('emits a chartEx part with the mandatory style + color-style sidecars', async () => {
		const { zip } = await buildFunnel()
		const cxPath = chartExPath(zip)
		expect(cxPath).toBe('ppt/charts/chartEx1.xml')

		const contentTypes = await readEntry(zip, '[Content_Types].xml')
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>'
		)
		const entries = listEntries(zip)
		expect(entries).toContain('ppt/charts/style1.xml')
		expect(entries).toContain('ppt/charts/colors1.xml')

		const slideRels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
		expect(slideRels).toContain(
			'Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="/ppt/charts/chartEx1.xml"'
		)
		// classic chart parts must NOT appear
		expect(listEntries(zip).some((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))).toBe(false)
	})

	test('slide Choice requires the funnel feature namespace cx2 (not waterfall cx1)', async () => {
		const { zip } = await buildFunnel()
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).toContain(
			'<mc:Choice xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" Requires="cx2">'
		)
		expect(slide).not.toContain('2015/9/8/chartex')
		// still the AlternateContent + Fallback wrapper
		expect(slide).toContain('<mc:Fallback>')
		expect(slide).toContain('This chart requires PowerPoint 2016 or newer to display.')
	})

	test('chartEx XML has the funnel layout and a single category axis (no value axis)', async () => {
		const { zip } = await buildFunnel()
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('layoutId="funnel"')
		// one category axis, PowerPoint numbers it id="1"; no value axis, no gridlines
		expect(xml).toContain('<cx:axis id="1"><cx:catScaling gapWidth="2.19"/><cx:tickLabels/></cx:axis>')
		expect(xml).not.toContain('cx:valScaling')
		expect(xml).not.toContain('cx:majorGridlines')
		expect(xml).not.toContain('<cx:axis id="0"') // the funnel axis is id 1, not waterfall's id-0 cat axis
		// funnel has no subtotals (a waterfall-only layoutPr)
		expect(xml).not.toContain('cx:subtotals')
		// shared cache shape is intact
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$6</cx:f>')
		expect(xml).toContain('<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$6</cx:f>')
	})

	test('omits data labels when values are not requested', async () => {
		const { zip } = await buildFunnel({ showValue: false })
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).not.toContain('cx:dataLabels')
	})
})
