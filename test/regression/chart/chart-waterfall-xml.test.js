import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries } from '../../helpers.js'
import { chartExPath } from './chart-parts.js'

// Waterfall is the first chartEx (cx:) chart type. Unlike the classic 2007 charts it emits a
// SEPARATE part (`ppt/charts/chartExN.xml`) in the Office-2016 chart-extension namespace, gets the
// `chartex+xml` content type + the MS chartEx rel type, and is referenced from the slide through
// `<mc:AlternateContent>` (Choice + Fallback). This pins that whole wiring plus the schema-shaped
// details the OOXML validator is strict about (externalData is a leaf placed BEFORE cx:data, with
// only an r:id — no `autoUpdate`).

const DATA = [{ name: 'Cash Flow', labels: ['Start', 'Q1', 'Q2', 'Q3', 'End'], values: [100, 40, -30, 20, 130] }]

async function buildWaterfall(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, {
			type: 'waterfall',
			x: 1,
			y: 1,
			w: 8,
			h: 4,
			showValue: true,
			showLegend: true,
			legendPos: 't',
			subtotals: [0, 4],
			...extra,
		})
	})
}

describe('waterfall (chartEx) chart', () => {
	test('emits a chartEx part with the chartex content type and MS rel type', async () => {
		const { zip } = await buildWaterfall()
		const cxPath = chartExPath(zip)
		expect(cxPath).toBe('ppt/charts/chartEx1.xml')

		const contentTypes = await readEntry(zip, '[Content_Types].xml')
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>'
		)

		const slideRels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
		expect(slideRels).toContain(
			'Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="/ppt/charts/chartEx1.xml"'
		)
		// classic chart parts / rel type must NOT appear
		expect(listEntries(zip).some((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))).toBe(false)
	})

	test('emits the mandatory chart-style + color-style sidecar parts', async () => {
		// A chartEx part without these renders as corrupt in PowerPoint (0x80070570), even though it
		// passes schema validation — so they are required, not optional.
		const { zip } = await buildWaterfall()
		const entries = listEntries(zip)
		expect(entries).toContain('ppt/charts/style1.xml')
		expect(entries).toContain('ppt/charts/colors1.xml')

		const contentTypes = await readEntry(zip, '[Content_Types].xml')
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>'
		)
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>'
		)

		const cxRels = await readEntry(zip, 'ppt/charts/_rels/chartEx1.xml.rels')
		expect(cxRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"')
		expect(cxRels).toContain(
			'Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"'
		)
		expect(cxRels).toContain(
			'Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"'
		)
	})

	test('slide references the chart via mc:AlternateContent (Choice + Fallback)', async () => {
		const { zip } = await buildWaterfall()
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).toContain(
			'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
		)
		expect(slide).toContain(
			'<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
		)
		expect(slide).toContain('<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">')
		expect(slide).toContain(
			'<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="rId1"/>'
		)
		// Fallback shows a placeholder shape for non-2016 consumers
		expect(slide).toContain('<mc:Fallback>')
		expect(slide).toContain('This chart requires PowerPoint 2016 or newer to display.')
	})

	test('chartEx XML has the waterfall layout, subtotals, and schema-shaped chartData', async () => {
		const { zip } = await buildWaterfall()
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"')
		expect(xml).toContain('layoutId="waterfall"')
		expect(xml).toContain('<cx:subtotals><cx:idx val="0"/><cx:idx val="4"/></cx:subtotals>')
		// externalData is a leaf with only r:id, and precedes <cx:data>
		expect(xml).toContain('<cx:chartData><cx:externalData r:id="rId1"/><cx:data id="0">')
		expect(xml).not.toContain('autoUpdate')
		// category + value caches
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$6</cx:f>')
		expect(xml).toContain('<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$6</cx:f>')
		expect(xml).toContain('<cx:pt idx="2">-30</cx:pt>')
		// legend + data labels honored
		expect(xml).toContain('<cx:legend pos="t" align="ctr" overlay="0"/>')
		expect(xml).toContain('<cx:visibility seriesName="0" categoryName="0" value="1"/>')
	})

	test('omits subtotals/legend/dataLabels when not requested', async () => {
		const { zip } = await buildWaterfall({ showValue: false, showLegend: false, subtotals: undefined })
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).not.toContain('cx:subtotals')
		expect(xml).not.toContain('cx:legend')
		expect(xml).not.toContain('cx:dataLabels')
	})

	test('invalid subtotal indices are dropped', async () => {
		const { zip } = await buildWaterfall({ subtotals: [0, -1, 2.5, 3] })
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('<cx:subtotals><cx:idx val="0"/><cx:idx val="3"/></cx:subtotals>')
	})
})
