import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries } from '../../helpers.js'
import { chartExPath } from './chart-parts.js'

// Region map (`regionMap`, a filled/geographic map) is a chartEx (cx:) layout. Confirmed against a
// map PowerPoint authored (`Shapes.AddChart2(-1, 140, …)`) and read back as ChartType 140, the
// pieces SPECIFIC to a region map are:
//   - the numeric dimension is tagged `type="colorVal"` (the value drives each region's fill color),
//     NOT the flat layouts' `type="val"` nor the hierarchical `type="size"`;
//   - the series carries `layoutId="regionMap"` and a <cx:layoutPr> holding a single <cx:geography>
//     hint whose cultureLanguage / cultureRegion / attribution attributes are all schema-REQUIRED;
//   - it is axis-free (geography, not a Cartesian scale) — no <cx:axis> at all;
//   - the DATA is a plain flat labeled category chart (region names in col A + values in col B), the
//     same <cx:strDim>/<cx:numDim> shape as a bar chart — only the colorVal tag + geography differ.
// PowerPoint itself also nests a <cx:geoCache> (a base64 Bing-geometry blob produced by an online
// lookup); that blob cannot be reproduced offline, so it is intentionally OMITTED — PowerPoint
// re-resolves the geography from the region names on open (verified over COM: the deck opens and
// reads back as a regionMap either way). Unlike pareto/histogram, a region map has NO
// schema-vs-PowerPoint divergence — it validates cleanly. The map needs the LATER chartex feature
// namespace (cx4 = …/2016/5/10/chartex) on the slide's <mc:Choice Requires>, unlike the cx1/cx2 the
// earlier layouts use. The shared chartEx wiring (content type, MS rel type, style/colors sidecars,
// mc:Fallback) is asserted in full by chart-waterfall-xml; here we pin the regionMap deltas.

const DATA = [
	{
		name: 'Sales',
		labels: ['United States', 'Canada', 'Mexico', 'Brazil'],
		values: [100, 60, 40, 55],
	},
]

async function buildMap(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, { type: 'regionMap', x: 1, y: 1, w: 8, h: 4.5, ...extra })
	})
}

describe('region map (chartEx) chart', () => {
	test('emits a chartEx part with the mandatory sidecars, no classic chart part', async () => {
		const { zip } = await buildMap()
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

	test('slide Choice requires the LATER chartex feature namespace cx4', async () => {
		const { zip } = await buildMap()
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).toContain(
			'<mc:Choice xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" Requires="cx4">'
		)
		expect(slide).toContain('<mc:Fallback>')
	})

	test('numeric dimension is tagged colorVal (value drives region fill), over a flat labeled layout', async () => {
		const { zip } = await buildMap()
		const xml = await readEntry(zip, chartExPath(zip))
		// Region names in col A, values in col B — the flat labeled shape.
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$5</cx:f>')
		expect(xml).toContain('<cx:pt idx="0">United States</cx:pt>')
		// The map-specific tag: colorVal, NOT val or size.
		expect(xml).toContain('<cx:numDim type="colorVal"><cx:f>Sheet1!$B$2:$B$5</cx:f>')
		expect(xml).toContain('<cx:pt idx="0">100</cx:pt>')
		expect(xml).not.toContain('type="val"')
		expect(xml).not.toContain('type="size"')
		// Series name is the value column's header — column B.
		expect(xml).toContain('<cx:f>Sheet1!$B$1</cx:f>')
	})

	test('single regionMap series with a default geography hint and no axes', async () => {
		const { zip } = await buildMap()
		const xml = await readEntry(zip, chartExPath(zip))
		expect((xml.match(/<cx:series /g) || []).length).toBe(1)
		expect(xml).toContain('layoutId="regionMap"')
		// All three geography attributes are schema-required; defaults are en-US / US.
		expect(xml).toContain(
			'<cx:layoutPr><cx:geography cultureLanguage="en-US" cultureRegion="US" attribution="Powered by Bing"/></cx:layoutPr>'
		)
		// The un-reproducible Bing geometry cache is intentionally omitted.
		expect(xml).not.toContain('<cx:geoCache')
		// A region map plots on geography — no Cartesian axis.
		expect(xml).not.toContain('<cx:axis')
	})

	test('geography opt overrides the culture language and region', async () => {
		const { zip } = await buildMap({ geography: { cultureLanguage: 'fr-FR', cultureRegion: 'FR' } })
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain(
			'<cx:layoutPr><cx:geography cultureLanguage="fr-FR" cultureRegion="FR" attribution="Powered by Bing"/></cx:layoutPr>'
		)
	})
})
