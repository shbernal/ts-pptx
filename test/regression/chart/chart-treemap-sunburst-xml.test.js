import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries } from '../../helpers.js'
import { chartExPath } from './chart-parts.js'

// Treemap + sunburst are the hierarchical chartEx (cx:) layouts. They ride the same subsystem
// waterfall/funnel introduced but exercise the parts SPECIFIC to nested categories (confirmed
// against charts PowerPoint itself authored, `Shapes.AddChart2(-1, 117|120, …)`):
//   - the category dimension carries MULTIPLE <cx:lvl> (one per hierarchy level, LEAF FIRST) under
//     a SINGLE <cx:f> range that spans every label column;
//   - the numeric dimension is tagged type="size" (not the flat layouts' type="val"), and its
//     column — plus the series-name <cx:tx> cell — sit AFTER all the label columns;
//   - both are axis-free; treemap adds <cx:layoutPr><cx:parentLabelLayout>, sunburst has neither;
//   - both declare the base feature namespace cx1 (like waterfall), NOT funnel's cx2.
// The shared chartEx wiring (content type, MS rel type, style/colors sidecars, the mc:Fallback) is
// asserted in full by chart-waterfall-xml / chart-funnel-xml; here we pin the hierarchical deltas.

// 3-level hierarchy: labels[0] is the leaf (most granular), labels[last] is the root. Six leaves.
const DATA = [
	{
		name: 'Population',
		labels: [
			['Seattle', 'Portland', 'SF', 'LA', 'Austin', 'Dallas'], // leaf
			['WA', 'OR', 'CA', 'CA', 'TX', 'TX'], // state
			['West', 'West', 'West', 'West', 'South', 'South'], // region (root)
		],
		values: [750, 650, 880, 3900, 970, 1340],
	},
]

async function buildHier(type, extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, { type, x: 1, y: 1, w: 8, h: 4.5, showValue: true, ...extra })
	})
}

describe('treemap + sunburst (hierarchical chartEx) charts', () => {
	test('emit a chartEx part with the mandatory style + color-style sidecars, no classic chart part', async () => {
		for (const type of ['treemap', 'sunburst']) {
			const { zip } = await buildHier(type)
			const cxPath = chartExPath(zip)
			expect(cxPath).toBe('ppt/charts/chartEx1.xml')

			const contentTypes = await readEntry(zip, '[Content_Types].xml')
			expect(contentTypes).toContain(
				'<Override PartName="/ppt/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>'
			)
			const entries = listEntries(zip)
			expect(entries).toContain('ppt/charts/style1.xml')
			expect(entries).toContain('ppt/charts/colors1.xml')
			expect(entries.some((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))).toBe(false)
		}
	})

	test('slide Choice requires the base feature namespace cx1 (like waterfall, not funnel cx2)', async () => {
		for (const type of ['treemap', 'sunburst']) {
			const { zip } = await buildHier(type)
			const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
			expect(slide).toContain(
				'<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
			)
			expect(slide).not.toContain('2015/10/21/chartex') // funnel's cx2 level
			expect(slide).toContain('<mc:Fallback>')
		}
	})

	test('category dimension has one <cx:lvl> per level (leaf first) under a single spanning <cx:f>', async () => {
		const { zip } = await buildHier('treemap')
		const xml = await readEntry(zip, chartExPath(zip))
		// Single formula spans all three label columns A..C over the six data rows.
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$C$7</cx:f>')
		// Three levels, emitted leaf-first: level 0 is the leaf (city), level 2 is the root (region).
		const lvls = [...xml.matchAll(/<cx:lvl[^>]*>(.*?)<\/cx:lvl>/gs)].map((m) => m[1])
		expect(lvls.length).toBe(4) // 3 category levels + 1 value level
		expect(lvls[0]).toContain('<cx:pt idx="0">Seattle</cx:pt>')
		expect(lvls[2]).toContain('<cx:pt idx="0">West</cx:pt>')
		expect(lvls[2]).toContain('<cx:pt idx="5">South</cx:pt>')
	})

	test('value dimension is type="size" in the column after the labels; series name references that column', async () => {
		const { zip } = await buildHier('treemap')
		const xml = await readEntry(zip, chartExPath(zip))
		// Hierarchical layouts tag the numeric dim "size", not the flat layouts' "val".
		expect(xml).toContain('<cx:numDim type="size"><cx:f>Sheet1!$D$2:$D$7</cx:f>')
		expect(xml).not.toContain('type="val"')
		// Series name (<cx:tx>) points at the value column's header cell (D1 for 3 label columns).
		expect(xml).toContain('<cx:f>Sheet1!$D$1</cx:f>')
	})

	test('treemap has parentLabelLayout and no axes', async () => {
		const { zip } = await buildHier('treemap')
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('layoutId="treemap"')
		expect(xml).toContain('<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>')
		expect(xml).not.toContain('<cx:axis')
	})

	test('sunburst has neither parentLabelLayout nor axes', async () => {
		const { zip } = await buildHier('sunburst')
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('layoutId="sunburst"')
		expect(xml).not.toContain('cx:parentLabelLayout')
		expect(xml).not.toContain('<cx:axis')
	})

	test('a flat one-level category set degrades to the column-A/B cache', async () => {
		// The multi-level generalization must still produce the flat shape for a single label level.
		const { zip } = await build((p) => {
			p.addSlide().addChart([{ name: 'N', labels: ['a', 'b', 'c'], values: [1, 2, 3] }], {
				type: 'treemap',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
			})
		})
		const xml = await readEntry(zip, chartExPath(zip))
		expect(xml).toContain('<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$4</cx:f>')
		expect(xml).toContain('<cx:numDim type="size"><cx:f>Sheet1!$B$2:$B$4</cx:f>')
	})
})
