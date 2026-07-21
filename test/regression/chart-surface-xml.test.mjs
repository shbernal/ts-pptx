import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries } from '../helpers.js'

// Surface is a CLASSIC (c:) chart, not a chartEx layout. Confirmed against the four surface charts
// PowerPoint authors (`Shapes.AddChart2(-1, {83|84|85|86}, …)`) and reads back as ChartType
// 83/84/85/86, the pieces specific to a surface chart are:
//   - a `<c:surface3DChart>` (3-D surface, `surface3D` default true) or `<c:surfaceChart>` (2-D
//     contour / top view, `surface3D: false`), whose first child is `<c:wireframe val="0|1"/>`;
//   - it is a 3-D scene like bar3D — it needs a SERIES axis (`<c:serAx>`) as the third axis, plus a
//     `<c:view3D>` and floor/side/back walls before the plotArea;
//   - the 2-D contour's view3D looks straight down the value axis (rotX 90).
// PowerPoint also writes a cosmetic `<c:bandFmts>` color-band list; it is optional (PowerPoint
// regenerates the bands on open — verified over COM the deck opens and reads back either way), so it
// is intentionally omitted. A surface chart has NO schema-vs-PowerPoint divergence — it validates
// cleanly. The generic classic-chart wiring (content type, chart rel, xlsx) is covered elsewhere.

const LABELS = ['A', 'B', 'C', 'D']
const DATA = [
	{ name: 'Series 1', labels: LABELS, values: [4.3, 2.5, 3.5, 4.5] },
	{ name: 'Series 2', labels: LABELS, values: [2.4, 4.4, 1.8, 2.8] },
	{ name: 'Series 3', labels: LABELS, values: [2, 2, 3, 5] },
]

async function buildSurface(extra = {}) {
	return build((p) => {
		p.addSlide().addChart(DATA, { type: 'surface', x: 1, y: 1, w: 8, h: 4.5, ...extra })
	})
}

describe('surface (classic) chart', () => {
	test('emits a classic chart part (no chartEx, no AlternateContent)', async () => {
		const { zip } = await buildSurface()
		const entries = listEntries(zip)
		expect(entries).toContain('ppt/charts/chart1.xml')
		expect(entries.some((p) => /^ppt\/charts\/chartEx\d+\.xml$/.test(p))).toBe(false)
		const contentTypes = await readEntry(zip, '[Content_Types].xml')
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
		)
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).not.toContain('<mc:AlternateContent')
	})

	test('default is a 3-D surface: surface3DChart, view3D + walls, wireframe off, 3 series over cat/val/ser axes', async () => {
		const { zip } = await buildSurface()
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect(xml).toContain('<c:surface3DChart>')
		expect(xml).not.toContain('<c:surfaceChart>')
		expect(xml).toContain('<c:wireframe val="0"/>')
		// A surface is a 3-D scene: view3D + all three walls precede the plotArea.
		expect(xml).toContain('<c:view3D>')
		expect(xml).toContain('<c:floor>')
		expect(xml).toContain('<c:sideWall>')
		expect(xml).toContain('<c:backWall>')
		expect((xml.match(/<c:ser>/g) || []).length).toBe(3)
		// Three axes: category, value, series.
		expect(xml).toContain('<c:catAx>')
		expect(xml).toContain('<c:valAx>')
		expect(xml).toContain('<c:serAx>')
		// The surface plot references all three axis ids.
		expect(xml).toMatch(
			/<c:surface3DChart>[\s\S]*<c:axId val="2094734554"\/><c:axId val="2094734552"\/><c:axId val="2094734556"\/>[\s\S]*<\/c:surface3DChart>/
		)
		// The cosmetic color-band list is omitted (PowerPoint regenerates it on open).
		expect(xml).not.toContain('<c:bandFmts>')
	})

	test('surfaceWireframe draws the mesh (wireframe val 1)', async () => {
		const { zip } = await buildSurface({ surfaceWireframe: true })
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect(xml).toContain('<c:surface3DChart>')
		expect(xml).toContain('<c:wireframe val="1"/>')
	})

	test('surface3D:false is a 2-D contour: surfaceChart, top-down view3D (rotX 90)', async () => {
		const { zip } = await buildSurface({ surface3D: false })
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect(xml).toContain('<c:surfaceChart>')
		expect(xml).not.toContain('<c:surface3DChart>')
		// The contour looks straight down the value axis.
		expect(xml).toContain('<c:rotX val="90"/>')
		expect(xml).toContain('<c:perspective val="0"/>')
		// A contour still needs its series axis.
		expect(xml).toContain('<c:serAx>')
	})
})
