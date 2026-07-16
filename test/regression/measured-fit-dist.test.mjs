// Measured-fit through the BUILT package (dist), driven by a repo-shipped font so
// the metrics-baking branches run deterministically on every platform.
//
// The sibling suites (measure-text-api / measured-fit-integration / table-cell-fit)
// either import from `src/` (which the coverage gate — scoped to `dist/**`— does not
// count) or gate their metrics-baked assertions behind an `fc-match`-resolved Aptos
// that is absent on most machines, so the whole `applyMeasuredFit` / `measureText` /
// `computeTableLayout` core stays uncovered in `dist`. Silkscreen (OFL, committed
// under test/read/fixtures/fonts) removes that dependency: opentype.js parses it, so
// `registerFontMetrics` gives us real advances and the export-time bake runs here.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'

const REG_PATH = fileURLToPath(new URL('../read/fixtures/fonts/Silkscreen-Regular.ttf', import.meta.url))
const BOLD_PATH = fileURLToPath(new URL('../read/fixtures/fonts/Silkscreen-Bold.ttf', import.meta.url))

const EMU_PER_IN = 914400
const EMU_PER_PT = 12700

// Silkscreen is a wide pixel font, so this string overflows a small box handily.
const OVERFLOW = 'The quick brown fox jumps over the lazy dog. '.repeat(6).trim()

async function slide1Xml(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	return zip.file('ppt/slides/slide1.xml').async('string')
}

/** First non-group xfrm (the spTree opens with a zero-size group). */
function firstXfrm(xml) {
	const re = /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g
	for (const m of xml.matchAll(re)) {
		const cx = Number(m[3])
		const cy = Number(m[4])
		if (cx > 0 || cy > 0) return { offY: Number(m[2]), cx, cy }
	}
	throw new Error('no non-zero xfrm found')
}

const szValues = (xml) => [...xml.matchAll(/sz="(\d+)"/g)].map((m) => Number(m[1]))

async function pptxWithSilkscreen() {
	const pres = new PptxGenJS()
	await pres.registerFontMetrics('Silkscreen', REG_PATH)
	await pres.registerFontMetrics('Silkscreen', BOLD_PATH, { bold: true })
	return pres
}

describe('measureText() through dist (Silkscreen metrics)', () => {
	test('named registered face → measurable with a wrapped multi-line height', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText(OVERFLOW, { wIn: 2, fontSize: 18, fontFace: 'Silkscreen' })
		expect(m.measurable).toBe(true)
		expect(m.lineCount).toBeGreaterThan(1)
		expect(m.heightIn).toBeGreaterThan(0)
		expect(m.widestLineIn).toBeGreaterThan(0)
	})

	test('short text at a wide width → single line', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText('Hi', { wIn: 8, fontSize: 12, fontFace: 'Silkscreen' })
		expect(m.measurable).toBe(true)
		expect(m.lineCount).toBe(1)
	})

	test('fitsBox / shrinkScaleFor helpers', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText(OVERFLOW, { wIn: 2, fontSize: 18, fontFace: 'Silkscreen' })
		expect(m.fitsBox(0.3)).toBe(false) // too short — overflows
		expect(m.fitsBox(50)).toBe(true) // absurdly tall — fits
		const scale = m.shrinkScaleFor(0.5)
		expect(scale).toBeGreaterThan(0)
		expect(scale).toBeLessThanOrEqual(100)
	})

	test('unnamed (theme-default) face → not measurable', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText('x', { wIn: 5, fontSize: 12 })
		expect(m.measurable).toBe(false)
		expect(m.fitsBox(1)).toBe(false)
		expect(m.shrinkScaleFor(1)).toBe(100)
	})

	test('overflowsBox: true when overflowing, false when it fits, false when unmeasurable', async () => {
		const pres = await pptxWithSilkscreen()
		expect(pres.overflowsBox(OVERFLOW, { wIn: 2, hIn: 0.3, fontSize: 18, fontFace: 'Silkscreen' })).toBe(true)
		expect(pres.overflowsBox('Hi', { wIn: 6, hIn: 3, fontSize: 12, fontFace: 'Silkscreen' })).toBe(false)
		// Unmeasurable (unnamed) face never false-positives.
		expect(pres.overflowsBox(OVERFLOW, { wIn: 2, hIn: 0.3, fontSize: 18 })).toBe(false)
	})

	test('run array with per-run bold uses the bold metrics variant', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText([{ text: 'Bold heading', options: { bold: true } }], {
			wIn: 3,
			fontSize: 18,
			fontFace: 'Silkscreen',
		})
		expect(m.measurable).toBe(true)
		expect(m.widestLineIn).toBeGreaterThan(0)
	})
})

describe("applyMeasuredFit: fit:'shrink' through dist export", () => {
	test('overflow → baked fontScale on the 2.5% grid, below 100%', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Silkscreen', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		const m = xml.match(/<a:normAutofit fontScale="(\d+)"/)
		expect(m).not.toBeNull()
		const scale = Number(m[1])
		expect(scale).toBeGreaterThanOrEqual(25000)
		expect(scale).toBeLessThan(100000)
		expect(scale % 2500).toBe(0)
	})

	test('text that fits → bare flag (no needless shrink)', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addText('Hi', { x: 1, y: 1, w: 6, h: 3, fontFace: 'Silkscreen', fontSize: 12, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('unregistered NAMED face → heuristic still bakes an approximate scale', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		// Deck opted into measured fit, so an unregistered *named* face falls back to the
		// conservative average-advance heuristic rather than degrading to the bare flag.
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Helvetica', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('fontScale')
	})

	test('unnamed (theme-default) face stays unmeasurable → bare flag', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('bytes source (Uint8Array) works the same as a path', async () => {
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Silkscreen', new Uint8Array(readFileSync(REG_PATH)))
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Silkscreen', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toMatch(/<a:normAutofit fontScale="\d+"/)
	})

	test('grouped overflow text still receives a baked fontScale', async () => {
		const pres = await pptxWithSilkscreen()
		pres.addSlide().addGroup([
			{ rect: { x: 1, y: 1, w: 3, h: 1 } },
			{
				text: {
					text: OVERFLOW,
					options: { x: 1, y: 1, w: 3, h: 1, fontFace: 'Silkscreen', fontSize: 18, fit: 'shrink' },
				},
			},
		])
		const xml = await slide1Xml(pres)
		expect(xml).toMatch(/<p:grpSp>.*<a:normAutofit fontScale="\d+"/s)
	})

	test('nested grouped overflow text still receives a baked fontScale', async () => {
		const pres = await pptxWithSilkscreen()
		pres.addSlide().addGroup([
			{ rect: { x: 1, y: 1, w: 3, h: 1 } },
			{
				group: {
					children: [
						{ rect: { x: 1, y: 1, w: 3, h: 1 } },
						{
							text: {
								text: OVERFLOW,
								options: {
									x: 1,
									y: 1,
									w: 3,
									h: 1,
									fontFace: 'Silkscreen',
									fontSize: 18,
									fit: 'shrink',
								},
							},
						},
					],
				},
			},
		])
		const xml = await slide1Xml(pres)
		expect((xml.match(/<p:grpSp>/g) ?? []).length).toBe(2)
		expect(xml).toMatch(/<a:normAutofit fontScale="\d+"/)
	})
})

describe("applyMeasuredFit: fit:'resize' through dist export", () => {
	test('overflow + top anchor → box grows downward, origin fixed', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, {
			x: 1,
			y: 1,
			w: 3,
			h: 1,
			fontFace: 'Silkscreen',
			fontSize: 18,
			fit: 'resize',
			valign: 'top',
		})
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:spAutoFit/>')
		const { offY, cy } = firstXfrm(xml)
		expect(cy).toBeGreaterThan(1 * EMU_PER_IN)
		expect(offY).toBe(1 * EMU_PER_IN)
	})

	test('short text + centered anchor → box shrinks, origin shifts by half the delta', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addText('Hi', { x: 1, y: 2, w: 6, h: 3, fontFace: 'Silkscreen', fontSize: 18, fit: 'resize' })
		const xml = await slide1Xml(pres)
		const { offY, cy } = firstXfrm(xml)
		expect(cy).toBeLessThan(3 * EMU_PER_IN)
		expect(cy).toBeGreaterThan(20 * EMU_PER_PT)
		const delta = 3 * EMU_PER_IN - cy
		expect(offY).toBeCloseTo(2 * EMU_PER_IN + delta / 2, -2)
	})
})

describe("applyMeasuredFit: TableCellProps.fit:'shrink' through dist export", () => {
	const LONG = 'This is a deliberately long cell sentence that overflows a short fixed-height table row.'

	test('fixed-height row + overflow → reduced literal font size baked onto the cell', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addTable([[{ text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }]], {
			x: 0.5,
			y: 0.5,
			w: 3,
			rowH: [0.4],
		})
		const xml = await slide1Xml(pres)
		const sizes = szValues(xml)
		expect(sizes.length).toBeGreaterThan(0)
		// 18pt == sz 1800; a shrink bakes something strictly smaller.
		expect(Math.min(...sizes)).toBeLessThan(1800)
	})

	test('auto-height row (no rowH / table h) → cell font size untouched (row grows instead)', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addTable([[{ text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }]], {
			x: 0.5,
			y: 0.5,
			w: 3,
		})
		const xml = await slide1Xml(pres)
		expect(szValues(xml)).toContain(1800) // authored 18pt preserved
	})

	test('table-level fit:shrink propagates to cells that set none', async () => {
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addTable([[{ text: LONG, options: { fontFace: 'Silkscreen', fontSize: 18 } }]], {
			x: 0.5,
			y: 0.5,
			w: 3,
			rowH: [0.4],
			fit: 'shrink',
		})
		const xml = await slide1Xml(pres)
		expect(Math.min(...szValues(xml))).toBeLessThan(1800)
	})
})

describe('tableLayout() through dist (Silkscreen metrics)', () => {
	test('auto-height rows estimated from real metrics → heightExact false, positive heights', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [
			[{ text: 'Header A', options: { fontFace: 'Silkscreen', fontSize: 14 } }, { text: 'Header B' }],
			[{ text: OVERFLOW, options: { fontFace: 'Silkscreen', fontSize: 14 } }, { text: 'short' }],
		]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 8, colW: [4, 4] })
		expect(res.cells).toHaveLength(4)
		expect(res.heightExact).toBe(false)
		expect(res.heightIn).toBeGreaterThan(0)
		// The wrapping OVERFLOW cell drives its row taller than the short-text header row.
		const rowH = (r) => res.cells.find((c) => c.row === r && c.col === 0).hIn
		expect(rowH(1)).toBeGreaterThan(rowH(0))
	})

	test('pinned rowH → exact heights', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [
			[{ text: 'a' }, { text: 'b' }],
			[{ text: 'c' }, { text: 'd' }],
		]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 8, rowH: [0.5, 0.75] })
		expect(res.heightExact).toBe(true)
		expect(res.heightIn).toBeCloseTo(1.25, 5)
	})
})
