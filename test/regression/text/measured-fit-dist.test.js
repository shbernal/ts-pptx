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
import TsPptx from '../../../dist/node.js'
// The `ts-pptx/measure` entry publishes the calibrated constants the bake uses, so a test
// can state "inflated by the height safety factor" instead of re-pinning its value here.
import { HEIGHT_SAFETY_FACTOR } from '../../../dist/measure.js'

const REG_PATH = fileURLToPath(new URL('../../read/fixtures/fonts/Silkscreen-Regular.ttf', import.meta.url))
const BOLD_PATH = fileURLToPath(new URL('../../read/fixtures/fonts/Silkscreen-Bold.ttf', import.meta.url))

const EMU_PER_IN = 914400
const EMU_PER_PT = 12700

// Silkscreen is a wide pixel font, so this string overflows a small box handily.
const OVERFLOW = 'The quick brown fox jumps over the lazy dog. '.repeat(6).trim()

async function slide1Xml(pres) {
	const buf = await pres.toBytes()
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
	const pres = new TsPptx()
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
		const pres = new TsPptx()
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

// The input-shape matrix. Everything above drives one canonical text box; these drive the
// *option* surface that reshapes the measured box or the paragraph list before the solver
// sees it — `\n` / `breakLine` splitting, `margin` insets, `wrap:false`, the vertical anchor,
// line spacing, and the degenerate boxes that make a box unmeasurable. Each assertion is a
// comparison against the same deck without the option, so it pins the option's *effect*
// rather than a magic number: the model's constants are calibrated elsewhere
// (autofit-calibration-oracle) and must stay free to move.

/** Baked `<a:normAutofit fontScale>` (thousandths of a percent), or null for the bare flag. */
function bakedScale(xml) {
	const m = xml.match(/<a:normAutofit fontScale="(\d+)"/)
	return m ? Number(m[1]) : null
}

/** Build one Silkscreen text box with `fit:'shrink'` and return its baked scale. */
async function shrinkScale(extraOpts, text = OVERFLOW) {
	const pres = await pptxWithSilkscreen()
	pres.addSlide().addText(text, {
		x: 1,
		y: 1,
		w: 3,
		h: 1,
		fontFace: 'Silkscreen',
		fontSize: 18,
		fit: 'shrink',
		...extraOpts,
	})
	return bakedScale(await slide1Xml(pres))
}

describe('measured fit: paragraph splitting', () => {
	test('a "\\n" inside a run starts a new paragraph (more lines, more height)', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 8, fontSize: 18, fontFace: 'Silkscreen' }
		const oneLine = pres.measureText('alpha beta', opts)
		const twoLines = pres.measureText('alpha\nbeta', opts)
		// Wide enough that neither wraps — so the extra line is the newline's doing, not the wrap.
		expect(oneLine.lineCount).toBe(1)
		expect(twoLines.lineCount).toBe(2)
		expect(twoLines.heightIn).toBeGreaterThan(oneLine.heightIn)
	})

	test('"\\r\\n" is normalized to the same split as "\\n"', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 8, fontSize: 18, fontFace: 'Silkscreen' }
		expect(pres.measureText('alpha\r\nbeta', opts).heightIn).toBe(pres.measureText('alpha\nbeta', opts).heightIn)
	})

	test('`breakLine` on a run ends its paragraph just as a "\\n" does', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 8, fontSize: 18, fontFace: 'Silkscreen' }
		const broken = pres.measureText([{ text: 'alpha', options: { breakLine: true } }, { text: 'beta' }], opts)
		const joined = pres.measureText([{ text: 'alpha' }, { text: 'beta' }], opts)
		expect(broken.lineCount).toBe(2)
		expect(joined.lineCount).toBe(1)
	})

	test('a trailing "\\n" leaves the final empty paragraph in the count', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 8, fontSize: 18, fontFace: 'Silkscreen' }
		expect(pres.measureText('alpha\n', opts).lineCount).toBe(2)
	})

	test('the split reaches the export bake: the same text needs a smaller scale when broken', async () => {
		const flowed = await shrinkScale({}, 'alpha beta gamma delta epsilon zeta')
		const broken = await shrinkScale({}, 'alpha\nbeta\ngamma\ndelta\nepsilon\nzeta')
		expect(broken).not.toBeNull()
		expect(broken).toBeLessThan(flowed)
	})
})

describe('measured fit: text-frame insets from `margin`', () => {
	test('a larger `margin` shrinks the inner box, so the text shrinks further', async () => {
		const none = await shrinkScale({ margin: 0 })
		const roomy = await shrinkScale({ margin: 0.4 })
		expect(none).not.toBeNull()
		expect(roomy).toBeLessThan(none)
	})

	test('`margin` as a [T,R,B,L] array matches the equivalent scalar', async () => {
		expect(await shrinkScale({ margin: [0.4, 0.4, 0.4, 0.4] })).toBe(await shrinkScale({ margin: 0.4 }))
	})

	test('the array is read as [T,R,B,L] — a top/bottom pair is not a left/right pair', async () => {
		// Vertical margin eats the height the shrink solver is fitting against; horizontal
		// margin eats width and forces earlier wrapping. Both shrink, but not by the same
		// amount — which is what proves the four slots are not interchangeable.
		const vertical = await shrinkScale({ margin: [0.4, 0, 0.4, 0] })
		const horizontal = await shrinkScale({ margin: [0, 0.4, 0, 0.4] })
		expect(vertical).not.toBe(horizontal)
	})

	test('`margin` larger than the box makes it unmeasurable → bare flag, no bogus scale', async () => {
		// Insets exceeding w/h drive innerWidthPt/innerHeightPt negative. Baking a scale off a
		// negative box would emit a nonsense fontScale; the box must fall back to the bare flag.
		expect(await shrinkScale({ w: 1, h: 0.5, margin: 0.6 })).toBeNull()
	})
})

describe('measured fit: degenerate boxes', () => {
	test('a zero-width box is not measured (bare flag)', async () => {
		expect(await shrinkScale({ w: 0 })).toBeNull()
	})

	test('a zero-height box is not measured (bare flag)', async () => {
		expect(await shrinkScale({ h: 0 })).toBeNull()
	})

	test('a box with room vertically but none horizontally is not measured either', async () => {
		// Left+right insets exceed `w` while top+bottom leave `h` healthy, so the width check
		// is the only thing standing between the solver and a negative layout width.
		expect(await shrinkScale({ w: 1, h: 5, margin: 0.6 })).toBeNull()
	})

	test('empty text is not measured (bare flag)', async () => {
		expect(await shrinkScale({}, '')).toBeNull()
	})
})

describe('measured fit: `wrap: false`', () => {
	test('a non-wrapping line that overflows horizontally still shrinks', async () => {
		// The point of the wrap:false handling. With wrap:false the text is one line, so it
		// always fits the box HEIGHT and `normAutofit`'s vertical test never fires — the line
		// just spills out the side. The solver has to check the widest line against the box
		// WIDTH instead. The box here is deliberately tall enough that height alone would say
		// "fits", so a baked scale can only have come from the width check.
		const scale = await shrinkScale({ wrap: false, w: 2, h: 6 }, 'nowrapping single line of text')
		expect(scale).not.toBeNull()
		expect(scale).toBeLessThan(100000)
	})

	test('the same box wrapping normally needs no shrink', async () => {
		expect(await shrinkScale({ w: 2, h: 6 }, 'nowrapping single line of text')).toBeNull()
	})
})

describe("measured fit: fit:'resize' vertical anchor", () => {
	test('bottom anchor → the box grows upward, its bottom edge fixed', async () => {
		const pres = await pptxWithSilkscreen()
		pres.addSlide().addText(OVERFLOW, {
			x: 1,
			y: 2,
			w: 3,
			h: 1,
			fontFace: 'Silkscreen',
			fontSize: 18,
			fit: 'resize',
			valign: 'bottom',
		})
		const { offY, cy } = firstXfrm(await slide1Xml(pres))
		expect(cy).toBeGreaterThan(1 * EMU_PER_IN)
		// Top moves up by the whole delta, so the bottom edge lands where it was authored.
		expect(offY).toBeLessThan(2 * EMU_PER_IN)
		expect(offY + cy).toBeCloseTo(3 * EMU_PER_IN, -2)
	})

	test("fit:'resize' with an unmeasurable face leaves the authored height alone", async () => {
		const pres = await pptxWithSilkscreen()
		// Registry is non-empty (the deck opted in), but an unnamed face cannot be guessed —
		// so resize has nothing to bake and must not touch the box.
		pres.addSlide().addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontSize: 18, fit: 'resize' })
		const { offY, cy } = firstXfrm(await slide1Xml(pres))
		expect(cy).toBe(1 * EMU_PER_IN)
		expect(offY).toBe(1 * EMU_PER_IN)
	})
})

describe('measured fit: line spacing', () => {
	test('exact `lineSpacing` (points) overrides the calibrated single-line pitch', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 2, fontSize: 18, fontFace: 'Silkscreen' }
		const natural = pres.measureText(OVERFLOW, opts)
		const wide = pres.measureText(OVERFLOW, { ...opts, lineSpacing: 60 })
		expect(wide.lineCount).toBe(natural.lineCount) // spacing does not change wrapping
		expect(wide.heightIn).toBeGreaterThan(natural.heightIn)
	})

	test('`lineSpacingMultiple` scales the pitch proportionally', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 2, fontSize: 18, fontFace: 'Silkscreen' }
		const single = pres.measureText(OVERFLOW, opts)
		const double = pres.measureText(OVERFLOW, { ...opts, lineSpacingMultiple: 2 })
		expect(double.heightIn).toBeCloseTo(single.heightIn * 2, 5)
	})

	test('paragraph space before/after adds to the measured height', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 8, fontSize: 18, fontFace: 'Silkscreen' }
		const bare = pres.measureText('alpha\nbeta', opts)
		const spaced = pres.measureText('alpha\nbeta', { ...opts, paraSpaceBefore: 12, paraSpaceAfter: 6 })
		// Two paragraphs x (12 + 6)pt = 36pt = 0.5in, and the reported height carries the
		// conservative height inflation — which applies to the paragraph spacing as well as
		// to the line pitch. Naming the factor rather than hardcoding 0.52 keeps this test
		// from re-pinning a calibration constant that autofit-calibration-oracle owns.
		expect(spaced.heightIn - bare.heightIn).toBeCloseTo(0.5 * HEIGHT_SAFETY_FACTOR, 5)
	})

	test('wider line spacing reaches the export bake', async () => {
		const natural = await shrinkScale({})
		const wide = await shrinkScale({ lineSpacingMultiple: 2 })
		expect(wide).toBeLessThan(natural)
	})
})

describe('measured fit: solver floor and edges', () => {
	test('text that cannot fit even at the floor bakes the floor rather than giving up', async () => {
		const scale = await shrinkScale({ w: 1, h: 0.4 }, OVERFLOW.repeat(4))
		expect(scale).toBe(25000) // MIN_FONT_SCALE_PCT
	})

	test('an over-long unbreakable word character-wraps instead of overflowing', async () => {
		const pres = await pptxWithSilkscreen()
		// No whitespace to break at, so the greedy layout has to split mid-word. Without the
		// character-wrap fallback this measures as a single (enormous) line and reports a
		// height the box would satisfy while the text visibly spills.
		const m = pres.measureText('W'.repeat(120), { wIn: 1.5, fontSize: 18, fontFace: 'Silkscreen' })
		expect(m.lineCount).toBeGreaterThan(1)
		expect(m.widestLineIn).toBeLessThanOrEqual(1.5)
	})

	test('a tab is measured as a space-width gap, not as zero width', async () => {
		const pres = await pptxWithSilkscreen()
		const opts = { wIn: 20, fontSize: 18, fontFace: 'Silkscreen' }
		const tabbed = pres.measureText('alpha\tbeta', opts)
		const joined = pres.measureText('alphabeta', opts)
		expect(tabbed.lineCount).toBe(1)
		expect(tabbed.widestLineIn).toBeGreaterThan(joined.widestLineIn)
	})

	test('shrinkScaleFor returns 100 when the text already fits', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText('Hi', { wIn: 8, fontSize: 12, fontFace: 'Silkscreen' })
		expect(m.shrinkScaleFor(50)).toBe(100)
	})

	test('an unregistered named face is measured heuristically and reported', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText(OVERFLOW, { wIn: 2, fontSize: 18, fontFace: 'Helvetica' })
		expect(m.measurable).toBe(true)
		expect(m.approximatedFaces).toEqual(['Helvetica'])
		// The snapshot is caller-owned: re-entering the resolver must not grow the array.
		m.shrinkScaleFor(0.5)
		expect(m.approximatedFaces).toEqual(['Helvetica'])
	})

	test('an empty run list is unmeasurable', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText([], { wIn: 8, fontSize: 12, fontFace: 'Silkscreen' })
		expect(m.measurable).toBe(false)
		expect(m.lineCount).toBe(0)
	})

	test('a zero-width measure box is unmeasurable rather than a divide-by-zero', async () => {
		const pres = await pptxWithSilkscreen()
		const m = pres.measureText('alpha', { wIn: 0, fontSize: 12, fontFace: 'Silkscreen' })
		expect(m.measurable).toBe(false)
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

	test('rowH as a scalar pins every row to the same height', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [[{ text: 'a' }], [{ text: 'b' }], [{ text: 'c' }]]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4, rowH: 0.5 })
		expect(res.heightExact).toBe(true)
		expect(res.heightIn).toBeCloseTo(1.5, 5)
	})

	test('table `h` with no rowH splits evenly across the rows', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [[{ text: 'a' }], [{ text: 'b' }]]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4, h: 3 })
		expect(res.heightExact).toBe(true)
		expect(res.heightIn).toBeCloseTo(3, 4)
		expect(res.cells[0].hIn).toBeCloseTo(1.5, 4)
	})

	test('a rowH array slot that is missing falls back to the even split of table `h`', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [[{ text: 'a' }], [{ text: 'b' }]]
		// Row 0 is pinned at 0.5in; row 1 has no slot, so it takes h/numRows rather than
		// collapsing to auto-height (which would silently drop the caller's `h`).
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4, h: 3, rowH: [0.5] })
		expect(res.heightExact).toBe(true)
		expect(res.cells[0].hIn).toBeCloseTo(0.5, 4)
		expect(res.cells[1].hIn).toBeCloseTo(1.5, 4)
	})

	test('omitted x/y/w default to the origin and 75% of the slide', async () => {
		const pres = await pptxWithSilkscreen()
		const res = pres.tableLayout([[{ text: 'a' }, { text: 'b' }]], {})
		expect(res.cells[0].xIn).toBe(0)
		expect(res.cells[0].yIn).toBe(0)
		expect(res.widthIn).toBeCloseTo(10 * 0.75, 4) // 10in slide default
	})

	test('empty input returns an empty result rather than throwing', async () => {
		const pres = await pptxWithSilkscreen()
		for (const rows of [[], [[]]]) {
			const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4 })
			expect(res).toEqual({ cells: [], widthIn: 0, heightIn: 0, heightExact: true })
		}
	})

	test('a row hole (sparse array from a JS caller) is skipped, not treated as a row of cells', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [[{ text: 'a' }]]
		rows[2] = [{ text: 'c' }] // leaves rows[1] as a hole
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4, rowH: 0.5 })
		expect(res.cells).toHaveLength(2)
		expect(res.cells.map((c) => c.row)).toEqual([0, 2])
	})

	test('cells beyond the grid width established by row 0 are dropped', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [
			[{ text: 'a' }, { text: 'b' }],
			[{ text: 'c' }, { text: 'd' }, { text: 'e' }], // 'e' has no column to land in
		]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4, rowH: 0.5 })
		expect(res.cells).toHaveLength(4)
	})

	test('a row covered entirely by a rowspan still gets a default line of height', async () => {
		const pres = await pptxWithSilkscreen()
		// Row 1 has no cell of its own — every column is occupied by the rowspan above it.
		// Without a floor it would be zero-height and the spanned cell would lose half its box.
		const rows = [[{ text: 'a', options: { rowspan: 2 } }], []]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4 })
		expect(res.cells).toHaveLength(1)
		expect(res.heightIn).toBeGreaterThan(0)
		expect(res.cells[0].hIn).toBe(res.heightIn) // the span covers both rows
	})

	test('a cell narrower than its own margins falls back to one line, not a negative height', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [[{ text: OVERFLOW, options: { margin: 0.9 } }]]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 1 })
		expect(res.heightIn).toBeGreaterThan(0)
	})

	test('a cell with no text at all measures as one line', async () => {
		const pres = await pptxWithSilkscreen()
		const res = pres.tableLayout([[{}, { text: 'b' }]], { x: 1, y: 1, w: 4 })
		expect(res.cells).toHaveLength(2)
		expect(res.cells[0].hIn).toBeGreaterThan(0)
	})

	test('a rowH array hole with no table `h` leaves that row auto-height', async () => {
		const pres = await pptxWithSilkscreen()
		const rows = [[{ text: 'a' }], [{ text: 'b' }]]
		const res = pres.tableLayout(rows, { x: 1, y: 1, w: 4, rowH: [0.5] })
		expect(res.cells[0].heightExact).toBe(true)
		expect(res.cells[1].heightExact).toBe(false) // estimated, not pinned
		expect(res.heightExact).toBe(false)
	})

	test('a cell whose text is a single run object is normalized like a one-run array', async () => {
		// `normalizeRuns` documents three shapes — string, TextProps, TextProps[]. `addTable`
		// only ever produces the first and third, so the middle one is reachable only here,
		// where the caller's rows go in unnormalized.
		const pres = await pptxWithSilkscreen()
		const opts = { fontFace: 'Silkscreen', fontSize: 14 }
		// @ts-expect-error `TableCell.text` is typed string | TableCell[]; a single run object
		// is the untyped-caller shape normalizeRuns still documents and handles
		const single = pres.tableLayout([[{ text: { text: OVERFLOW, options: opts } }]], { x: 1, y: 1, w: 3 })
		const asArray = pres.tableLayout([[{ text: [{ text: OVERFLOW, options: opts }] }]], { x: 1, y: 1, w: 3 })
		expect(single.heightIn).toBe(asArray.heightIn)
		expect(single.heightIn).toBeGreaterThan(0)
	})

	test('a cell whose text is neither a string nor runs measures as one line', async () => {
		const pres = await pptxWithSilkscreen()
		// @ts-expect-error verifying an untyped caller's junk `text` degrades to one line
		const res = pres.tableLayout([[{ text: true }, { text: 'b' }]], { x: 1, y: 1, w: 4 })
		expect(res.cells[0].hIn).toBeGreaterThan(0)
	})

	test('a scalar cell `margin` is applied to all four sides', async () => {
		const pres = await pptxWithSilkscreen()
		const tight = pres.tableLayout([[{ text: OVERFLOW, options: { margin: 0 } }]], { x: 1, y: 1, w: 4 })
		const padded = pres.tableLayout([[{ text: OVERFLOW, options: { margin: 0.3 } }]], { x: 1, y: 1, w: 4 })
		// Less inner width (more wrapping) and more vertical padding both add height.
		expect(padded.heightIn).toBeGreaterThan(tight.heightIn)
	})

	test('a malformed cell `margin` falls back to the default rather than producing NaN geometry', async () => {
		const pres = await pptxWithSilkscreen()
		const good = pres.tableLayout([[{ text: OVERFLOW }]], { x: 1, y: 1, w: 4 })
		// Wrong length, wrong element type, and a NaN slot — the three ways an untyped caller
		// reaches `resolveCellInsetsEmu` with something it cannot turn into four insets.
		for (const margin of [
			[0.1, 0.1],
			['a', 'b', 'c', 'd'],
			[0.1, 0.1, 0.1, NaN],
		]) {
			// @ts-expect-error deliberately malformed `Margin` from an untyped caller
			const res = pres.tableLayout([[{ text: OVERFLOW, options: { margin } }]], { x: 1, y: 1, w: 4 })
			expect(res.heightIn).toBe(good.heightIn)
		}
	})
})

describe('applyMeasuredFit: the table cell shrink pass', () => {
	const LONG = 'This is a deliberately long cell sentence that overflows a short fixed-height table row.'

	/** Smallest `sz` (hundredths of a point) emitted on slide 1. */
	const minSz = (xml) => Math.min(...szValues(xml))

	async function tableDeck(rows, opts) {
		const pres = await pptxWithSilkscreen()
		pres.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 3, ...opts })
		return slide1Xml(pres)
	}

	test('a cell without fit is left alone while its shrinking neighbor is baked', async () => {
		const xml = await tableDeck(
			[
				[
					{ text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } },
					{ text: LONG, options: { fontFace: 'Silkscreen', fontSize: 18 } },
				],
			],
			{ w: 6, rowH: [0.4] }
		)
		const sizes = szValues(xml)
		expect(sizes).toContain(1800) // the opted-out cell keeps its authored size
		expect(Math.min(...sizes)).toBeLessThan(1800) // the opted-in one shrank
	})

	test('a fixed row height from table `h` (not rowH) still drives the shrink', async () => {
		const xml = await tableDeck([[{ text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }]], {
			h: 0.4,
		})
		expect(minSz(xml)).toBeLessThan(1800)
	})

	test('a scalar rowH pins the row the same way an array does', async () => {
		const cell = { text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }
		const scalar = await tableDeck([[{ ...cell }]], { rowH: 0.4 })
		const array = await tableDeck([[{ ...cell }]], { rowH: [0.4] })
		expect(minSz(scalar)).toBe(minSz(array))
	})

	test('a table with no `w` is measured against the 75% default width', async () => {
		const xml = await tableDeck([[{ text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }]], {
			w: undefined,
			rowH: [0.4],
		})
		expect(minSz(xml)).toBeLessThan(1800)
	})

	test('per-run font sizes in a rich cell are each scaled, not just the cell default', async () => {
		// A cell whose `text` is a run array: `scaleCellFontSizes` has to walk the runs and
		// clone each one. Mutating them in place would corrupt every other cell sharing the
		// table's options object, and skipping them would leave the runs at their authored
		// size while the cell default shrank — a visibly mixed-size cell.
		const xml = await tableDeck(
			[
				[
					{
						text: [
							{ text: LONG, options: { fontSize: 18 } },
							{ text: LONG, options: { fontSize: 24 } },
							{ text: LONG }, // no size of its own — inherits the cell default
						],
					},
				],
			],
			{ rowH: [0.4], fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 }
		)
		const sizes = szValues(xml)
		expect(sizes.every((s) => s < 1800)).toBe(true)
		expect(sizes).not.toContain(2400)
	})

	test('a cell narrower than its margins is skipped rather than shrunk against a negative box', async () => {
		const xml = await tableDeck(
			[[{ text: LONG, options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18, margin: 0.9 } }]],
			{ w: 1, rowH: [0.4] }
		)
		expect(szValues(xml)).toContain(1800)
	})

	test('an empty cell is still measured — its blank line has to fit the row too', async () => {
		// `addTable` normalizes a cell with no `text` to '' (define/table.ts), so an empty cell
		// is a one-line cell, not an unmeasurable one: at 18pt its line is taller than this
		// 0.4in row and the pass shrinks it like any other. The unmeasurable-cell guard is
		// reachable only through `tableLayout()`, which takes caller rows unnormalized —
		// covered above by "a cell with no text at all measures as one line".
		const xml = await tableDeck([[{ options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }]], {
			rowH: [0.4],
		})
		expect(Math.max(...szValues(xml))).toBeLessThan(1800)
	})

	test('a cell with an unmeasurable face keeps its authored size', async () => {
		// The deck registered metrics, so the table pass runs — but this cell names no face,
		// and an unnamed theme-default face cannot be guessed. Baking a guess here would
		// shrink text that may well have fit.
		const xml = await tableDeck([[{ text: LONG, options: { fit: 'shrink', fontSize: 18 } }]], { rowH: [0.4] })
		expect(szValues(xml)).toContain(1800)
	})

	test('a cell that already fits keeps its authored size', async () => {
		const xml = await tableDeck([[{ text: 'Hi', options: { fit: 'shrink', fontFace: 'Silkscreen', fontSize: 18 } }]], {
			rowH: [2],
		})
		expect(szValues(xml)).toContain(1800)
	})

	test('a non-text, non-table object on the slide is passed over', async () => {
		// The fit pass dispatches on `_type`; anything that is not text/table/group has no
		// measurable box and must not be walked into.
		const pres = await pptxWithSilkscreen()
		const slide = pres.addSlide()
		slide.addImage({ data: 'image/png;base64,iVBORw0KGgo=', x: 1, y: 1, w: 1, h: 1 })
		slide.addText(OVERFLOW, { x: 1, y: 3, w: 3, h: 1, fontFace: 'Silkscreen', fontSize: 18, fit: 'shrink' })
		expect(await slide1Xml(pres)).toMatch(/<a:normAutofit fontScale="\d+"/)
	})
})
