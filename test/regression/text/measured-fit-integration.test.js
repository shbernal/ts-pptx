// End-to-end: fit:'shrink' through the public API. The bare-flag (no metrics) and
// degrade-with-warning paths need no font at all; the baked-fontScale and resized-box
// assertions hand a real font file to `registerFontMetrics`, which takes a path or bytes
// and so cannot be fed the recorded-advance sidecar the read-side oracles fall back to.
//
// Every case runs in every lane. What the measured arm asserts is the SHAPE of the bake —
// a scale inside the 25%..100% band and on the 2.5% grid, a box that grew or shrank in the
// right direction — none of which is a claim about one face's advances. So it measures
// genuine Aptos where the machine has it and the committed Silkscreen fixture otherwise,
// rather than gating on a face no hosted runner will ever carry. Gating on it is what let
// nine of these report PASSED everywhere: on Linux, which resolves Aptos to DejaVu and
// rejects it, and on Windows, where there is no `fc-match` to answer at all.
import { readFileSync } from 'node:fs'
import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../../dist/node.js'
import { resolveGenuineFontFile } from '../../read/font-oracle.js'

async function slide1Xml(pres) {
	const buf = await pres.toBytes()
	const zip = await JSZip.loadAsync(buf)
	return zip.file('ppt/slides/slide1.xml').async('string')
}

const EMU_PER_IN = 914400
const EMU_PER_PT = 12700

/**
 * Extract the first real shape's xfrm off.y / ext.cy (EMU). The spTree opens with a
 * group `<a:off 0 0/><a:ext 0 0/>`, so skip the zero-size group and take the first
 * xfrm with a non-zero extent.
 */
function firstXfrm(xml) {
	const re = /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g
	for (const m of xml.matchAll(re)) {
		const cx = Number(m[3])
		const cy = Number(m[4])
		if (cx > 0 || cy > 0) return { offY: Number(m[2]), cx, cy }
	}
	throw new Error('no non-zero xfrm found')
}

const OVERFLOW = 'The quick brown fox jumps over the lazy dog. '.repeat(8).trim()

// Aptos comes through the same lookup the read-side oracles use — the only one that finds
// the per-user Office install, and the only one that answers on Windows at all.
const APTOS = resolveGenuineFontFile({ family: 'Aptos' })
const FACE = APTOS ? 'Aptos' : 'Silkscreen'
const FACE_FILE = APTOS ?? 'test/read/fixtures/fonts/Silkscreen-Regular.ttf'

describe("measured fit: fit:'shrink' integration", () => {
	test('no registered metrics → bare <a:normAutofit/> (unchanged behavior)', async () => {
		const pres = new TsPptx()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('DOCUMENTED DIVERGENCE: empty registry → measureText predicts a shrink the export does not bake', async () => {
		// Intentional, not a bug (docs/measured-text-fit.md "Unregistered-font heuristic"):
		// applyMeasuredFit reads "no metrics" as "never opted into measured fit" and bakes
		// nothing, while measureText is a read-only query that stays useful with zero setup.
		// Pinned so the asymmetry cannot change silently; approximatedFaces is how a caller
		// that needs exact numbers detects it.
		const pres = new TsPptx()
		const m = pres.measureText(OVERFLOW, { wIn: 3, fontSize: 18, fontFace: 'Aptos' })
		expect(m.measurable).toBe(true)
		expect(m.approximatedFaces).toEqual(['Aptos']) // nothing registered → the face was guessed
		expect(m.shrinkScaleFor(1)).toBeLessThan(100) // the query predicts a shrink…

		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>') // …but the export bakes the bare flag
		expect(xml).not.toContain('fontScale')
	})

	test('registered metrics for a DIFFERENT named face → heuristic shrink (P3: approximate, no throw)', async () => {
		const pres = new TsPptx()
		// The deck has opted into measured fit (some face registered), so an unregistered
		// *named* face now falls back to the conservative average-advance heuristic and
		// still bakes an approximate fontScale rather than degrading to the bare flag.
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Helvetica', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('fontScale')
		expect(xml).not.toContain('<a:normAutofit/>') // bare flag replaced by the baked (heuristic) scale
	})

	test('unnamed (theme-default) face stays unmeasurable → bare flag (heuristic does not guess the face)', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		// No fontFace → we cannot know which face the theme resolves to, so no heuristic.
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('registered metrics + overflow → baked fontScale < 100%', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: FACE, fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		const m = xml.match(/<a:normAutofit fontScale="(\d+)"/)
		expect(m).not.toBeNull()
		const scale = Number(m[1])
		// On the 2.5% grid (×1000), between the 25% floor and below 100%.
		expect(scale).toBeGreaterThanOrEqual(25000)
		expect(scale).toBeLessThan(100000)
		expect(scale % 2500).toBe(0)
	})

	test('wrap:false single line too wide → horizontal fontScale baked (dn-autofit-shrink-horizontal)', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		// One long line, no wrap, in a box that is plenty TALL (2in) but too NARROW
		// (2in): the single line fits the height and only the horizontal check catches it.
		const ONE_LINE = 'This single line is deliberately far too wide to ever fit the narrow box'
		slide.addText(ONE_LINE, { x: 1, y: 1, w: 2, h: 2, fontFace: FACE, fontSize: 28, fit: 'shrink', wrap: false })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('wrap="none"') // it is genuinely a non-wrapping frame
		const m = xml.match(/<a:normAutofit fontScale="(\d+)"/)
		expect(m).not.toBeNull() // before the fix this was a bare <a:normAutofit/>
		const scale = Number(m[1])
		expect(scale).toBeGreaterThanOrEqual(25000)
		expect(scale).toBeLessThan(100000)
		expect(scale % 2500).toBe(0)
		expect(firstXfrm(xml).cy).toBe(2 * EMU_PER_IN) // shrink, not resize: box height unchanged
	})

	test('registered metrics + text that fits → bare flag (no needless shrink)', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		slide.addText('Hi', { x: 1, y: 1, w: 6, h: 3, fontFace: FACE, fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('bytes source works (registerFontMetrics accepts Uint8Array)', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, new Uint8Array(readFileSync(FACE_FILE)))
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: FACE, fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toMatch(/<a:normAutofit fontScale="\d+"/)
	})
})

describe("measured fit: fit:'resize' integration", () => {
	test('no registered metrics → bare <a:spAutoFit/>, authored height unchanged', async () => {
		const pres = new TsPptx()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'resize' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:spAutoFit/>')
		expect(firstXfrm(xml).cy).toBe(1 * EMU_PER_IN) // unchanged 1in box
	})

	test('registered metrics + overflow → box grows past the authored height', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: FACE, fontSize: 18, fit: 'resize', valign: 'top' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:spAutoFit/>')
		const { offY, cy } = firstXfrm(xml)
		expect(cy).toBeGreaterThan(1 * EMU_PER_IN) // multi-line overflow → taller box
		expect(offY).toBe(1 * EMU_PER_IN) // anchor top → origin fixed, grows downward
	})

	test('registered metrics + short text → box shrinks to fit (spAutoFit semantics)', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		slide.addText('Hi', { x: 1, y: 2, w: 6, h: 3, fontFace: FACE, fontSize: 18, fit: 'resize', valign: 'top' })
		const xml = await slide1Xml(pres)
		const { offY, cy } = firstXfrm(xml)
		// One 18pt line + insets ≈ 30pt, far less than the authored 3in box.
		expect(cy).toBeLessThan(3 * EMU_PER_IN)
		expect(cy).toBeGreaterThan(20 * EMU_PER_PT)
		expect(offY).toBe(2 * EMU_PER_IN) // top anchor keeps the origin
	})

	test('centered anchor splits the height delta across off.y', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics(FACE, FACE_FILE)
		const slide = pres.addSlide()
		// Default (no valign) resolves to centered anchor → origin shifts by half the delta.
		slide.addText('Hi', { x: 1, y: 2, w: 6, h: 3, fontFace: FACE, fontSize: 18, fit: 'resize' })
		const xml = await slide1Xml(pres)
		const { offY, cy } = firstXfrm(xml)
		const delta = 3 * EMU_PER_IN - cy // positive (box shrank)
		expect(offY).toBeCloseTo(2 * EMU_PER_IN + delta / 2, -2) // top moved down by half the shrink
	})
})
