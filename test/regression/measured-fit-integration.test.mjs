// End-to-end: fit:'shrink' through the public API. The bare-flag (no metrics) and
// degrade-with-warning paths are CI-safe; the baked-fontScale assertions need a
// real font and are skipped when fc-match cannot resolve genuine Aptos.
import { readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'

async function slide1Xml(pres) {
	const buf = await pres.stream()
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

function aptosPath() {
	try {
		const out = execFileSync('fc-match', ['-f', '%{family}\t%{file}', 'Aptos'], { encoding: 'utf8' })
		const [fam, file] = out.split('\t')
		return fam && file && fam.toLowerCase().includes('aptos') ? file.trim() : null
	} catch {
		return null
	}
}

describe("measured fit: fit:'shrink' integration", () => {
	test('no registered metrics → bare <a:normAutofit/> (unchanged behavior)', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('registered metrics for a DIFFERENT face → bare flag + degrade (no throw)', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', path)
		const slide = pres.addSlide()
		// Box uses an unregistered face → unmeasurable → bare flag retained.
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Helvetica', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('registered metrics + overflow → baked fontScale < 100%', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', path)
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		const m = xml.match(/<a:normAutofit fontScale="(\d+)"/)
		expect(m).not.toBeNull()
		const scale = Number(m[1])
		// On the 2.5% grid (×1000), between the 25% floor and below 100%.
		expect(scale).toBeGreaterThanOrEqual(25000)
		expect(scale).toBeLessThan(100000)
		expect(scale % 2500).toBe(0)
	})

	test('registered metrics + text that fits → bare flag (no needless shrink)', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', path)
		const slide = pres.addSlide()
		slide.addText('Hi', { x: 1, y: 1, w: 6, h: 3, fontFace: 'Aptos', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:normAutofit/>')
		expect(xml).not.toContain('fontScale')
	})

	test('bytes source works (registerFontMetrics accepts Uint8Array)', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', new Uint8Array(readFileSync(path)))
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'shrink' })
		const xml = await slide1Xml(pres)
		expect(xml).toMatch(/<a:normAutofit fontScale="\d+"/)
	})
})

describe("measured fit: fit:'resize' integration", () => {
	test('no registered metrics → bare <a:spAutoFit/>, authored height unchanged', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'resize' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:spAutoFit/>')
		expect(firstXfrm(xml).cy).toBe(1 * EMU_PER_IN) // unchanged 1in box
	})

	test('registered metrics + overflow → box grows past the authored height', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', path)
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Aptos', fontSize: 18, fit: 'resize', valign: 'top' })
		const xml = await slide1Xml(pres)
		expect(xml).toContain('<a:spAutoFit/>')
		const { offY, cy } = firstXfrm(xml)
		expect(cy).toBeGreaterThan(1 * EMU_PER_IN) // multi-line overflow → taller box
		expect(offY).toBe(1 * EMU_PER_IN) // anchor top → origin fixed, grows downward
	})

	test('registered metrics + short text → box shrinks to fit (spAutoFit semantics)', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', path)
		const slide = pres.addSlide()
		slide.addText('Hi', { x: 1, y: 2, w: 6, h: 3, fontFace: 'Aptos', fontSize: 18, fit: 'resize', valign: 'top' })
		const xml = await slide1Xml(pres)
		const { offY, cy } = firstXfrm(xml)
		// One 18pt line + insets ≈ 30pt, far less than the authored 3in box.
		expect(cy).toBeLessThan(3 * EMU_PER_IN)
		expect(cy).toBeGreaterThan(20 * EMU_PER_PT)
		expect(offY).toBe(2 * EMU_PER_IN) // top anchor keeps the origin
	})

	test('centered anchor splits the height delta across off.y', async () => {
		const path = aptosPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('Aptos', path)
		const slide = pres.addSlide()
		// Default (no valign) resolves to centered anchor → origin shifts by half the delta.
		slide.addText('Hi', { x: 1, y: 2, w: 6, h: 3, fontFace: 'Aptos', fontSize: 18, fit: 'resize' })
		const xml = await slide1Xml(pres)
		const { offY, cy } = firstXfrm(xml)
		const delta = 3 * EMU_PER_IN - cy // positive (box shrank)
		expect(offY).toBeCloseTo(2 * EMU_PER_IN + delta / 2, -2) // top moved down by half the shrink
	})
})
