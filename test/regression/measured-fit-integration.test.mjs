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
