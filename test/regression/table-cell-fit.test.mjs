// End-to-end: TableCellProps.fit:'shrink' through the public API. PowerPoint has no
// text-autofit for table cells, so the library bakes a REDUCED literal font size onto
// the cell's runs. The no-metrics no-op path is CI-safe and always runs; the
// baked-size assertions need a real font and are skipped when none resolves.
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

const szValues = (xml) => [...xml.matchAll(/sz="(\d+)"/g)].map((m) => Number(m[1]))
const LONG = 'This is a deliberately long cell sentence that overflows a short fixed-height table row.'

/** Resolve a usable .ttf for some common Linux/CI family (opentype.js parses plain TTF). */
function usableFontPath() {
	for (const fam of ['DejaVu Sans', 'Liberation Sans', 'Noto Sans', 'Arial']) {
		try {
			const out = execFileSync('fc-match', ['-f', '%{file}', fam], {
				encoding: 'utf8',
				stdio: ['ignore', 'pipe', 'ignore'],
			})
			const file = out.trim()
			if (file && file.toLowerCase().endsWith('.ttf')) return file
		} catch {
			/* try next */
		}
	}
	return null
}

describe('measured fit: TableCellProps.fit', () => {
	test('no registered metrics → cell font size unchanged (no-op)', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addTable([[{ text: LONG, options: { fontFace: 'Aptos', fontSize: 18, fit: 'shrink' } }]], {
			x: 0.5,
			y: 0.5,
			w: 3,
			h: 0.7,
			colW: [3],
		})
		const sizes = szValues(await slide1Xml(pres))
		expect(sizes.length).toBeGreaterThan(0)
		expect(sizes.every((s) => s === 1800)).toBe(true)
	})

	test('registered metrics + overflow in a fixed-height row → baked size < authored', async () => {
		const path = usableFontPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('CellFont', new Uint8Array(readFileSync(path)))
		const slide = pres.addSlide()
		slide.addTable([[{ text: LONG, options: { fontFace: 'CellFont', fontSize: 18, fit: 'shrink' } }]], {
			x: 0.5,
			y: 0.5,
			w: 3,
			h: 0.7,
			colW: [3],
		})
		const sizes = szValues(await slide1Xml(pres))
		expect(sizes.length).toBeGreaterThan(0)
		// Every emitted run size must be below the authored 18pt (1800).
		expect(sizes.every((s) => s < 1800)).toBe(true)
		expect(Math.max(...sizes)).toBeGreaterThan(0)
	})

	test('auto-height row (no rowH / table h) → no shrink (the row grows instead)', async () => {
		const path = usableFontPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('CellFont', new Uint8Array(readFileSync(path)))
		const slide = pres.addSlide()
		// No `h` and no `rowH` → unconstrained height → nothing to shrink against.
		slide.addTable([[{ text: LONG, options: { fontFace: 'CellFont', fontSize: 18, fit: 'shrink' } }]], {
			x: 0.5,
			y: 0.5,
			colW: [3],
		})
		const sizes = szValues(await slide1Xml(pres))
		expect(sizes.every((s) => s === 1800)).toBe(true)
	})

	test('table-level fit:shrink cascades to a cell with no explicit fit', async () => {
		const path = usableFontPath()
		if (!path) return expect(true).toBe(true)
		const pres = new PptxGenJS()
		await pres.registerFontMetrics('CellFont', new Uint8Array(readFileSync(path)))
		const slide = pres.addSlide()
		slide.addTable([[{ text: LONG, options: { fontFace: 'CellFont', fontSize: 18 } }]], {
			x: 0.5,
			y: 0.5,
			w: 3,
			h: 0.7,
			colW: [3],
			fit: 'shrink',
		})
		const sizes = szValues(await slide1Xml(pres))
		expect(sizes.every((s) => s < 1800)).toBe(true)
	})
})
