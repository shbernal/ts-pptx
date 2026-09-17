// End-to-end: TableCellProps.fit:'shrink' through the public API. PowerPoint has no
// text-autofit for table cells, so the library bakes a REDUCED literal font size onto
// the cell's runs.
//
// Every case here runs on every machine. What the measured ones assert is the SHAPE of the
// bake — below the authored size, or untouched — not a particular face's advances, so the
// registered face is the committed Silkscreen fixture rather than whatever fontconfig
// happens to answer with. Resolving the font off the machine is what let all three of them
// report PASSED on a runner that had measured nothing.
import { readFileSync } from 'node:fs'
import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../../dist/node.js'

async function slide1Xml(pres) {
	const buf = await pres.toBytes()
	const zip = await JSZip.loadAsync(buf)
	return zip.file('ppt/slides/slide1.xml').async('string')
}

const szValues = (xml) => [...xml.matchAll(/sz="(\d+)"/g)].map((m) => Number(m[1]))
const LONG = 'This is a deliberately long cell sentence that overflows a short fixed-height table row.'

/** A committed face, so the measured arm of this file has no machine-dependent input. */
const CELL_FONT = new Uint8Array(readFileSync('test/read/fixtures/fonts/Silkscreen-Regular.ttf'))

describe('measured fit: TableCellProps.fit', () => {
	test('no registered metrics → cell font size unchanged (no-op)', async () => {
		const pres = new TsPptx()
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
		const pres = new TsPptx()
		await pres.registerFontMetrics('CellFont', CELL_FONT)
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
		const pres = new TsPptx()
		await pres.registerFontMetrics('CellFont', CELL_FONT)
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

	test('a tiny cell shrinks no further than the smallest size a run can carry', async () => {
		// The shrink had no minimum, so a 1pt cell taken to the 25% scale floor was baked as 0.2pt.
		// The emitter's clamp then caught it and warned "fontSize 0.2 is outside the valid range" --
		// a value the caller never wrote and could do nothing about. Their 1pt was perfectly valid.
		const warnings = []
		const originalWarn = console.warn
		console.warn = (...args) => warnings.push(String(args[0]))
		let sizes
		try {
			const pres = new TsPptx()
			await pres.registerFontMetrics('CellFont', CELL_FONT)
			const slide = pres.addSlide()
			slide.addTable([[{ text: LONG, options: { fontFace: 'CellFont', fontSize: 1, fit: 'shrink' } }]], {
				x: 0.5,
				y: 0.5,
				w: 0.6,
				h: 0.2,
				colW: [0.6],
				rowH: 0.2,
			})
			sizes = szValues(await slide1Xml(pres))
		} finally {
			console.warn = originalWarn
		}
		expect(sizes.length).toBeGreaterThan(0)
		// 100 hundredths is 1pt, the bottom of ST_TextFontSize.
		expect(Math.min(...sizes)).toBe(100)
		expect(warnings.filter((w) => w.includes('fontSize'))).toEqual([])
	})

	test('table-level fit:shrink cascades to a cell with no explicit fit', async () => {
		const pres = new TsPptx()
		await pres.registerFontMetrics('CellFont', CELL_FONT)
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
