// Layout-time table-cell geometry API (upstream-issue-1169). Two layers:
//  A) the core computeTableLayout(rows, opts, layout, registry) against src, with
//     SYNTHETIC metrics so the suite is reproducible and needs no font files;
//  B) the pptx.tableLayout() instance method through dist (heuristic path).
// Width geometry is exact (shared resolveTableColWidthsEmu); auto-height rows are
// conservative estimates flagged heightExact:false, exact when rowH/table h pins them.
import { describe, test, expect } from 'vitest'
import { computeTableLayout } from '../../src/measure-fit.ts'
import { FontMetricsRegistry } from '../../src/font-metrics.ts'
import PptxGenJS from '../../dist/node.js'

// Monospace synthetic metrics: every code point advances `emPerChar` ems.
const mono = (emPerChar = 0.5) => ({
	unitsPerEm: 1000,
	advanceWidthPt(text, sizePt, charSpacingPt = 0) {
		const n = [...text].length
		return n * emPerChar * sizePt + n * charSpacingPt
	},
})
const regWith = (face = 'Mono') => {
	const r = new FontMetricsRegistry()
	r.set(face, mono())
	return r
}
// Minimal layout — bare-number (inch) coords ignore the axis length, so only used
// to satisfy the signature. width/height in inches.
const LAYOUT = { name: 'test', width: 10, height: 5.625, _sizeW: 9144000, _sizeH: 5143500 }
const emptyReg = new FontMetricsRegistry()
const cellAt = (res, row, col) => res.cells.find((c) => c.row === row && c.col === col)

describe('computeTableLayout core — width geometry (exact)', () => {
	test('even column distribution: x positions cumulative, widths sum to w', () => {
		const rows = [
			[{ text: 'a' }, { text: 'b' }, { text: 'c' }],
			[{ text: 'd' }, { text: 'e' }, { text: 'f' }],
		]
		const res = computeTableLayout(rows, { x: 1, y: 1, w: 9, rowH: [0.5, 0.5] }, LAYOUT, emptyReg)
		expect(res.cells).toHaveLength(6)
		expect(res.widthIn).toBeCloseTo(9, 6)
		expect(res.heightIn).toBeCloseTo(1, 6)
		expect(res.heightExact).toBe(true)
		// Row 0: each column is 3in wide, x offset from table x=1.
		expect(cellAt(res, 0, 0)).toMatchObject({ xIn: 1, wIn: 3, yIn: 1, hIn: 0.5, heightExact: true })
		expect(cellAt(res, 0, 1).xIn).toBeCloseTo(4, 6)
		expect(cellAt(res, 0, 2).xIn).toBeCloseTo(7, 6)
		// Row 1 sits one row height down.
		expect(cellAt(res, 1, 0).yIn).toBeCloseTo(1.5, 6)
	})

	test('explicit colW array drives per-column x/width; scalar rowH is exact', () => {
		const rows = [[{ text: 'a' }, { text: 'b' }, { text: 'c' }]]
		const res = computeTableLayout(rows, { x: 0, y: 0, colW: [2, 3, 4], rowH: 0.4 }, LAYOUT, emptyReg)
		expect(cellAt(res, 0, 0)).toMatchObject({ xIn: 0, wIn: 2 })
		expect(cellAt(res, 0, 1)).toMatchObject({ xIn: 2, wIn: 3 })
		expect(cellAt(res, 0, 2)).toMatchObject({ xIn: 5, wIn: 4 })
		expect(res.widthIn).toBeCloseTo(9, 6)
		expect(res.heightIn).toBeCloseTo(0.4, 6)
		expect(res.heightExact).toBe(true)
	})
})

describe('computeTableLayout core — colspan / rowspan', () => {
	test('colspan: merged cell spans its columns; covered slot not emitted', () => {
		const rows = [
			[{ text: 'A', options: { colspan: 2 } }, { text: 'B' }],
			[{ text: 'c' }, { text: 'd' }, { text: 'e' }],
		]
		const res = computeTableLayout(rows, { x: 0, y: 0, colW: [2, 3, 4], rowH: [1, 1] }, LAYOUT, emptyReg)
		// A occupies cols 0-1 (2+3=5in wide); B starts at col 2.
		expect(cellAt(res, 0, 0)).toMatchObject({ col: 0, colSpan: 2, xIn: 0, wIn: 5 })
		expect(cellAt(res, 0, 2)).toMatchObject({ col: 2, colSpan: 1, xIn: 5, wIn: 4 })
		// Row 0 has exactly two emitted cells (col 1 is covered by A's colspan).
		expect(res.cells.filter((c) => c.row === 0)).toHaveLength(2)
		expect(cellAt(res, 0, 1)).toBeUndefined()
	})

	test('rowspan: merged cell spans rows; cell below lands in the next free column', () => {
		const rows = [[{ text: 'A', options: { rowspan: 2 } }, { text: 'B' }], [{ text: 'c' }]]
		const res = computeTableLayout(rows, { x: 0, y: 0, colW: [3, 3], rowH: [1, 1] }, LAYOUT, emptyReg)
		// A spans both rows → height 2in; B is one row.
		expect(cellAt(res, 0, 0)).toMatchObject({ row: 0, col: 0, rowSpan: 2, yIn: 0, hIn: 2 })
		expect(cellAt(res, 0, 1)).toMatchObject({ rowSpan: 1, hIn: 1 })
		// The row-1 cell skips col 0 (covered by A's rowspan) and lands in col 1.
		expect(cellAt(res, 1, 1)).toMatchObject({ row: 1, col: 1, yIn: 1, hIn: 1 })
		expect(cellAt(res, 1, 0)).toBeUndefined()
	})
})

describe('computeTableLayout core — auto-height estimation', () => {
	test('auto row (no rowH/h) is estimated and flagged heightExact:false', () => {
		const rows = [[{ text: 'aaaaaaaaaa bbbbbbbbbb cccccccccc dddddddddd eeeeeeeeee' }]]
		const res = computeTableLayout(
			rows,
			{ x: 0, y: 0, w: 2, colW: [2], fontFace: 'Mono', fontSize: 18 },
			LAYOUT,
			regWith()
		)
		const c = cellAt(res, 0, 0)
		expect(c.heightExact).toBe(false)
		expect(res.heightExact).toBe(false)
		// Wraps to multiple lines at 2in / 18pt mono → taller than a single line.
		const oneLine = (18 * 1.67) / 100 // gen-tables line-height heuristic, inches
		expect(c.hIn).toBeGreaterThan(oneLine)
	})

	test('explicit rowH makes an otherwise-auto table exact even with no metrics', () => {
		const rows = [[{ text: 'hello world' }]]
		const res = computeTableLayout(rows, { x: 0, y: 0, w: 4, rowH: 0.75 }, LAYOUT, emptyReg)
		expect(res.heightExact).toBe(true)
		expect(cellAt(res, 0, 0).hIn).toBeCloseTo(0.75, 6)
	})

	test('auto row with an unmeasurable font still yields a non-zero one-line floor', () => {
		const rows = [[{ text: 'hello world' }]]
		const res = computeTableLayout(rows, { x: 0, y: 0, w: 4 }, LAYOUT, emptyReg)
		const c = cellAt(res, 0, 0)
		expect(c.heightExact).toBe(false)
		expect(c.hIn).toBeGreaterThan(0)
	})

	test('empty rows return empty geometry', () => {
		expect(computeTableLayout([], { x: 0, y: 0, w: 4 }, LAYOUT, emptyReg)).toEqual({
			cells: [],
			widthIn: 0,
			heightIn: 0,
			heightExact: true,
		})
	})
})

describe('pptx.tableLayout() instance method (through dist)', () => {
	test('returns per-cell geometry using the default layout, no metrics required', () => {
		const pptx = new PptxGenJS()
		const rows = [[{ text: 'one' }, { text: 'two' }, { text: 'three' }]]
		const res = pptx.tableLayout(rows, { x: 1, y: 1, w: 8, colW: [2, 3, 3], rowH: 0.5 })
		expect(res.cells).toHaveLength(3)
		expect(cellAt(res, 0, 0)).toMatchObject({ xIn: 1, wIn: 2 })
		expect(cellAt(res, 0, 1)).toMatchObject({ xIn: 3, wIn: 3 })
		expect(cellAt(res, 0, 2)).toMatchObject({ xIn: 6, wIn: 3 })
		expect(res.heightExact).toBe(true)
	})
})
