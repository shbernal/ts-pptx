// `rowH` used to be read in four places that disagreed, and the one that disagreed loudest was
// `pptx.tableLayout()` — the public prediction API. For `rowH: [0, 2]` the writer baked a 2.0in
// first row (`0` is falsy, so it fell through to the even split of `h`) while `tableLayout()`
// reported 0.2004in and flagged it `heightExact: true`: a tenfold error on the number a caller
// places the next shape against, reported as pinned. A negative entry reached the file as
// `<a:tr h="-914400">`, and a stringified one (`rowH: ['1']`, reachable from untyped JS) was
// honoured by the writer and rejected by the layout API.
//
// `docs/measured-text-fit.md` states the invariant: "a layout-time prediction must never
// disagree with what the export then bakes". Column widths honoured it through
// `resolveTableColWidthsEmu`; row heights now honour it through `resolveTableRowHeightEmu`.
//
// So the assertion worth having is three-way, and it is what this file is: for one set of
// `TableProps`, the emitted `<a:tr h>`, `computeTableLayout()`'s per-row height, and the
// export-time measured-fit pass's notion of whether a row is fixed-height all have to agree.
// Checking any one of them alone is what let this drift.
import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import TsPptx, { setDiagnosticHandler } from '../../../dist/node.js'
import { computeTableLayout } from '../../../src/measure/table-fit.ts'
import { applyMeasuredFit } from '../../../src/measure/fit.ts'
import { FontMetricsRegistry } from '../../../src/measure/font-metrics.ts'

const EMU_PER_INCH = 914400
const LAYOUT = { name: 'test', width: 10, height: 5.625, _sizeW: 9144000, _sizeH: 5143500 }

/** Monospace synthetic metrics, so the fit pass is reproducible without a font file. */
const mono = () => ({
	unitsPerEm: 1000,
	advanceWidthPt: (text, sizePt, charSpacingPt = 0) => [...text].length * (0.5 * sizePt + charSpacingPt),
	hasCodepoint: () => true,
})

/** Enough text that a 2in-tall, 2in-wide cell cannot hold it at 18pt. */
const LONG = 'word '.repeat(200)
/** @returns {import('../../../src/types/index.js').TableRow[]} */
const ROWS = () => [[{ text: LONG, options: { fontFace: 'Mono', fontSize: 18, fit: 'shrink' } }], [{ text: 'second' }]]

/** Per-row `<a:tr h>` in inches, straight out of the built package. 0 is an auto-height row. */
async function writerRowHeightsIn(opts) {
	const pres = new TsPptx()
	pres.addSlide().addTable(ROWS(), opts)
	const zip = await JSZip.loadAsync(await pres.write({ outputType: 'nodebuffer' }))
	const xml = await zip.file('ppt/slides/slide1.xml').async('string')
	return [...xml.matchAll(/<a:tr h="(-?\d+)"/g)].map((m) => Number(m[1]) / EMU_PER_INCH)
}

/** Per-row height (inches) and exactness, from the public prediction API's engine. */
function layoutRows(opts) {
	const res = computeTableLayout(ROWS(), opts, LAYOUT, new FontMetricsRegistry())
	return [0, 1].map((row) => {
		const cell = res.cells.find((c) => c.row === row && c.col === 0)
		return { hIn: cell?.hIn ?? 0, exact: cell?.heightExact ?? false }
	})
}

/**
 * Whether the export-time fit pass treated row 0 as fixed-height, observed the only way it is
 * observable: a `fit:'shrink'` cell in an auto-height row is skipped (the row grows instead), so
 * a reduced font size means the pass resolved a real height for that row.
 */
function fitShrankRow0(opts) {
	const pres = new TsPptx()
	pres.addSlide().addTable(ROWS(), opts)
	const registry = new FontMetricsRegistry()
	registry.set('Mono', mono())
	// The pass runs over the internal slide list, which `write()` reaches through `gen/prepare.ts`.
	// Calling it directly is what lets synthetic metrics stand in for a font file.
	const slides = /** @type {any} */ (pres)._slides
	applyMeasuredFit(slides, registry)
	return slides[0]._slideObjects[0].arrTabRows[0][0].options.fontSize < 18
}

describe('rowH is read the same way by the writer, the layout API and the fit pass', () => {
	// `heightIn` is what the three have to agree on. `exact` is the second claim `tableLayout()`
	// makes, and it has to track the writer too: a row the file pins is exact, a row that grows to
	// fit is an estimate.
	const CASES = [
		{ name: 'a zero entry does not pin — the row is sized from `h`', opts: { rowH: [0, 2] }, heights: [2, 2] },
		{ name: 'a positive entry pins its own row', opts: { rowH: [1, 2] }, heights: [1, 2] },
		{ name: 'a negative entry does not pin, and never reaches the file', opts: { rowH: [-1, 2] }, heights: [2, 2] },
		{
			name: 'a numeric string pins, as every reading but one already had it',
			opts: { rowH: ['1', 2] },
			heights: [1, 2],
		},
		{ name: 'a scalar applies to every row', opts: { rowH: 1.5 }, heights: [1.5, 1.5] },
		{ name: 'no rowH with a table `h` splits it evenly', opts: {}, heights: [2, 2] },
	]

	for (const { name, opts, heights } of CASES) {
		test(name, async () => {
			const full = { x: 0.5, y: 0.5, w: 2, h: 4, ...opts }
			const writer = await writerRowHeightsIn(full)
			expect(writer).toEqual(heights)

			const layout = layoutRows(full)
			expect(layout.map((r) => r.hIn)).toEqual(heights)
			expect(layout.map((r) => r.exact)).toEqual([true, true])

			// Every case above pins row 0 to something, so the fit pass has a box to shrink into.
			expect(fitShrankRow0(full)).toBe(true)
		})
	}

	test('with no rowH and no `h`, all three agree the rows are auto-height', async () => {
		const opts = { x: 0.5, y: 0.5, w: 2 }
		// The writer spells auto-height as h="0" — the row grows to fit in PowerPoint.
		expect(await writerRowHeightsIn(opts)).toEqual([0, 0])
		// The layout API cannot leave a hole, so it estimates and says the estimate is not exact.
		const layout = layoutRows(opts)
		expect(layout.every((r) => r.hIn > 0)).toBe(true)
		expect(layout.map((r) => r.exact)).toEqual([false, false])
		// And the fit pass leaves the cell alone rather than shrinking it into a box it does not have.
		expect(fitShrankRow0(opts)).toBe(false)
	})
})

describe('`cy` is the table height to all three, not to one of them', () => {
	// `cy` is the already-resolved EMU height the auto-pager and the fit pass stamp onto a
	// table's options. Only the fit pass read it, so `addTable(rows, { cy })` with no `h` gave a
	// file whose rows are pinned and a `pptx.tableLayout()` that reported every row auto-height —
	// the same drift the one reading of `rowH` closed, arriving through a second option.
	const CY_OPTS = { x: 0.5, y: 0.5, w: 2, cy: 4 * EMU_PER_INCH }

	test('the layout API pins the rows the writer pins', async () => {
		expect(await writerRowHeightsIn(CY_OPTS)).toEqual([2, 2])
		expect(layoutRows(CY_OPTS)).toEqual([
			{ hIn: 2, exact: true },
			{ hIn: 2, exact: true },
		])
	})

	test('and the fit pass shrinks into the same box', () => {
		expect(fitShrankRow0(CY_OPTS)).toBe(true)
	})
})

describe('a colW entry that is not a width is reported', () => {
	// Two columns, so a two-entry `colW` matches the column count and survives
	// `addTableDefinition`'s mismatch check to reach the resolver.
	const TWO_COL = () => [[{ text: 'a' }, { text: 'b' }]]

	/** Codes reported while building a deck whose one table carries `colW`. */
	async function codesForColW(colW) {
		const codes = []
		setDiagnosticHandler((d) => codes.push(d.code))
		try {
			const pres = new TsPptx()
			pres.addSlide().addTable(TWO_COL(), { x: 0.5, y: 0.5, w: 4, h: 4, colW })
			await pres.write({ outputType: 'nodebuffer' })
		} finally {
			setDiagnosticHandler(null)
		}
		return codes
	}

	test('a non-numeric entry warns, like the analogous rowH entry', async () => {
		// The even split is not what a caller writing `NaN` meant, and `rowH` has said so for a
		// while; `colW` fell back silently.
		expect(await codesForColW([2, Number.NaN])).toContain('table/invalid-col-width')
	})

	test('a missing slot is silent — a sparse array distributes that column', async () => {
		const sparse = [2]
		sparse.length = 2
		expect(await codesForColW(sparse)).not.toContain('table/invalid-col-width')
	})
})

describe('rowH entries that are not heights are reported', () => {
	/** Codes reported while building a deck whose one table carries `rowH`. */
	async function codesForRowH(rowH) {
		const codes = []
		setDiagnosticHandler((d) => codes.push(d.code))
		try {
			const pres = new TsPptx()
			pres.addSlide().addTable(ROWS(), { x: 0.5, y: 0.5, w: 2, h: 4, rowH })
			await pres.write({ outputType: 'nodebuffer' })
		} finally {
			setDiagnosticHandler(null)
		}
		return codes
	}

	test('a zero, negative or non-numeric entry warns rather than silently picking a winner', async () => {
		// None of the readings this replaced was what a caller writing `0` meant, so the fallback
		// is stated rather than guessed at.
		for (const rowH of [
			[0, 2],
			[-1, 2],
			['x', 2],
		])
			expect(await codesForRowH(rowH)).toContain('table/invalid-row-height')
	})

	test('a missing array slot is silent — that is how an auto-height row is spelled', async () => {
		// The auto-pager builds per-slide `rowH` arrays with `undefined` holes for exactly this,
		// so warning on them would make every paged table with an auto row noisy.
		expect(await codesForRowH([undefined, 2])).not.toContain('table/invalid-row-height')
	})
})
