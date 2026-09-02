import { defineRegressionSuite, build, readEntry, assert, captureDiagnostics } from '../../helpers.js'

// Acceptance: when a table is sized with `w` (or nothing) but no explicit `colW`,
// the emitted <a:gridCol w=…> must be the table width split evenly in EMU — not the
// raw inches value used as EMU. The historical bug divided inches (`w=9`) and emitted
// `gridCol w="3"` (≈0 EMU), collapsing every auto-width table to a sliver.

const ONE_IN_EMU = 914400

function gridColWidths(xml) {
	return [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
}

defineRegressionSuite('Table column-width distribution', [
	{
		name: '`w` without `colW` splits the EMU width evenly across columns',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, w: 9, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assert(cols.length === 3, `expected 3 gridCols; got ${cols.length}`)
			const expected = Math.round((9 * ONE_IN_EMU) / 3) // 3 inches per column
			cols.forEach((w) => assert(w === expected, `expected gridCol w=${expected} EMU; got ${w}`))
			// Regression guard: the old bug emitted w="3" (raw inches as EMU).
			assert(!cols.includes(3), 'gridCol must not be the raw inches value treated as EMU (w="3")')
		},
	},
	{
		name: 'neither `w` nor `colW` (default full-slide width) still yields inch-scale columns',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C', 'D']], { x: 0.5, y: 0.5, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assert(cols.length === 4, `expected 4 gridCols; got ${cols.length}`)
			// Each column should be a sane fraction of the slide width, not ~0 EMU.
			cols.forEach((w) => assert(w > ONE_IN_EMU, `expected each gridCol > 1in EMU; got ${w}`))
		},
	},
	{
		// A scalar `colW` becomes `w = colW * colCount`, and that product used to be floored to
		// whole inches — discarding up to a full inch across the table. `colW` is documented as
		// inches with no rounding rule; the only rounding a length needs happens in `inch2Emu`.
		name: 'a fractional scalar `colW` is not floored to whole inches',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, colW: 2.4, h: 1 })
			})
			const cols = gridColWidths(await readEntry(zip, 'ppt/slides/slide1.xml'))
			const expected = Math.round(2.4 * ONE_IN_EMU)
			assert(
				cols.length === 3 && cols.every((w) => w === expected),
				`expected three ${expected} EMU columns (2.4in); got ${JSON.stringify(cols)}`
			)
		},
	},
	{
		// `colW: [3]` on a multi-column table means "this width for all". The branch read the
		// width by coercing the whole array (`Number([3])`) rather than reading `colW[0]`, which
		// is only correct for a one-element array by accident of array-to-primitive coercion.
		name: 'a one-element `colW` array spreads its own element, unfloored',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, colW: [1.5], h: 1 })
			})
			const cols = gridColWidths(await readEntry(zip, 'ppt/slides/slide1.xml'))
			const expected = Math.round(1.5 * ONE_IN_EMU)
			assert(
				cols.length === 3 && cols.every((w) => w === expected),
				`expected three ${expected} EMU columns (1.5in); got ${JSON.stringify(cols)}`
			)
		},
	},
	{
		// A non-numeric scalar used to become `w = NaN` and surface far downstream as
		// `coord/non-finite`, whose message describes a missing layout dimension and names
		// nothing the caller wrote.
		name: 'a non-numeric scalar `colW` warns about `colW` and falls back to the default width',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					const s = p.addSlide()
					s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, colW: 'wide', h: 1 })
				})
			)
			assert(
				codes.includes('table/invalid-col-width'),
				`expected table/invalid-col-width; got ${JSON.stringify(codes)}`
			)
			const cols = gridColWidths(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assert(
				cols.length === 3 && cols.every((w) => w > ONE_IN_EMU),
				`expected three sane columns; got ${JSON.stringify(cols)}`
			)
		},
	},
	{
		name: 'explicit `colW` array is still honored per column (inches → EMU)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, colW: [2, 3, 4], h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assert(
				cols.length === 3 && cols[0] === 2 * ONE_IN_EMU && cols[1] === 3 * ONE_IN_EMU && cols[2] === 4 * ONE_IN_EMU,
				`expected [2in,3in,4in] EMU; got ${JSON.stringify(cols)}`
			)
		},
	},
])
