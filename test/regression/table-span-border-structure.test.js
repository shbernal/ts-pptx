import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Verification suite for several historical upstream table reports that the current fork already
// emits correctly. These guard against regressing back into the reported symptoms:
//   #1290 — colspan > 1 produced extra empty cells/columns
//   #1318 — top/left borders missing on non-first rows/columns
//   #1055 — rowspan continuation cells lost borders
//   #1224 — border of the cell adjacent to a rowspan was not applied
// All four reduce to two invariants in the generated <a:tbl>: (1) the grid/cell counts match the
// declared spans, and (2) every cell — including span continuation (hMerge/vMerge) cells — carries
// the full four-sided border when a uniform border is configured.

const SOLID = { type: 'solid', pt: 1, color: 'FF0000' }
const BORDER4 = [SOLID, SOLID, SOLID, SOLID]

async function tableXml(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	return /<a:tbl>[\s\S]*<\/a:tbl>/.exec(xml)[0]
}

function cells(tblXml) {
	return tblXml.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g) || []
}

function hasAllFourBorders(cellXml) {
	return ['lnL', 'lnR', 'lnT', 'lnB'].every((s) => new RegExp(`<a:${s} w="\\d+"`).test(cellXml))
}

defineRegressionSuite('Table span + border structure (verified correct)', [
	{
		name: 'colspan produces correct grid/cell counts, no extra empty columns (#1290)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'A', options: { colspan: 2 } }, { text: 'C' }],
						[{ text: 'd' }, { text: 'e' }, { text: 'f' }],
					],
					{ x: 0, y: 0, w: 9, colW: [3, 3, 3] }
				)
			})
			const tbl = await tableXml(zip)
			assert((tbl.match(/<a:gridCol/g) || []).length === 3, 'expected exactly 3 grid columns')
			const rows = tbl.split('</a:tr>').filter((r) => r.includes('<a:tr'))
			assert(rows.length === 2, `expected 2 rows, got ${rows.length}`)
			rows.forEach((r, i) => {
				const tcCount = (r.match(/<a:tc[ >]/g) || []).length
				assert(tcCount === 3, `row ${i} should have 3 cells (incl. hMerge filler), got ${tcCount}`)
			})
			// Origin cell carries gridSpan=2 and exactly one hMerge filler exists in row 0.
			assert(/gridSpan="2"/.test(rows[0]), 'colspan origin cell should declare gridSpan="2"')
			assert((rows[0].match(/hMerge="1"/g) || []).length === 1, 'row 0 should have exactly one hMerge filler cell')
		},
	},
	{
		name: 'uniform border renders on every cell incl. non-first rows/columns (#1318)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'a' }, { text: 'b' }],
						[{ text: 'c' }, { text: 'd' }],
					],
					{ x: 0, y: 0, w: 6, border: BORDER4 }
				)
			})
			const tbl = await tableXml(zip)
			const tcs = cells(tbl)
			assert(tcs.length === 4, `expected 4 cells, got ${tcs.length}`)
			tcs.forEach((c, i) => assert(hasAllFourBorders(c), `cell ${i} must carry all four borders`))
		},
	},
	{
		name: 'rowspan continuation + adjacent cells keep borders (#1055, #1224)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'A', options: { rowspan: 2 } }, { text: 'B' }], [{ text: 'C' }]], {
					x: 0,
					y: 0,
					w: 6,
					border: BORDER4,
				})
			})
			const tbl = await tableXml(zip)
			const tcs = cells(tbl)
			// origin (rowSpan=2) + B + vMerge continuation + C = 4 cells
			assert(tcs.length === 4, `expected 4 cells, got ${tcs.length}`)
			assert(/rowSpan="2"/.test(tcs[0]), 'origin cell should declare rowSpan="2"')
			assert(/vMerge="1"/.test(tcs[2]), 'covered row should emit a vMerge continuation cell')
			tcs.forEach((c, i) =>
				assert(hasAllFourBorders(c), `cell ${i} (incl. vMerge continuation) must carry all four borders`)
			)
		},
	},
])
