import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Regression: borders (and fill) configured on a colspan/rowspan table cell must
// render across the whole merged region. PowerPoint defines a merged region's
// outer edges on the *covered* cells — the right edge of a colspan lives on the
// rightmost `hMerge` cell, the bottom edge of a rowspan on the lowest `vMerge`
// cell — so emitting an empty `<a:tcPr/>` for those dummy cells drops the
// configured border/fill on those edges.
//
// Reproduces upstream-issue-680.

// Pull the `<a:tcPr>…</a:tcPr>` of the dummy span cell carrying the given merge
// attribute (e.g. `hMerge="1"` / `vMerge="1"`). Covered cells have no text body,
// so the cell is just `<a:tc …><a:tcPr>…</a:tcPr></a:tc>`.
function coveredCellTcPr(xml, mergeAttr) {
	const re = new RegExp(`<a:tc ${mergeAttr}>(.*?)</a:tc>`)
	const m = xml.match(re)
	assert(m, `expected a covered cell with ${mergeAttr}; got: ${xml.slice(0, 600)}`)
	return m[1]
}

defineRegressionSuite('Table merged-cell borders', 'upstream-issue-680', [
	{
		name: 'colspan/rowspan covered cells inherit the origin cell border and fill',
		fn: async () => {
			// 3 columns:
			//   Row 0: [ A (rowspan=2, blue border, blue fill), B (colspan=2, red border, red fill) ]
			//   Row 1: [ A-continues (vMerge), b1, b2 ]
			const blue = makeBorder('0000FF')
			const red = makeBorder('FF0000')

			const { zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[
							{ text: 'A', options: { rowspan: 2, border: blue, fill: { color: 'E0E0FF' } } },
							{ text: 'B', options: { colspan: 2, border: red, fill: { color: 'FFE0E0' } } },
						],
						[{ text: 'b1' }, { text: 'b2' }],
					],
					{ x: 1, y: 1, w: 9 }
				)
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')

			// 1. Pre-fix the covered cells were emitted as empty `<a:tcPr/>`; that must be gone.
			assert(
				!xml.includes('<a:tcPr/>'),
				`expected no empty <a:tcPr/> on covered span cells (upstream-issue-680); got: ${xml.slice(0, 800)}`
			)

			// 2. The colspan's rightmost covered cell (hMerge) must carry the red border so the
			//    merged region's right edge renders, plus the red fill so the region fills uniformly.
			const hMergePr = coveredCellTcPr(xml, 'hMerge="1"')
			assert(hMergePr.includes('<a:lnR '), `expected lnR on the hMerge cell; got tcPr: ${hMergePr}`)
			assert(
				hMergePr.includes('val="FF0000"'),
				`expected the red border color on the hMerge cell; got tcPr: ${hMergePr}`
			)
			assert(
				hMergePr.includes('val="FFE0E0"'),
				`expected the red fill carried onto the hMerge cell; got tcPr: ${hMergePr}`
			)

			// 3. The rowspan's lower covered cell (vMerge) must carry the blue border so the merged
			//    region's bottom edge renders, plus the blue fill.
			const vMergePr = coveredCellTcPr(xml, 'vMerge="1"')
			assert(vMergePr.includes('<a:lnB '), `expected lnB on the vMerge cell; got tcPr: ${vMergePr}`)
			assert(
				vMergePr.includes('val="0000FF"'),
				`expected the blue border color on the vMerge cell; got tcPr: ${vMergePr}`
			)
			assert(
				vMergePr.includes('val="E0E0FF"'),
				`expected the blue fill carried onto the vMerge cell; got tcPr: ${vMergePr}`
			)
		},
	},
])

function makeBorder(color) {
	const side = { type: 'solid', color, pt: 2 }
	return [side, side, side, side]
}
