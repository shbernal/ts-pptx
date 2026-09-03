import {
	defineRegressionSuite,
	build,
	listEntries,
	readEntry,
	assert,
	assertEqual,
	captureDiagnostics,
} from '../../helpers.js'

// Every case here pins an ASYMMETRIC or non-default input. The bugs these cover all cancel out
// under the symmetric default margin and the default `x`, which is exactly why the rest of the
// suite stayed green while they shipped.

const SLIDE_XML = 'ppt/slides/slide1.xml'

function slideNames(zip) {
	return listEntries(zip)
		.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

async function rowsPerSlide(zip) {
	const counts = []
	for (const name of slideNames(zip)) counts.push(((await readEntry(zip, name)).match(/<a:tr /g) || []).length)
	return counts
}

/** The table's total grid width in EMU, summed off the emitted `<a:gridCol>`s. */
function gridWidthEmu(xml) {
	return [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].reduce((total, m) => total + Number(m[1]), 0)
}

/** Rows whose text is long enough to page, all identical so every page's budget is comparable. */
function uniformRows(count) {
	const rows = []
	for (let idx = 1; idx <= count; idx++) rows.push([`R${idx}`, 'lorem ipsum dolor sit amet'])
	return rows
}

defineRegressionSuite('Table geometry under asymmetric input', [
	{
		// `resolveSlideMarginsInches` returns TRBL. The usable-width helper read index 1 (right)
		// as the left-edge fallback and subtracted index 3 (left) as the right margin — both one
		// index off, which cancels for a symmetric margin. The sibling site in the definer had
		// already been found and fixed with a comment saying exactly this.
		name: 'a wide LEFT slide margin does not shrink the table as though it were the right one',
		fn: async () => {
			const widthFor = async (margin) => {
				const { zip } = await build((p) => {
					p.defineLayout({ name: 'TEN', width: 10, height: 5.625 })
					p.layout = 'TEN'
					p.defineSlideMaster({ title: 'M', margin })
					p.addSlide({ masterTitle: 'M' }).addTable([['a', 'b']], { x: 1, autoPage: false })
				})
				return gridWidthEmu(await readEntry(zip, SLIDE_XML))
			}
			// The table starts at x=1 either way, so only the RIGHT margin can change its width.
			const wideRight = await widthFor([0.5, 2, 0.5, 0.5])
			const wideLeft = await widthFor([0.5, 0.5, 0.5, 2])
			const symmetric = await widthFor([0.5, 0.5, 0.5, 0.5])
			assertEqual(wideLeft, symmetric, 'a wide left margin leaves a table at x=1 the same width')
			assert(wideRight < symmetric, `a wide right margin narrows it; got ${wideRight} against ${symmetric}`)
		},
	},
	{
		// `slideMargin` was read by the auto-pager and by nothing else, so whether a table
		// respected it depended on whether `autoPage` was on.
		name: 'slideMargin steers an un-paged table too',
		fn: async () => {
			const widthFor = async (slideMargin) => {
				const { zip } = await build((p) => {
					p.defineLayout({ name: 'TEN', width: 10, height: 5.625 })
					p.layout = 'TEN'
					p.addSlide().addTable([['a', 'b']], { autoPage: false, slideMargin })
				})
				return gridWidthEmu(await readEntry(zip, SLIDE_XML))
			}
			const narrow = await widthFor(0.25)
			const wide = await widthFor(2)
			assert(wide < narrow, `a wider slideMargin narrows the table; got ${wide} against ${narrow}`)
		},
	},
	{
		// `headerRow` is inline styling for row 0, baked into the cells at definition time. It was
		// carried onto every continuation page, where the recursive `addTable` re-ran the sugar
		// against THAT page's row 0 — an arbitrary body row — painting it as a header.
		name: 'headerRow styling is not reapplied to a body row on each continuation page',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineLayout({ name: 'TEN', width: 10, height: 5.625 })
				p.layout = 'TEN'
				p.addSlide().addTable(uniformRows(50), {
					x: 0.5,
					y: 0.5,
					w: 9,
					autoPage: true,
					fontSize: 16,
					headerRow: { bold: true, fill: 'FFFF00' },
				})
			})
			const names = slideNames(zip)
			assert(names.length >= 3, `expected several pages; got ${names.length}`)
			for (const [idx, name] of names.entries()) {
				const xml = await readEntry(zip, name)
				const fills = [...xml.matchAll(/<a:srgbClr val="FFFF00"\/>/g)].length
				if (idx === 0) assert(fills > 0, 'the real header row keeps its fill')
				else assertEqual(fills, 0, `page ${idx + 1} paints no body row as a header`)
				const firstRow = xml.match(/<a:tbl>[\s\S]*?<\/a:tr>/)?.[0] ?? ''
				assertEqual(
					/firstRow="1"/.test(xml),
					idx === 0,
					`page ${idx + 1} declares firstRow only where a header actually is; ${firstRow.slice(0, 120)}`
				)
			}
		},
	},
	{
		// A repeated header row read `_lineHeight` off the DEFINER's cells, which never carry it —
		// so it was priced at zero and each continuation page took the header for free and then
		// packed the same body rows the first page fits.
		name: 'a repeated header row costs the page budget it occupies',
		fn: async () => {
			const paged = (repeat) =>
				build((p) => {
					p.defineLayout({ name: 'TEN', width: 10, height: 5.625 })
					p.layout = 'TEN'
					p.addSlide().addTable(uniformRows(60), {
						x: 0.5,
						y: 0.5,
						w: 9,
						autoPage: true,
						fontSize: 16,
						autoPageSlideStartY: 0.5,
						autoPageRepeatHeader: repeat,
						autoPageHeaderRows: 1,
					})
				})
			const plain = await rowsPerSlide((await paged(false)).zip)
			const repeated = await rowsPerSlide((await paged(true)).zip)
			// Both drop the last page, which holds whatever remains and may be short.
			const plainFull = plain.slice(0, -1)
			const repeatedFull = repeated.slice(0, -1)
			assert(plainFull.length >= 2 && repeatedFull.length >= 2, 'expected several full pages either way')
			for (const count of repeatedFull.slice(1)) {
				assert(
					count <= plainFull[0],
					`a page carrying a repeated header fits no more rows than a page without one; ` +
						`got ${JSON.stringify(repeated)} against ${JSON.stringify(plain)}`
				)
			}
		},
	},
	{
		// Three rules for the paged height: the "after the initial slide" block was gated on
		// `> 1`, so it started on the THIRD page, and the explicit-`h` floor reached page one and
		// pages three and up. Page two was the one page that got neither.
		name: 'every page of an explicit-height table gets the same usable height',
		fn: async () => {
			const { zip } = await build((p) => {
				p.layout = 'LAYOUT_16x9'
				p.addSlide().addTable(uniformRows(60), { x: 0.5, y: 0.2, w: 8, h: 4, autoPage: true, fontSize: 18 })
			})
			const counts = await rowsPerSlide(zip)
			const full = counts.slice(0, -1)
			assert(full.length >= 3, `expected several full pages; got ${JSON.stringify(counts)}`)
			assert(
				new Set(full).size === 1,
				`every full page fits the same number of identical rows; got ${JSON.stringify(counts)}`
			)
		},
	},
	{
		// Six sites read a span, in six spellings that agreed only on the values that were
		// already valid: `Number(x) ? … : 1`, `Math.max(1, Number(x) || 1)`, and a bare
		// `x ?? 1` that passed a non-numeric span straight through to a column count. They
		// now read it through `resolveSpan`, the same rule the up-front check applies, so the
		// reading cannot drift apart again if that check is ever moved.
		name: 'a colspan that is not a whole number is reported and read as 1 everywhere',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addTable([[{ text: 'a', options: { colspan: 'x' } }, 'b', 'c']], { x: 1, y: 1, w: 6 })
				})
				return readEntry(zip, SLIDE_XML)
			})
			assertEqual([...xml.matchAll(/<a:gridCol /g)].length, 3, `three grid columns; got: ${xml}`)
			assert(!xml.includes('NaN'), `no NaN may reach the grid; got: ${xml}`)
			assert(codes.includes('table/span-out-of-range'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a cell margin that is not four finite inches is reported and falls back',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addTable([[{ text: 'a', options: { margin: [0.1, Number.NaN, 0.1, 0.1] } }]], {
						x: 1,
						y: 1,
						w: 6,
					})
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('NaN'), `no NaN may reach marL/R/T/B; got: ${xml}`)
			assert(codes.includes('table/invalid-margin'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		// A cell stating `margin: [0, …]` used to fall through to the table's margin in the pager
		// because each side was gated on truthiness.
		name: 'a cell margin of zero is a margin, not an absent one',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a', options: { margin: [0, 0, 0, 0] } }]], {
					x: 1,
					y: 1,
					w: 6,
					margin: [0.5, 0.5, 0.5, 0.5],
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(/marT="0"/.test(xml) && /marB="0"/.test(xml), `the cell keeps its own zero margins; got: ${xml}`)
		},
	},
])
