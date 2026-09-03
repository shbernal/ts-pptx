import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'

// The shape of every bug here: build the same literal into TWO objects and the second behaves
// differently from the first, because the first build wrote into it. `addTableDefinition` takes
// ownership of the TABLE options for exactly this reason and left the per-cell ones aliased.

const SLIDE_XML = 'ppt/slides/slide1.xml'

/** Each `<a:tbl>` on the slide, in document order. */
function tables(xml) {
	return [...xml.matchAll(/<a:tbl>[\s\S]*?<\/a:tbl>/g)].map((m) => m[0])
}

defineRegressionSuite('Table literals the caller still owns', [
	{
		name: 'a rows literal reused across two tables is not styled by the first',
		fn: async () => {
			const rows = [[{ text: 'A' }, { text: 'B' }]]
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(rows, { x: 0.5, y: 0.5, w: 6, color: 'FF0000', bold: true })
				s.addTable(rows, { x: 0.5, y: 3, w: 6, color: '0000FF' })
			})
			const [first, second] = tables(await readEntry(zip, SLIDE_XML))
			assert(first.includes('<a:srgbClr val="FF0000"/>'), `the first table is red; got ${first}`)
			assert(second.includes('<a:srgbClr val="0000FF"/>'), `the second table is blue; got ${second}`)
			assert(!second.includes('<a:srgbClr val="FF0000"/>'), `and carries none of the first's red; got ${second}`)
			assert(!/<a:rPr[^>]*\bb="1"/.test(second), `nor its bold; got ${second}`)
		},
	},
	{
		name: 'the caller cell objects come back as written',
		fn: async () => {
			const cell = { text: 'A', options: { colspan: 1 } }
			const rows = [[cell, { text: 'B' }]]
			const before = JSON.stringify(rows)
			await build((p) => {
				p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 6, color: 'FF0000', fontSize: 20, margin: 0.2 })
			})
			assertEqual(JSON.stringify(rows), before, 'the caller-owned rows after addTable')
		},
	},
	{
		name: 'a cell literal shared between two tables keeps its own options only',
		fn: async () => {
			const shared = { text: 'shared', options: { bold: true } }
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([[shared]], { x: 0.5, y: 0.5, w: 4, fontSize: 30 })
				s.addTable([[shared]], { x: 0.5, y: 3, w: 4 })
			})
			const [first, second] = tables(await readEntry(zip, SLIDE_XML))
			assert(/sz="3000"/.test(first), `the first table sizes at 30pt; got ${first}`)
			assert(!/sz="3000"/.test(second), `the second inherits nothing from it; got ${second}`)
			assertEqual(JSON.stringify(shared.options), JSON.stringify({ bold: true }), 'the shared cell options')
		},
	},
	{
		// The cell path has always been explicit that a null side is OMITTED (it keeps inheriting
		// from the built-in style) rather than erased. The table path filled the hole with an
		// explicit `{ type: 'none' }`, so one sparse tuple meant two different things depending on
		// where it was written — and, because the same array object reached every cell by
		// reference, the filling propagated into cells that had already captured it.
		name: 'a hole in a table-level border tuple is left alone, as it is on a cell',
		fn: async () => {
			const sparse = [
				{ type: 'solid', color: '333333', width: 1 },
				null,
				{ type: 'solid', color: '333333', width: 1 },
				null,
			]
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B']], { x: 0.5, y: 0.5, w: 6, tableStyle: 'MEDIUM_STYLE_2_ACCENT_1', border: sparse })
				s.addTable([[{ text: 'A', options: { border: sparse } }, 'B']], {
					x: 0.5,
					y: 3,
					w: 6,
					tableStyle: 'MEDIUM_STYLE_2_ACCENT_1',
				})
			})
			const [tableLevel, cellLevel] = tables(await readEntry(zip, SLIDE_XML))
			for (const [label, xml] of [
				['table-level', tableLevel],
				['cell-level', cellLevel],
			]) {
				assert(xml.includes('<a:lnT'), `${label}: the stated top edge is drawn`)
				assert(xml.includes('<a:lnB'), `${label}: the stated bottom edge is drawn`)
				assert(!xml.includes('<a:lnL'), `${label}: the hole leaves the left edge to the style; got ${xml}`)
				assert(!xml.includes('<a:lnR'), `${label}: the hole leaves the right edge to the style; got ${xml}`)
			}
		},
	},
	{
		name: 'a border tuple literal reused across two tables is not completed in place',
		fn: async () => {
			const sparse = [{ type: 'solid' }, null, null, null]
			const before = JSON.stringify(sparse)
			await build((p) => {
				const s = p.addSlide()
				s.addTable([['A']], { x: 0.5, y: 0.5, w: 4, border: sparse })
				s.addTable([['B']], { x: 0.5, y: 3, w: 4, border: sparse })
			})
			assertEqual(JSON.stringify(sparse), before, 'the caller-owned border tuple after two tables')
		},
	},
])
