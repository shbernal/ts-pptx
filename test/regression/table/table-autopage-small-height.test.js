import { setDiagnosticHandler, defineRegressionSuite, build, listEntries, readEntry, assert } from '../../helpers.js'

// Regression: an autoPage table whose height (`h`) is too small to fit even a single line of text
// must NOT emit a degenerate output (an empty `rows:[]` overflow page that made the recursive
// addTable throw "addTable: Array expected", or one row per slide forever). The paginator ignores
// the unusable height (falling back to the slide height) and warns.
//
// NOTE: `h` is the table's *height* (an extent), not a bottom coordinate — `y` does not eat into
// it (see table-autopage-mid-slide.test.js). So a small-but-
// usable `h` like 0.7" paginates normally regardless of `y`; only an `h` smaller than one line of
// the base font is genuinely unusable.

function rows(n) {
	return Array.from({ length: n }, (_, i) => [{ text: `Row ${i} col A` }, { text: `Row ${i} col B` }])
}

function slideCount(zip) {
	return listEntries(zip).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length
}

/** How many `<a:tr>` each slide carries, in slide order. */
async function rowsPerSlide(zip) {
	const names = listEntries(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))
	const counts = []
	for (const name of names) counts.push(((await readEntry(zip, name)).match(/<a:tr[ />]/g) || []).length)
	return counts
}

/** A long table under the unusable-`h` fallback, so every page's height comes from it. */
async function pagedRowCounts(opts) {
	const many = Array.from({ length: 120 }, (_, i) => [{ text: `Row ${i} col A` }, { text: `Row ${i} col B` }])
	const { zip } = await build((p) => {
		p.addSlide().addTable(many, { x: 0.5, w: 9, h: 0.1, colW: [4.5, 4.5], autoPage: true, fontSize: 12, ...opts })
	})
	return rowsPerSlide(zip)
}

defineRegressionSuite('Table autoPage tiny-height guard', [
	{
		name: 'h smaller than one line of text does not crash and emits no empty page (warns instead)',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let zip
			try {
				;({ zip } = await build((p) => {
					// h(0.1") cannot fit even one 12pt line (~0.2") → unusable usable height.
					p.addSlide().addTable(rows(12), {
						x: 0.5,
						y: 1.2,
						w: 9,
						h: 0.1,
						colW: [4.5, 4.5],
						margin: 0,
						slideMargin: 0,
						autoPage: true,
						fontSize: 12,
					})
				}))
			} finally {
				setDiagnosticHandler(null)
			}
			assert(slideCount(zip) >= 1, 'expected at least one slide, never zero or a crash')
			assert(
				warnings.some((w) => w.includes('leaves no room to paginate')),
				`expected a warning about unusable table height; got: ${JSON.stringify(warnings)}`
			)
		},
	},
	{
		name: 'a usable explicit h still paginates normally (no warning)',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let zip
			try {
				;({ zip } = await build((p) => {
					// Plenty of usable height: should paginate by content, not trip the guard.
					p.addSlide().addTable(rows(60), {
						x: 0.5,
						y: 0.5,
						w: 9,
						h: 6,
						colW: [4.5, 4.5],
						margin: 0,
						slideMargin: 0,
						autoPage: true,
						fontSize: 14,
					})
				}))
			} finally {
				setDiagnosticHandler(null)
			}
			assert(slideCount(zip) >= 2, 'a 60-row table should overflow to multiple slides')
			assert(
				!warnings.some((w) => w.includes('leaves no room to paginate')),
				'usable height must not trigger the guard warning'
			)
		},
	},
	{
		name: 'a `y` past the bottom of the slide falls back to the whole slide height',
		fn: async () => {
			// The guard's fallback is "slide height minus the start-Y minus the bottom margin".
			// Put `y` below the slide bottom (the slide is 5.625" tall) and that subtraction goes
			// negative — a usable height of -0.4" would page one row per slide forever. The
			// fallback's own fallback is the full slide height.
			//
			// Enough rows to overflow, so the guard is also exercised on a *subsequent* page,
			// where the start-Y comes from the top margin rather than `y`.
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let zip
			try {
				;({ zip } = await build((p) => {
					p.addSlide().addTable(rows(40), {
						x: 0.5,
						y: 6,
						w: 9,
						h: 0.1,
						colW: [4.5, 4.5],
						margin: 0,
						slideMargin: 0,
						autoPage: true,
						fontSize: 12,
					})
				}))
			} finally {
				setDiagnosticHandler(null)
			}
			// A full slide of usable height fits ~28 rows of 12pt, so 40 rows page but do not
			// degenerate to one-row-per-slide.
			const count = slideCount(zip)
			assert(count >= 2 && count < 40, `expected a few pages, not one per row; got ${count}`)
			assert(
				warnings.some((w) => w.includes('leaves no room to paginate')),
				`expected the unusable-height warning; got: ${JSON.stringify(warnings)}`
			)
		},
	},
	{
		name: 'cell margins taller than the usable height still page without an empty first slide',
		fn: async () => {
			// Row height starts at the row's top+bottom cell margins, before a single line is
			// placed. Margins totalling 1.8" against a 1.5" table height mean the very first
			// row overflows with *nothing* accumulated yet — the page-break path then has an
			// empty row buffer and an empty page buffer, and must push neither. An empty
			// `rows:[]` page is what makes the recursive addTable throw "Array expected".
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows(3), {
					x: 0.5,
					y: 0.5,
					w: 9,
					h: 1.5,
					colW: [4.5, 4.5],
					margin: [0.9, 0, 0.9, 0],
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})
			const count = slideCount(zip)
			assert(count >= 1, 'expected at least one slide, never zero or a crash')
			// Every page that exists must carry a row: no degenerate empty page.
			for (const name of listEntries(zip).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
				const xml = await readEntry(zip, name)
				assert(xml.includes('<a:tr '), `${name} has a table with no rows`)
			}
		},
	},
	// The fallback above computes its own start-Y, and it used to compute it differently from the
	// main path in the same function -- the shape that function's own comment records having
	// already fixed once. Two of the three cases below are what the difference cost.
	{
		name: 'autoPageSlideStartY: 0 is honoured by the fallback, not read as unset',
		fn: async () => {
			const stated = await pagedRowCounts({ y: 1.2, autoPageSlideStartY: 0 })
			const unset = await pagedRowCounts({ y: 1.2 })
			// Starting a continuation page at the very top of the slide buys height, so it must
			// fit MORE rows than letting it start at the top margin. The fallback spelled this
			// `autoPageSlideStartY || topMargin`, and `0` is falsy, so the two were identical.
			assert(
				stated[1] > unset[1],
				`a stated start-Y of 0 must give a taller continuation page; got ${stated} against ${unset}`
			)
		},
	},
	{
		name: 'a `y` above the top margin is not lost when the fallback pages',
		fn: async () => {
			// RULE: paging must not push a table DOWN past a `y` that is already above the margin.
			// The main path says so with `Math.min(y, topMargin)`; the fallback had no such clause,
			// so page one started at `y` and every page after it at the margin -- identical rows
			// that disagreed about how many of them fit.
			const counts = await pagedRowCounts({ y: 0.1 })
			const [first, ...rest] = counts
			for (const [idx, count] of rest.slice(0, -1).entries())
				assert(
					count === first,
					`continuation page ${idx + 2} holds ${count} rows against the first page's ${first}: ${counts}`
				)
		},
	},
	{
		name: 'a `y` below the top margin still starts continuations at the margin',
		fn: async () => {
			// The other side of that `Math.min`: `y` below the margin does NOT pull continuation
			// pages down with it, so they are taller than the first page rather than equal to it.
			const counts = await pagedRowCounts({ y: 1.2 })
			assert(counts[1] > counts[0], `continuation pages must reclaim the space below y; got ${counts}`)
		},
	},
])
