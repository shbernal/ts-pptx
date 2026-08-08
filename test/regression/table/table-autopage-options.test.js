import { defineRegressionSuite, build, readEntry, listEntries, assert, assertEqual } from '../../helpers.js'

// Exercises the option surface of the auto-paging engine (getSlidesForTableRows /
// parseTextToLines in src/gen/table/autopage.ts) through the public `addTable({autoPage:true})`
// path: uniform vs. array `colW`, `colW` without `w`, `colspan`, per-cell/table
// `margin`, `slideMargin`, `autoPageRepeatHeader`, per-cell `fontSize`, and degenerate
// cell text (empty / numeric / whitespace-only). These paths were unreached by the
// existing table suite, which clusters on rowspan/tiny-height/mid-slide geometry.
//
// margin:0 / slideMargin:0 and fontSize:12 (~0.2 in per line) are used where the
// assertion depends on deterministic row heights.

function slideFiles(zip) {
	return listEntries(zip)
		.filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

function gridColCount(xml) {
	return (xml.match(/<a:gridCol\b/g) || []).length
}

// N two-column body rows with distinctive text.
function bodyRows(n) {
	return Array.from({ length: n }, (_, i) => [{ text: `Row ${i} A` }, { text: `Row ${i} B` }])
}

/**
 * Run `fn` with `console.log` captured, returning the lines it emitted. Restoring in a
 * `finally` matters: a throwing build must not leave the rest of the suite stubbed.
 * (Same shape as the `console.warn` capture in connector-shape.test.js.)
 */
async function captureLog(fn) {
	const orig = console.log
	const lines = []
	// `String(a)` rather than JSON: the final dump logs whole TableRowSlide objects, and
	// only the string lines are asserted on below.
	console.log = (...args) => lines.push(args.map((a) => (typeof a === 'string' ? a : String(a))).join(' '))
	try {
		await fn()
	} finally {
		console.log = orig
	}
	return lines
}

/** First captured line matching `re`, or '' — keeps the assertions readable. */
function lineMatching(lines, re) {
	return lines.find((l) => re.test(l)) || ''
}

defineRegressionSuite('Table autoPage option surface', [
	{
		name: 'uniform numeric colW is applied to every column and paginates',
		fn: async () => {
			const { zip } = await build((p) => {
				// colW as a single number → the engine fans it out to one width per column.
				p.addSlide().addTable(bodyRows(40), {
					x: 0.5,
					y: 0.5,
					h: 5,
					colW: 1.5,
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})
			const files = slideFiles(zip)
			assert(files.length >= 2, `expected overflow to multiple slides; got ${files.length}`)
			for (const name of files) {
				const xml = await readEntry(zip, name)
				assertEqual(gridColCount(xml), 2, `every page keeps 2 columns (${name})`)
				// 1.5 in → 1371600 EMU per column.
				assert(xml.includes('<a:gridCol w="1371600"'), `expected 1.5in (1371600 EMU) columns on ${name}`)
			}
		},
	},
	{
		name: 'array colW with no `w` derives table width from the columns',
		fn: async () => {
			const { zip } = await build((p) => {
				// No `w`: total width must come from summing colW.
				p.addSlide().addTable(bodyRows(30), {
					x: 0.5,
					y: 0.5,
					h: 5,
					colW: [2, 3],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})
			const files = slideFiles(zip)
			assert(files.length >= 2, `expected overflow to multiple slides; got ${files.length}`)
			const xml = await readEntry(zip, files[0])
			assert(xml.includes('<a:gridCol w="1828800"'), 'expected a 2in (1828800 EMU) column')
			assert(xml.includes('<a:gridCol w="2743200"'), 'expected a 3in (2743200 EMU) column')
		},
	},
	{
		name: 'a colspan in the header row sets the column count for every page',
		fn: async () => {
			const rows = [
				[{ text: 'Wide Header', options: { colspan: 2 } }],
				...Array.from({ length: 30 }, (_, i) => [{ text: `L${i}` }, { text: `R${i}` }]),
			]
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.5,
					y: 0.5,
					h: 5,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})
			const files = slideFiles(zip)
			assert(files.length >= 2, `expected overflow to multiple slides; got ${files.length}`)
			for (const name of files) {
				const xml = await readEntry(zip, name)
				// numCols is derived from the header colspan (2), so continuation pages keep 2 columns.
				assertEqual(gridColCount(xml), 2, `colspan header must set 2 columns on ${name}`)
			}
		},
	},
	{
		name: 'cell and table margins consume vertical space (fewer rows per page)',
		fn: async () => {
			async function pageCount(useMargins) {
				const { zip } = await build((p) => {
					const rows = Array.from({ length: 24 }, (_, i) => [
						// One cell carries its own margin; the paginator takes the max of cell vs table margin.
						{ text: `Row ${i} A`, options: useMargins ? { margin: [0.15, 0.05, 0.15, 0.05] } : {} },
						{ text: `Row ${i} B` },
					])
					p.addSlide().addTable(rows, {
						x: 0.5,
						y: 0.5,
						h: 4,
						colW: [4.5, 4.5],
						margin: useMargins ? [0.1, 0.05, 0.1, 0.05] : 0,
						slideMargin: 0,
						autoPage: true,
						fontSize: 12,
					})
				})
				return slideFiles(zip).length
			}
			const withMargins = await pageCount(true)
			const without = await pageCount(false)
			assert(withMargins >= 2, `margined table should still paginate; got ${withMargins}`)
			// Per-row top+bottom margins eat vertical space, so fewer rows fit per page → at least as many pages.
			assert(
				withMargins >= without,
				`margins must not increase rows-per-page: withMargins=${withMargins} without=${without}`
			)
		},
	},
	{
		name: 'slideMargin as a 4-tuple bounds the usable area (no explicit h)',
		fn: async () => {
			const { zip } = await build((p) => {
				// No `h`: usable height comes from the slide height minus the slideMargin tuple.
				p.addSlide().addTable(bodyRows(60), {
					x: 0.5,
					w: 9,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: [0.5, 0.5, 0.5, 0.5],
					autoPage: true,
					fontSize: 14,
				})
			})
			assert(slideFiles(zip).length >= 2, 'a 60-row table with margins should overflow')
		},
	},
	{
		name: 'autoPageRepeatHeader repeats the header row on every continuation slide',
		fn: async () => {
			const rows = [
				[
					{ text: 'HEADER-A', options: { bold: true } },
					{ text: 'HEADER-B', options: { bold: true } },
				],
				...bodyRows(40),
			]
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.5,
					y: 0.5,
					h: 4,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					autoPageRepeatHeader: true,
					fontSize: 12,
				})
			})
			const files = slideFiles(zip)
			assert(files.length >= 2, `expected overflow to multiple slides; got ${files.length}`)
			// The header text must reappear on the second page, not only the first.
			const page2 = await readEntry(zip, files[1])
			assert(page2.includes('HEADER-A'), `expected repeated header on ${files[1]}; got: ${page2.slice(0, 400)}`)
		},
	},
	{
		name: 'degenerate cell text (empty / numeric / whitespace) does not crash autoPage',
		fn: async () => {
			const rows = [
				[{ text: '' }, { text: 2024 }],
				[{ text: '   ' }, { text: 'ok', options: { fontSize: 18 } }],
				...bodyRows(30),
			]
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.5,
					y: 0.5,
					h: 5,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					autoPageCharWeight: 0.2,
					fontSize: 12,
				})
			})
			const files = slideFiles(zip)
			assert(files.length >= 1, 'expected at least one slide')
			const page1 = await readEntry(zip, files[0])
			// A nonzero numeric cell renders its digits; the per-cell fontSize cell renders its text.
			assert(page1.includes('2024'), 'expected the numeric cell text "2024" to render')
			assert(page1.includes('>ok<'), 'expected the per-cell fontSize cell text to render')
		},
	},
	{
		name: 'no `y` and no `h` paginates using the full slide height',
		fn: async () => {
			const { zip } = await build((p) => {
				// Neither y nor h given → first-page start falls back to the top margin and the
				// usable height is the slide height between margins.
				p.addSlide().addTable(bodyRows(80), {
					x: 0.5,
					w: 9,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})
			assert(slideFiles(zip).length >= 2, 'an 80-row full-height table should overflow')
		},
	},
	{
		name: '`autoPageLineWeight` inflates the estimated line height and pages sooner',
		fn: async () => {
			// The estimator's line height is `fontSize * (LINEH_MODIFIER + autoPageLineWeight)`,
			// the caller's escape hatch when a font runs taller than the built-in ratio. The
			// assertion is a comparison against the same deck without the option, so it pins the
			// option's *effect* and leaves the ratio itself free to move.
			async function pageCount(extra) {
				const { zip } = await build((p) => {
					p.addSlide().addTable(bodyRows(40), {
						x: 0.5,
						y: 0.5,
						w: 9,
						h: 2,
						colW: [4.5, 4.5],
						margin: 0,
						slideMargin: 0,
						autoPage: true,
						fontSize: 12,
						...extra,
					})
				})
				return slideFiles(zip).length
			}

			const plain = await pageCount({})
			const weighted = await pageCount({ autoPageLineWeight: 0.5 })
			assert(plain > 1, `expected the baseline deck to page at all; got ${plain}`)
			assert(weighted > plain, `a positive line weight should need more pages; got ${weighted} vs ${plain}`)
		},
	},
	{
		name: "a master's scalar `margin` applies to all four sides of the paging area",
		fn: async () => {
			// A master margin outranks `slideMargin`, and may be a single number rather than the
			// [T,R,B,L] array. Only the bottom margin narrows the paging area, so comparing two
			// masters is what proves the scalar was fanned out rather than merely accepted.
			// No `h`: with an explicit height the paging area is clamped to it and the margin
			// would not show.
			async function pageCount(margin) {
				const { zip } = await build((p) => {
					p.defineSlideMaster({ title: `AP_MARGIN_${margin}`, margin })
					p.addSlide({ masterTitle: `AP_MARGIN_${margin}` }).addTable(bodyRows(60), {
						x: 0.5,
						y: 0.5,
						w: 9,
						colW: [4.5, 4.5],
						margin: 0,
						autoPage: true,
						fontSize: 12,
					})
				})
				return slideFiles(zip).length
			}

			const narrow = await pageCount(0.25)
			const wide = await pageCount(2)
			assert(wide > narrow, `a 2" master margin should leave room for fewer rows; got ${wide} vs ${narrow}`)
		},
	},
	{
		name: 'a non-numeric `slideMargin` falls back to the default margins rather than NaN geometry',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable(bodyRows(30), {
					x: 0.5,
					y: 0.5,
					w: 9,
					colW: [4.5, 4.5],
					margin: 0,
					// A string is the untyped-caller shape the engine's own `isNaN` check absorbs.
					slideMargin: 'nope',
					autoPage: true,
					fontSize: 12,
				})
			})
			const files = slideFiles(zip)
			assert(files.length >= 1, 'expected the table to still be emitted')
			for (const name of files) {
				const xml = await readEntry(zip, name)
				assert(!xml.includes('NaN'), `${name} contains NaN geometry`)
			}
		},
	},
	{
		name: 'a row with no cells is dropped instead of emitting a cell-less <a:tr>',
		fn: async () => {
			// An empty row reaches the pager with no cells to walk, so it produces no line and
			// no row buffer to flush. A row element with fewer cells than the grid has columns
			// is exactly the malformation PowerPoint offers to "repair", so the row must be
			// dropped, not emitted empty.
			const rows = [[{ text: 'A0' }, { text: 'B0' }], [], [{ text: 'A2' }, { text: 'B2' }]]
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.5,
					y: 0.5,
					w: 9,
					h: 3,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertEqual((xml.match(/<a:tr\b/g) || []).length, 2, 'expected only the two populated rows')
			assertEqual(gridColCount(xml), 2, 'the grid should still describe both columns')
			assert(xml.includes('>A0<') && xml.includes('>A2<'), 'both populated rows must survive')
		},
	},
	// --- `verbose` -----------------------------------------------------------------
	// `verbose` is a documented (dev-only) `TableProps` flag, and its trace does real
	// arithmetic — `.toFixed()` on props that may legitimately be percentage *strings*.
	// The two cases below pin that the dump survives both input shapes and that what it
	// reports agrees with what was emitted. They pin the trace's *shape*, not the emitted
	// OOXML: nothing here asserts a formatted number's value, so the layout constants stay
	// free to move.
	{
		name: '`verbose` traces the auto-paging run for numeric props and its slide count matches',
		fn: async () => {
			let files = []
			const lines = await captureLog(async () => {
				const { zip } = await build((p) => {
					// `colW` as an array with no `w` is the one shape that reaches the width
					// calc's own trace line — with `w` set, that step is skipped entirely.
					p.addSlide().addTable(bodyRows(40), {
						x: 0.5,
						y: 0.5,
						h: 5,
						colW: [4.5, 4.5],
						margin: 0.05,
						slideMargin: 0.5,
						autoPageSlideStartY: 0.6,
						autoPageCharWeight: 0.2,
						autoPage: true,
						fontSize: 12,
						verbose: true,
					})
				})
				files = slideFiles(zip)
			})

			assert(lines.includes('[[VERBOSE MODE]]'), 'expected the trace header; got: ' + lines.slice(0, 3).join(' | '))
			// A numeric prop goes through the `.toFixed(1)` arm. Only "it printed a number"
			// is asserted — the value is scaled as though `x` were already EMU, which it is
			// not on this fork (`define/table.ts` keeps raw inches through to emission), so
			// pinning the digits would cement a stale unit into the suite.
			assert(
				/\| tableProps\.x .*= [\d.]+$/.test(lineMatching(lines, /\| tableProps\.x /)),
				'expected numeric `x` to print through the number arm; got: ' + lineMatching(lines, /\| tableProps\.x /)
			)
			assertEqual(
				lineMatching(lines, /\| numCols /).replace(/.*= /, ''),
				'2',
				'the trace should report the column count it derived'
			)
			assert(files.length >= 2, `expected overflow to multiple slides; got ${files.length}`)
			assert(
				lines.some((l) => l.includes('NEW SLIDE CREATED')),
				'a paginating table should trace its page breaks'
			)
			assert(
				lines.some((l) => /ROW \[0\]: START/.test(l)),
				'expected the per-row trace'
			)
			// The dump is the engine's own account of what it produced, so it is worth
			// checking against the package rather than merely asserting it printed.
			assertEqual(
				lineMatching(lines, /FINAL: tableRowSlides\.length/).replace(/.*= /, ''),
				String(files.length),
				'the traced page count should match the slides actually emitted'
			)
			// The cell-wrapping trace ([1/4]..[4/4] inside `parseTextToLines`) never appears:
			// its only call site passes `verbose: false` outright. Asserted so those arms are
			// demonstrably unreachable rather than merely untested — if the flag is ever
			// wired through, this is the line that says so.
			assert(
				!lines.some((l) => l.startsWith('[1/4]') || l.startsWith('[4/4]')),
				'the cell-wrapping trace is not wired to `verbose`; expected no [n/4] lines'
			)
		},
	},
	{
		name: '`verbose` prints percentage-string props verbatim instead of NaN',
		fn: async () => {
			const lines = await captureLog(async () => {
				// Percentage strings are a supported `Coord`. The trace guards each one with a
				// `typeof === 'number'` check; without it, `('5%' / 914400).toFixed(1)` would
				// put NaN into a diagnostic meant to explain a layout.
				await build((p) => {
					p.addSlide().addTable(bodyRows(40), {
						x: '5%',
						y: '5%',
						w: 9, // numeric `w` beside string x/y/h — the other side of the same guard
						h: '60%',
						margin: 0,
						autoPage: true,
						fontSize: 12,
						verbose: true,
					})
				})
			})

			assertEqual(lineMatching(lines, /\| tableProps\.x /).replace(/.*= /, ''), '5%', 'string `x` should pass through')
			assertEqual(lineMatching(lines, /\| tableProps\.h /).replace(/.*= /, ''), '60%', 'string `h` should pass through')
			assert(
				/\| tableProps\.w .*= [\d.]+$/.test(lineMatching(lines, /\| tableProps\.w /)),
				'numeric `w` should still print through the number arm'
			)
			assert(
				!lines.some((l) => l.includes('NaN')),
				'no traced value should be NaN; got: ' + lines.filter((l) => l.includes('NaN')).join(' | ')
			)
		},
	},
])
