import { defineRegressionSuite, build, listEntries, readEntry, assert, assertEqual } from '../../helpers.js'

// Regression: upstream gitbrent/PptxGenJS#1200 — "tableToSlides autoPaging not working": a table
// paged onto several slides, but the rows on the generated slides ran off the bottom edge.
//
// The cause is not in the DOM half of that report, which is why it took so long to pin. The pager
// charges each row its cells' top/bottom margins before deciding whether the row fits
// (`emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu`), and on a page break it did that and then
// immediately zeroed the accumulator. So the first row of every continuation slide — and only that
// row — was placed for free. One row's margins is not much on its own, but it is enough to let a
// continuation page accept a row the page has no room for, and the deeper the cell padding the
// further that row hangs off the slide.
//
// ── The oracle ─────────────────────────────────────────────────────────────────────────────────
//
// "Does the table overflow the slide" is a question about a rendered page, and this repo has no
// render oracle for one. It does not need one here. Every row below is identical, and with
// `autoPageSlideStartY` equal to the top margin every page has the identical usable height — so
// however many rows fit, that number is the same on every full page. The first page's budget was
// never in doubt (it charges margins for every row it takes, first one included), so the defect
// shows as an arithmetic disagreement between page 1 and its continuations about how many
// identical rows fit an identical space. Before the fix: 10, then 11, 11, 11.
//
// This is the DOM-free repro the dismissal asked for, and it makes the same case in the
// browser (test/browser/table-autopage.spec.mjs) a confirmation rather than the evidence.

/** Enough rows to fill several pages, all identical so every page's budget is comparable. */
function uniformRows(count, margin) {
	const options = { fontSize: 16, margin: [margin, margin, margin, margin] }
	const rows = []
	for (let idx = 1; idx <= count; idx++) {
		rows.push([
			{ text: `R${idx}`, options },
			// Short enough that no cell can wrap at the column widths below: a wrapped cell makes
			// its row taller than its neighbours', and then "the same number of rows" stops being
			// the same amount of vertical space and the comparison below means nothing.
			{ text: `V${idx}`, options },
		])
	}
	return rows
}

function slideXmlNames(zip) {
	return listEntries(zip)
		.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

async function rowsPerSlide(zip) {
	const counts = []
	for (const name of slideXmlNames(zip)) {
		counts.push(((await readEntry(zip, name)).match(/<a:tr /g) || []).length)
	}
	return counts
}

/** Every cell's text, slide by slide, so nothing can be dropped or duplicated unnoticed. */
async function cellTextPerSlide(zip) {
	const perSlide = []
	for (const name of slideXmlNames(zip)) {
		const xml = await readEntry(zip, name)
		perSlide.push([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) => match[1]))
	}
	return perSlide
}

const paged = (margin, rowCount = 60) =>
	build((pptx) => {
		pptx.defineLayout({ name: 'L1200', width: 10, height: 5.625 })
		pptx.layout = 'L1200'
		pptx.addSlide().addTable(uniformRows(rowCount, margin), {
			x: 0.5,
			y: 0.5,
			w: 9,
			colW: [3, 6],
			autoPage: true,
			autoPageSlideStartY: 0.5,
		})
	})

defineRegressionSuite('Table autoPage continuation-slide row budget (gitbrent/PptxGenJS#1200)', [
	{
		name: 'a continuation slide takes no more rows than the first slide did, at the same geometry',
		fn: async () => {
			const { zip } = await paged(0.0833) // 8px of cell padding, the shape an HTML table arrives in
			const counts = await rowsPerSlide(zip)
			assert(counts.length >= 3, `expected several pages; got ${JSON.stringify(counts)}`)

			// All but the last: the last holds whatever remains and is allowed to be short.
			const full = counts.slice(0, -1)
			const first = full[0]
			assert(
				full.every((count) => count === first),
				'every full page must hold the same number of identical rows; the pages disagreeing means a ' +
					`continuation page took a row it had no room for (gitbrent/PptxGenJS#1200). Got ${JSON.stringify(counts)}`
			)
			// Guard on the shape of the fixture itself: one row per page would make the assertion
			// above trivially true, and would mean the budget arithmetic collapsed rather than held.
			assert(first > 3, `expected a real page budget; got ${first} rows per page`)
		},
	},
	{
		name: 'the deeper the cell margins, the more the old accounting lost — still one budget',
		fn: async () => {
			// The bug leaked exactly one row's worth of top+bottom margin per page, so a fixture with
			// deeper padding is where it did the most damage. Ordinary padding at 0.2in per side, not
			// an extreme: >= 1in would be read as a legacy points value and warn.
			const { zip } = await paged(0.2)
			const counts = await rowsPerSlide(zip)
			const full = counts.slice(0, -1)
			assert(full.length >= 2, `expected several full pages; got ${JSON.stringify(counts)}`)
			assert(
				full.every((count) => count === full[0]),
				`deep cell margins must not buy a continuation page an extra row; got ${JSON.stringify(counts)}`
			)
		},
	},
	{
		name: 'pagination moves every row across exactly once, in order',
		fn: async () => {
			const { zip } = await paged(0.0833, 30)
			const seen = (await cellTextPerSlide(zip)).flat().filter((text) => /^R\d+$/.test(text))
			assertEqual(
				seen.join(','),
				Array.from({ length: 30 }, (_, idx) => `R${idx + 1}`).join(','),
				'every source row should appear once, in order, across the generated slides'
			)
		},
	},
])
