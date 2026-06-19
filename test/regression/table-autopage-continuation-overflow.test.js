import { defineRegressionSuite, build, listEntries, readEntry, assert } from '../helpers.js'

// Regression (upstream gitbrent/PptxGenJS#1319): autoPage created a continuation slide, but the
// reporter saw rows on that *new* slide overflow past the bottom — pagination was applied to the
// first slide only, not to subsequently generated ones.
//
// The current auto-pager recomputes the usable table height (`calcSlideTabH`) for every generated
// slide, so continuation slides honor the same row budget. This locks that: with many rows and no
// explicit `h`, the middle (fully-filled) continuation slides must all carry an equal row count —
// a stable budget, not a count that grows slide-over-slide (the overflow symptom).

const header = [
	{ text: 'Name', options: { bold: true } },
	{ text: 'Detail', options: { bold: true } },
]

function rows(n) {
	const r = [header]
	for (let i = 0; i < n; i++) {
		r.push([
			{ text: `Row ${i} name` },
			{ text: `Row ${i} has a somewhat longer description that may wrap across a couple of lines in the cell ${i}` },
		])
	}
	return r
}

function slideXmlNames(zip) {
	return listEntries(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

async function rowsPerSlide(zip) {
	const counts = []
	for (const name of slideXmlNames(zip)) {
		const xml = await readEntry(zip, name)
		counts.push((xml.match(/<a:tr /g) || []).length)
	}
	return counts
}

defineRegressionSuite('Table autoPage continuation-slide overflow (upstream #1319)', [
	{
		name: 'continuation slides honor the row budget (no progressive overflow), matching the reported case',
		fn: async () => {
			// The reporter's exact option shape: no explicit `h`, repeated header, newSlideStartY.
			const { zip } = await build((p) => {
				p.defineLayout({ name: 'L1319', width: 10, height: 5.625 })
				p.layout = 'L1319'
				p.addSlide().addTable(rows(30), {
					x: 0.5,
					y: 0.5,
					w: 9,
					autoPage: true,
					autoPageRepeatHeader: true,
					newSlideStartY: 0.5,
				})
			})
			const counts = await rowsPerSlide(zip)
			assert(counts.length >= 3, `expected several overflow slides; got ${JSON.stringify(counts)}`)

			// The "middle" slides are the ones filled to capacity (the first may start differently and
			// the last holds the remainder). They must all carry the SAME row count: a fixed per-slide
			// budget. A count that climbs from one continuation slide to the next is the overflow bug.
			const middle = counts.slice(1, -1)
			const first = middle[0]
			assert(
				middle.every((c) => c === first),
				`continuation slides must share one stable row budget (no progressive overflow); got ${JSON.stringify(counts)}`
			)
		},
	},
])
