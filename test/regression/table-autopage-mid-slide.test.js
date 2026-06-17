import { defineRegressionSuite, build, listEntries, readEntry, assert } from '../helpers.js'

// Regression (upstream gitbrent/PptxGenJS#1264): an autoPage table that starts mid-slide with an
// explicit height `h` rendered only a few rows on the FIRST slide while later slides filled up.
//
// Root cause: the first-slide usable-height calc subtracted the start-Y from `h`, but `h` is the
// table's height (an extent), not a bottom coordinate. A table at y=3" with h=4" got a first page
// of only h - y - margin ≈ 0.5" instead of the full 4". Later slides already clamped to `h`, so
// they looked correct — making the first page the obvious outlier.

function rows(n) {
	return Array.from({ length: n }, (_, i) => [{ text: `Row ${i} col A` }, { text: `Row ${i} col B` }])
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

const baseOpts = {
	x: 0.5,
	w: 9,
	h: 4,
	colW: [4.5, 4.5],
	margin: 0,
	slideMargin: 0,
	autoPage: true,
	fontSize: 12,
}

defineRegressionSuite('Table autoPage mid-slide first-page row count (upstream #1264)', [
	{
		name: 'a table starting mid-slide fills its first page (not just a few rows)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows(40), { ...baseOpts, y: 3 })
			})
			const counts = await rowsPerSlide(zip)
			// The first slide must hold a full page worth of rows for the explicit h, not a sliver.
			assert(counts.length >= 2, `expected overflow to multiple slides; got ${JSON.stringify(counts)}`)
			assert(
				counts[0] >= 15,
				`first slide should fill the explicit h (~19 rows), not a few; got ${JSON.stringify(counts)}`
			)
		},
	},
	{
		name: 'first-page row count is independent of where the table starts (y)',
		fn: async () => {
			const { zip: zipTop } = await build((p) => {
				p.addSlide().addTable(rows(40), { ...baseOpts, y: 0.5 })
			})
			const { zip: zipMid } = await build((p) => {
				p.addSlide().addTable(rows(40), { ...baseOpts, y: 3 })
			})
			const top = await rowsPerSlide(zipTop)
			const mid = await rowsPerSlide(zipMid)
			// With an explicit `h`, the usable area is the same regardless of `y`, so the first
			// page (and therefore the whole pagination) must match.
			assert(
				top[0] === mid[0],
				`first-page row count must not depend on y when h is explicit; top=${JSON.stringify(top)} mid=${JSON.stringify(mid)}`
			)
		},
	},
])
