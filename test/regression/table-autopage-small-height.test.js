import { defineRegressionSuite, build, listEntries, assert } from '../helpers.js'

// Regression: an autoPage table whose height (`h`) — combined with `y` and margins — leaves no
// usable vertical space must NOT emit a degenerate empty overflow page. That empty `rows:[]`
// page previously made the recursive addTable throw "addTable: Array expected". The paginator
// now ignores the unusable height (falling back to the slide height) and warns.

function rows(n) {
	return Array.from({ length: n }, (_, i) => [{ text: `Row ${i} col A` }, { text: `Row ${i} col B` }])
}

function slideCount(zip) {
	return listEntries(zip).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length
}

defineRegressionSuite('Table autoPage tiny-height guard', [
	{
		name: 'tiny h + large y does not crash and emits no empty page (warns instead)',
		fn: async () => {
			const warnings = []
			const orig = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			let zip
			try {
				;({ zip } = await build((p) => {
					// y(1.2) + bottom margin already exceed h(0.7) → negative usable height.
					p.addSlide().addTable(rows(12), {
						x: 0.5,
						y: 1.2,
						w: 9,
						h: 0.7,
						colW: [4.5, 4.5],
						margin: 0,
						slideMargin: 0,
						autoPage: true,
						fontSize: 12,
					})
				}))
			} finally {
				console.warn = orig
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
			const orig = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
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
				console.warn = orig
			}
			assert(slideCount(zip) >= 2, 'a 60-row table should overflow to multiple slides')
			assert(
				!warnings.some((w) => w.includes('leaves no room to paginate')),
				'usable height must not trigger the guard warning'
			)
		},
	},
])
