import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'

// What a cell inherits from its table, and the agreement between the two paths that resolve it.
//
// The emitter builds the bag it hands to the text-body writer; the measured-fit pass builds the
// effective text it lays out. Each had its own key list, and the measure one's docstring claimed
// to mirror the emitter's while naming four keys the emitter did not carry — so a table-level
// `italic` was measured with italic metrics and emitted upright, and the layout the library
// reported was one the file disagreed with.

const SLIDE_XML = 'ppt/slides/slide1.xml'

/** The first `<a:rPr>` of the first cell on slide 1. */
async function firstRunProps(zip) {
	const xml = await readEntry(zip, SLIDE_XML)
	const rPr = /<a:rPr[^>]*>/.exec(xml)
	assert(rPr, 'expected a text run on the slide; got: ' + xml)
	return rPr[0]
}

defineRegressionSuite('table-level text options a cell inherits', [
	{
		name: 'a table-level `italic` reaches the cell it was measured against',
		fn: async () => {
			// The reachable half of the divergence: `italic` is on `TextBaseProps`, so `TableProps`
			// has it and a typed caller can set it — and it was inherited by the fitter and by
			// nothing else, so it changed the computed layout and never appeared in the file.
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'cell' }]], { x: 1, y: 1, w: 6, h: 1, italic: true })
			})
			assert(/ i="1"/.test(await firstRunProps(zip)), 'the run is italic')
		},
	},
	{
		name: "a cell's own value still wins over the table's",
		fn: async () => {
			// Inheritance fills gaps; it does not overwrite. The guard is `=== undefined`, so a
			// cell's `italic: false` is something the cell said, not something it left unset.
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'cell', options: { italic: false } }]], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
					italic: true,
				})
			})
			assert(!/ i="1"/.test(await firstRunProps(zip)), 'the cell said no')
		},
	},
	{
		name: 'a string cell and an object cell inherit the same values',
		fn: async () => {
			// Two cells with the same text rendering differently by how they were spelled is the
			// failure this pins; the string form takes the same resolution as the object form.
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'same' }, 'same']], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
					italic: true,
					bold: true,
					fontFace: 'Georgia',
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			const runs = [...xml.matchAll(/<a:rPr[^>]*>/g)].map((m) => m[0])
			assertEqual(runs.length, 2, 'two cells, two runs')
			assertEqual(runs[0], runs[1], 'and the same run properties on both')
		},
	},
])
