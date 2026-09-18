import { ChartType } from '../../../dist/node.js'
import { assertEqual, build, defineRegressionSuite } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// A combo subchart with `showLegend: false` deletes its series' legend entries, and
// `c:legendEntry/c:idx` counts positions in the legend rather than series indexes: PowerPoint,
// deleting the first entry of a legend that lists series 1, 2 and 0, writes idx 0
// (test/read/fixtures/chart-legend-entry.pptx). Where each series sits in the legend was read off
// PowerPoint renders (test/read/fixtures/authoring/probe-combo-legend-order.mjs): a scatter's X
// row is not a series, one axis group lists its subcharts in the order given, and across axis
// groups they are ranked by chart type -- area, bar, then line and radar tied, then scatter --
// with the primary axis group listed first between two of equal rank. That ranking was measured
// over every ordered pair in both directions; before it, only bar's lead over line and scatter
// was known, so an area or radar subchart kept its given order and a deletion could address the
// wrong entry.

const labels = ['P', 'Q', 'R']
const series = (...names) => names.map((name, i) => ({ name, labels, values: [1 + i, 2 + i, 3 + i] }))
const secondary = { secondaryValAxis: true, secondaryCatAxis: true }
const axes = { valAxes: [{}, {}], catAxes: [{}, { catAxisHidden: true }] }
const hidden = { showLegend: false }
const scatter = (options) => ({
	type: ChartType.scatter,
	data: [
		{ name: 'X', values: [1, 2, 3] },
		{ name: 'Y', labels, values: [2, 3, 4] },
	],
	options,
})

/** The `c:legendEntry/c:idx` values the combo `types` writes, comma-joined. */
async function legendEntries(types, secondaryAxes) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(types, { x: 1, y: 1, w: 8, h: 4, showLegend: true, ...(secondaryAxes ? axes : {}) })
	})
	const xml = await chartXml(zip)
	return [...xml.matchAll(/<c:legendEntry><c:idx val="(\d+)"\/>/g)].map(([, idx]) => idx).join(',')
}

defineRegressionSuite('Combo chart legend entries', [
	{
		name: 'on one axis group a hidden subchart deletes the positions its series hold in the given order',
		fn: async () => {
			const types = [
				{ type: ChartType.bar, data: series('B', 'C'), options: {} },
				{ type: ChartType.line, data: series('A'), options: hidden },
			]
			assertEqual(await legendEntries(types, false), '2', 'legend B C A: A is entry 2')
		},
	},
	{
		name: 'a hidden scatter deletes one entry for its Y series, listed after the bars on the other axis group',
		fn: async () => {
			const bar = { type: ChartType.bar, data: series('B', 'C'), options: {} }
			assertEqual(await legendEntries([scatter({ ...secondary, ...hidden }), bar], true), '2', 'scatter given first')
			assertEqual(await legendEntries([bar, scatter({ ...secondary, ...hidden })], true), '2', 'scatter given last')
		},
	},
	{
		name: 'across axis groups a hidden bar subchart is listed first, wherever it is given',
		fn: async () => {
			const types = [
				{ type: ChartType.line, data: series('A1', 'A2'), options: secondary },
				{ type: ChartType.bar, data: series('B'), options: hidden },
			]
			assertEqual(await legendEntries(types, true), '0', 'legend B A1 A2: B is entry 0')
		},
	},
	{
		// Each of these is one measured pairing, with the SECOND subchart hidden so the entry index
		// it writes says where the writer thinks that subchart's series sits.
		name: 'across axis groups the rank is by chart type, and area outranks bar',
		fn: async () => {
			const one = (type, name, options) => ({ type, data: series(name), options })
			/** `first` on the secondary axis group, `second` on the primary and hidden. */
			const pair = (first, second) => legendEntries([one(first, 'A', secondary), one(second, 'B', hidden)], true)

			assertEqual(await pair(ChartType.area, ChartType.bar), '1', 'area leads bar')
			assertEqual(await pair(ChartType.bar, ChartType.area), '0', 'and still leads it from the other axis')
			assertEqual(await pair(ChartType.bar, ChartType.line), '1', 'bar leads line')
			assertEqual(await pair(ChartType.bar, ChartType.radar), '1', 'and radar')
			// A scatter's first data row is X values rather than a series, so it needs the builder
			// above: a one-row scatter holds no series at all and writes no entry to compare.
			assertEqual(
				await legendEntries([one(ChartType.line, 'A', secondary), scatter(hidden)], true),
				'1',
				'line leads scatter'
			)
			assertEqual(
				await legendEntries([one(ChartType.radar, 'A', secondary), scatter(hidden)], true),
				'1',
				'radar leads scatter'
			)
		},
	},
	{
		name: 'line and radar tie, and the primary axis group is listed first between them',
		fn: async () => {
			// Both orderings of this pair listed the subchart on the PRIMARY axis first, which is
			// not a rank: it is the tiebreak. The same holds for two subcharts of one type.
			const one = (type, name, options) => ({ type, data: series(name), options })
			const pair = (first, second) => legendEntries([one(first, 'A', secondary), one(second, 'B', hidden)], true)

			assertEqual(await pair(ChartType.line, ChartType.radar), '0', 'the primary radar is listed first')
			assertEqual(await pair(ChartType.radar, ChartType.line), '0', 'and so is the primary line')
			assertEqual(await pair(ChartType.bar, ChartType.bar), '0', 'two bars rank equal, primary first')
		},
	},
])
