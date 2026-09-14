import { ChartType } from '../../../dist/node.js'
import { assertEqual, build, defineRegressionSuite } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// A combo subchart with `showLegend: false` deletes its series' legend entries, and
// `c:legendEntry/c:idx` counts positions in the legend rather than series indexes: PowerPoint,
// deleting the first entry of a legend that lists series 1, 2 and 0, writes idx 0
// (test/read/fixtures/chart-legend-entry.pptx). Where each series sits in the legend was read off
// PowerPoint renders (test/read/fixtures/authoring/probe-combo-legend-order.mjs): a scatter's X
// row is not a series, one axis group lists its subcharts in the order given, and across axis
// groups a bar subchart is listed first.

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
])
