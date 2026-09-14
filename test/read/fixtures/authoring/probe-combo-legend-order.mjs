// Probe; produces no committed fixture. Builds the combo charts `probe-combo-legend-order.ps1`
// renders in PowerPoint, to find the order a combo chart's legend lists its series in, which is
// what `c:legendEntry/c:idx` counts (`../chart-legend-entry.pptx`).
//
// Measured 2026-09-14, legend read off each PNG:
//   same-axis-line-bar      line A, bar B C, one axis group          A B C   (plotArea order)
//   same-axis-bar-line      bar B C, line A, one axis group          B C A   (plotArea order)
//   secondary-line-bar      line A on the secondary axis, bar B C    B C A
//   secondary-bar-line      bar B C on the secondary axis, line A    B C A
//   secondary-bar1-line2    bar B on the secondary axis, line A1 A2  B A1 A2
//   secondary-line2-bar1    line A1 A2 on the secondary axis, bar B  B A1 A2
//   secondary-scatter-bar   scatter Y on the secondary axis, bar B C B C Y
// Within one axis group the legend follows plotArea order. Across axis groups the bar group is
// listed first whichever axis it is on and however many series either group has, so the order is
// by chart type, not by axis or by plotArea. Only bar ahead of line and ahead of scatter is
// measured; how the other types rank against each other is not.
import fs from 'node:fs'
import TsPptx, { ChartType } from '../../../../dist/node.js'

const outDir = new URL('../../../../.tmp/combo-legend-order/', import.meta.url)
fs.mkdirSync(outDir, { recursive: true })

const labels = ['P', 'Q', 'R']
const secondary = { secondaryValAxis: true, secondaryCatAxis: true }
const axes = { valAxes: [{}, {}], catAxes: [{}, { catAxisHidden: true }] }
const series = (...names) => names.map((name, i) => ({ name, labels, values: [1 + i, 2 + i, 3 + i] }))
const bar = (names, options = {}) => ({ type: ChartType.bar, data: series(...names), options })
const line = (names, options = {}) => ({ type: ChartType.line, data: series(...names), options })
const scatter = (options = {}) => ({
	type: ChartType.scatter,
	data: [
		{ name: 'X', values: [1, 2, 3] },
		{ name: 'Y', labels, values: [2, 3, 4] },
	],
	options,
})

const decks = {
	'same-axis-line-bar': { types: [line(['A']), bar(['B', 'C'])], secondaryAxes: false },
	'same-axis-bar-line': { types: [bar(['B', 'C']), line(['A'])], secondaryAxes: false },
	'secondary-line-bar': { types: [line(['A'], secondary), bar(['B', 'C'])], secondaryAxes: true },
	'secondary-bar-line': { types: [bar(['B', 'C'], secondary), line(['A'])], secondaryAxes: true },
	'secondary-bar1-line2': { types: [bar(['B'], secondary), line(['A1', 'A2'])], secondaryAxes: true },
	'secondary-line2-bar1': { types: [line(['A1', 'A2'], secondary), bar(['B'])], secondaryAxes: true },
	'secondary-scatter-bar': { types: [scatter(secondary), bar(['B', 'C'])], secondaryAxes: true },
}
for (const [name, { types, secondaryAxes }] of Object.entries(decks)) {
	const pres = new TsPptx()
	pres
		.addSlide()
		.addChart(types, { x: 0.5, y: 0.5, w: 9, h: 4.5, showLegend: true, legendPos: 'b', ...(secondaryAxes ? axes : {}) })
	fs.writeFileSync(new URL(`${name}.pptx`, outDir), Buffer.from(await pres.toBytes()))
	console.log('wrote', name)
}
