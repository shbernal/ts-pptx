// Probe; produces no committed fixture. Builds the combo charts `probe-combo-legend-order.ps1`
// renders in PowerPoint, to find the order a combo chart's legend lists its series in, which is
// what `c:legendEntry/c:idx` counts (`../chart-legend-entry.pptx`).
//
// Every series is given a distinct solid colour, so the legend order can be read off the exported
// PNG's swatches by `read-combo-legend-order.mjs` rather than off the picture by eye. The first
// round of this probe was read by eye, which is why only three pairings were measured; a pixel
// reader is what made the rest tractable.
//
// Measured 2026-09-14, legend read off each PNG by eye:
//   same-axis-line-bar      line A, bar B C, one axis group          A B C   (plotArea order)
//   same-axis-bar-line      bar B C, line A, one axis group          B C A   (plotArea order)
//   secondary-line-bar      line A on the secondary axis, bar B C    B C A
//   secondary-bar-line      bar B C on the secondary axis, line A    B C A
//   secondary-bar1-line2    bar B on the secondary axis, line A1 A2  B A1 A2
//   secondary-line2-bar1    line A1 A2 on the secondary axis, bar B  B A1 A2
//   secondary-scatter-bar   scatter Y on the secondary axis, bar B C B C Y
//
// Measured 2026-09-18 by swatch colour, read by `read-combo-legend-order.mjs`. Every ordered pair
// of the five types that can share a combo was built twice, once with each subchart on the
// secondary axis group, so a type rank can be told apart from an axis-group rule. Twenty pairs
// plus five controls, all readable. The findings:
//
//   RANK, by chart type, and it beats the axis group in every pairing measured:
//       area  <  bar  <  { line = radar }  <  scatter
//   Each of the ten unordered pairs was measured in both directions and both agreed, except
//   line against radar: `rank-line-radar` and `rank-radar-line` BOTH listed the second-given
//   subchart first, which is not a rank at all. Line and radar tie.
//
//   THE TIEBREAK, for two subcharts of equal rank: the PRIMARY axis group is listed first,
//   whatever order they were given in. `ctrl-line-radar` and `ctrl-radar-line` put the
//   first-given subchart on the primary axis and both kept the given order, which with the two
//   `rank-` decks above pins it to the axis rather than to the order. `ctrl-same-bar-1st` and
//   `ctrl-same-bar-2nd` confirm it for two subcharts of the SAME type.
//
//   AREA OUTRANKS BAR, which the 2026-09-14 round could not have seen: it only ever paired bar
//   against line and scatter. `ctrl-bar-area` puts bar on the primary axis and area still leads.
//
//   BUBBLE cannot appear in a combo at all -- the write path refuses it, citing PowerPoint -- so
//   its rank is moot. The ten refused pairings are recorded rather than dropped, so nobody
//   re-runs them looking for an answer that does not exist.
//
// That ranking and its tiebreak are what `legendOrder` in `src/gen/chart/chart-xml.ts` encodes.
import fs from 'node:fs'
import TsPptx, { ChartType } from '../../../../dist/node.js'

const outDir = new URL('../../../../.tmp/combo-legend-order/', import.meta.url)
fs.rmSync(outDir, { recursive: true, force: true })
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

/**
 * The six chart types a combo can hold, each as a one-series subchart whose series is painted
 * `color`. One series apiece keeps the swatch-to-type mapping unambiguous.
 */
const ONE_SERIES = {
	bar: (name, color, options) => ({
		type: ChartType.bar,
		data: series(name),
		options: { ...options, chartColors: [color] },
	}),
	line: (name, color, options) => ({
		type: ChartType.line,
		data: series(name),
		options: { ...options, chartColors: [color] },
	}),
	area: (name, color, options) => ({
		type: ChartType.area,
		data: series(name),
		options: { ...options, chartColors: [color] },
	}),
	radar: (name, color, options) => ({
		type: ChartType.radar,
		data: series(name),
		options: { ...options, chartColors: [color] },
	}),
	scatter: (name, color, options) => ({
		type: ChartType.scatter,
		data: [
			{ name: `${name}x`, values: [1, 2, 3] },
			{ name, labels, values: [2, 3, 4] },
		],
		options: { ...options, chartColors: [color] },
	}),
	bubble: (name, color, options) => ({
		type: ChartType.bubble,
		data: [
			{ name: `${name}x`, values: [1, 2, 3] },
			{ name, labels, values: [2, 3, 4], sizes: [4, 5, 6] },
		],
		options: { ...options, chartColors: [color] },
	}),
}

/** The colour each position in the *given* order is painted, so a swatch names its subchart. */
const COLORS = ['FF0000', '0000FF']

const TYPES = Object.keys(ONE_SERIES)

/** Every ordered pair of distinct types: the rank is not assumed symmetric until it is measured. */
const pairs = []
for (const first of TYPES) {
	for (const second of TYPES) {
		if (first !== second) pairs.push([first, second])
	}
}

const decks = {
	'same-axis-line-bar': { types: [line(['A']), bar(['B', 'C'])], secondaryAxes: false },
	'same-axis-bar-line': { types: [bar(['B', 'C']), line(['A'])], secondaryAxes: false },
	'secondary-line-bar': { types: [line(['A'], secondary), bar(['B', 'C'])], secondaryAxes: true },
	'secondary-bar-line': { types: [bar(['B', 'C'], secondary), line(['A'])], secondaryAxes: true },
	'secondary-bar1-line2': { types: [bar(['B'], secondary), line(['A1', 'A2'])], secondaryAxes: true },
	'secondary-line2-bar1': { types: [line(['A1', 'A2'], secondary), bar(['B'])], secondaryAxes: true },
	'secondary-scatter-bar': { types: [scatter(secondary), bar(['B', 'C'])], secondaryAxes: true },
}

// The pair decks, named `rank-<first>-<second>`: `first` is given first and sits on the secondary
// axis group, so the two are never in one group and the legend has to rank them by type.
for (const [first, second] of pairs) {
	decks[`rank-${first}-${second}`] = {
		types: [ONE_SERIES[first](`${first}1`, COLORS[0], secondary), ONE_SERIES[second](`${second}2`, COLORS[1], {})],
		secondaryAxes: true,
	}
}

const refused = []
for (const [name, { types, secondaryAxes }] of Object.entries(decks)) {
	const pres = new TsPptx()
	try {
		pres.addSlide().addChart(types, {
			x: 0.5,
			y: 0.5,
			w: 9,
			h: 4.5,
			showLegend: true,
			legendPos: 'b',
			...(secondaryAxes ? axes : {}),
		})
		fs.writeFileSync(new URL(`${name}.pptx`, outDir), Buffer.from(await pres.toBytes()))
		console.log('wrote', name)
	} catch (error) {
		// A combination the write path refuses is an answer too: record it rather than dropping it,
		// so the next reader does not spend the run finding out again.
		refused.push(`${name}: ${error instanceof Error ? error.message : String(error)}`)
		console.log('REFUSED', name, '-', error instanceof Error ? error.message : String(error))
	}
}
if (refused.length > 0) {
	fs.writeFileSync(new URL('refused.txt', outDir), `${refused.join('\n')}\n`)
}
