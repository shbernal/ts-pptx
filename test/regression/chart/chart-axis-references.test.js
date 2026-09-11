import { DOMParser } from '@xmldom/xmldom'
import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, captureDiagnostics } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// Every `<c:axId>` a plot group carries has to name an axis the plot area emits, and the category
// and value axes a plot group names have to cross each other. PowerPoint enforces the second: a
// combo line on the primary category axis and the secondary value axis opened as corrupt
// (0x80070570), whether or not a secondary category axis was emitted as well.
//
// The axes used to be decided by the `valAxes` override list rather than by what the plots
// reference, so a 3-D bar or surface chart with `valAxes` lost its series axis, and a
// `secondaryValAxis` subchart with one `valAxes` entry lost the secondary value axis. A combo
// asking for only one secondary axis referenced a crossing axis that was never emitted at all.

const parser = new DOMParser()
const AXIS_TAGS = ['c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx']
const LABELS = ['a', 'b', 'c']

/** Every disagreement between the plot groups' axis references and the axes the part defines. */
function axisProblems(xml) {
	const doc = parser.parseFromString(xml, 'text/xml')
	const plotArea = doc.getElementsByTagName('c:plotArea')[0]
	const children = Array.from(plotArea.childNodes).filter((node) => node.nodeType === 1)
	const child = (node, name) => Array.from(node.childNodes).find((c) => c.nodeName === name)
	const val = (node, name) => child(node, name)?.getAttribute('val')

	/** axis id -> { tag, crossAx } */
	const axes = new Map()
	for (const node of children.filter((n) => AXIS_TAGS.includes(n.nodeName))) {
		const id = val(node, 'c:axId')
		if (axes.has(id)) return [`axis id ${id} is defined twice`]
		axes.set(id, { tag: node.nodeName, crossAx: val(node, 'c:crossAx') })
	}

	const problems = []
	for (const [id, axis] of axes) {
		if (!axes.has(axis.crossAx)) problems.push(`${axis.tag} ${id} crosses ${axis.crossAx}, which no axis carries`)
	}
	for (const plot of children.filter((n) => n.nodeName.endsWith('Chart'))) {
		const refs = Array.from(plot.childNodes)
			.filter((n) => n.nodeName === 'c:axId')
			.map((n) => /** @type {any} */ (n).getAttribute('val'))
		for (const ref of refs) {
			if (!axes.has(ref)) problems.push(`${plot.nodeName} references axis ${ref}, which no axis carries`)
		}
		const [catId, valId] = refs
		const cat = axes.get(catId)
		const value = axes.get(valId)
		if (cat && value && (cat.crossAx !== valId || value.crossAx !== catId))
			problems.push(`${plot.nodeName} plots on ${catId} and ${valId}, which do not cross each other`)
	}
	return problems
}

/** A single-type chart, or a bar + line combo whose line takes `lineOptions`. */
function chartCall(type, lineOptions) {
	if (type !== 'combo') {
		const data =
			type === ChartType.scatter
				? [
						{ name: 'X', values: [1, 2, 3] },
						{ name: 'Y', values: [4, 5, 6] },
					]
				: [{ name: 'S', labels: LABELS, values: [1, 2, 3] }]
		return { types: type, data }
	}
	return {
		types: [
			{ type: ChartType.bar, data: [{ name: 'Bar', labels: LABELS, values: [1, 2, 3] }], options: {} },
			{ type: ChartType.line, data: [{ name: 'Line', labels: LABELS, values: [40, 50, 60] }], options: lineOptions },
		],
	}
}

const SINGLE_TYPES = [
	ChartType.bar,
	ChartType.bar3d,
	ChartType.line,
	ChartType.area,
	ChartType.radar,
	ChartType.scatter,
	ChartType.surface,
]
const SECONDARY_FLAGS = [
	['no secondary axis', {}],
	['secondaryValAxis', { secondaryValAxis: true }],
	['secondaryCatAxis', { secondaryCatAxis: true }],
	['both secondary axes', { secondaryValAxis: true, secondaryCatAxis: true }],
]
const OVERRIDES = [
	['no valAxes', {}],
	['one valAxes entry', { valAxes: [{ valAxisTitle: 'V', showValAxisTitle: true }] }],
	[
		'two valAxes and catAxes entries',
		{
			valAxes: [{ valAxisTitle: 'V1' }, { valAxisTitle: 'V2' }],
			catAxes: [{ catAxisTitle: 'C1' }, { catAxisTitle: 'C2' }],
		},
	],
]

/** `[name, call, overrides]` for every shape the matrix covers. */
const MATRIX = [
	...SINGLE_TYPES.map((type) => [String(type), chartCall(type)]),
	...SECONDARY_FLAGS.map(([flags, lineOptions]) => [`bar + line combo, ${flags}`, chartCall('combo', lineOptions)]),
].flatMap(([name, call]) => OVERRIDES.map(([overrideName, overrides]) => [`${name}, ${overrideName}`, call, overrides]))

defineRegressionSuite('Chart axis references resolve to emitted, crossing axes', [
	{
		name: 'every plot group references axes the plot area emits, in crossing pairs',
		fn: async () => {
			const failures = []
			for (const [name, call, overrides] of MATRIX) {
				const frame = { x: 1, y: 1, w: 6, h: 4, ...overrides }
				const options = Array.isArray(call.types) ? frame : { ...frame, type: call.types }
				let xml
				try {
					const { zip } = await build((p) => {
						if (Array.isArray(call.types)) p.addSlide().addChart(call.types, options)
						else p.addSlide().addChart(call.data, options)
					})
					xml = await chartXml(zip)
				} catch (error) {
					// Two value axes where no subchart plots on the secondary pair is refused on purpose.
					if (/secondary-axis-unused|Secondary axis must be used/.test(String(error?.code ?? error?.message))) continue
					failures.push(`${name}: threw ${error?.message}`)
					continue
				}
				const problems = axisProblems(xml)
				if (problems.length) failures.push(`${name}:\n    ${problems.join('\n    ')}`)
			}
			assert(failures.length === 0, `\n  ${failures.join('\n  ')}`)
		},
	},
	{
		// The axis a subchart did not ask for is emitted hidden, the way PowerPoint shows a series
		// moved onto its secondary axis; an override can still show it.
		name: 'the secondary axis no subchart asked for is hidden unless an override shows it',
		fn: async () => {
			const secondaryCatDelete = async (overrides) => {
				const { zip } = await build((p) =>
					p.addSlide().addChart(chartCall('combo', { secondaryValAxis: true }).types, {
						x: 1,
						y: 1,
						w: 6,
						h: 4,
						...overrides,
					})
				)
				const xml = await chartXml(zip)
				const axis = xml.match(/<c:catAx><c:axId val="2094734555"\/>[\s\S]*?<\/c:catAx>/)?.[0] ?? ''
				return axis.match(/<c:delete val="(\d)"\/>/)?.[1]
			}
			assert((await secondaryCatDelete({})) === '1', 'the unrequested secondary category axis is hidden')
			const shown = await secondaryCatDelete({
				valAxes: [{}, {}],
				catAxes: [{}, { catAxisHidden: false }],
			})
			assert(shown === '0', `catAxes[1] can show it; got delete=${shown}`)
		},
	},
	{
		name: 'axis overrides past the second warn that they are ignored',
		fn: async () => {
			const { codes } = await captureDiagnostics(() =>
				build((p) =>
					p.addSlide().addChart(chartCall('combo', { secondaryValAxis: true }).types, {
						x: 1,
						y: 1,
						w: 6,
						h: 4,
						valAxes: [{}, {}, {}],
					})
				)
			)
			assert(codes.includes('chart/option-not-supported'), `expected a warning; got ${JSON.stringify(codes)}`)
		},
	},
])
