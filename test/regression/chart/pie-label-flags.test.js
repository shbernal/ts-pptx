import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertEqual, captureDiagnostics } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// A pie labels its *points*, not its series, so it emits two kinds of `<c:dLbls>`: one per point
// carrying the overrides, and one at the plot level carrying the defaults. The per-point block
// read the caller's `showLabel`/`showPercent`/`showValue`/`showSerName`; the plot-level block
// wrote `catName: 1, percent: 1` as constants. The constants were masked while every point had a
// `<c:dLbl>` to override them, which holds only while the pie has labels — so an unlabelled pie
// came back with both of the caller's `false`s inverted.
//
// The same input keyed both sheet ranges on the label count, so `count + 1` made them run
// backwards (`Sheet1!$A$2:$A$1`), and emitted no `<c:dPt>` at all.

const BASE = { x: 1, y: 1, w: 6, h: 4 }

/** Every `<c:show*>` flag in the plot-level `<c:dLbls>` (the one that is not inside a `<c:dLbl>`). */
function plotLevelFlags(xml) {
	const withoutPerPoint = xml.replace(/<c:dLbl>[\s\S]*?<\/c:dLbl>/g, '')
	return Object.fromEntries(
		[...withoutPerPoint.matchAll(/<c:show(\w+) val="([01])"\/>/g)].map((m) => [m[1], Number(m[2])])
	)
}

defineRegressionSuite('Pie data-label flags and slice count', [
	{
		name: "the plot-level dLbls follows the caller's show flags, labels or not",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S', values: [10, 20, 30] }], {
					...BASE,
					type: ChartType.pie,
					showPercent: false,
					showLabel: false,
				})
			})
			const flags = plotLevelFlags(await chartXml(zip))
			assertEqual(flags.CatName, 0, 'showLabel:false is not a category name')
			assertEqual(flags.Percent, 0, 'showPercent:false is not a percentage')
		},
	},
	{
		name: 'and still says yes when the caller asked',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S', labels: ['a', 'b'], values: [10, 20] }], {
					...BASE,
					type: ChartType.pie,
					showPercent: true,
					showLabel: true,
				})
			})
			const flags = plotLevelFlags(await chartXml(zip))
			assertEqual(flags.CatName, 1, 'showLabel:true is a category name')
			assertEqual(flags.Percent, 1, 'showPercent:true is a percentage')
		},
	},
	{
		name: 'an unlabelled pie slices its values, with forward sheet ranges',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S', values: [10, 20, 30] }], { ...BASE, type: ChartType.pie })
			})
			const xml = await chartXml(zip)
			assert(!/\$A\$2:\$A\$1|\$B\$2:\$B\$1/.test(xml), 'no reversed range; got: ' + xml)
			assert(xml.includes('Sheet1!$B$2:$B$4'), 'the value range spans the three slices; got: ' + xml)
			assert(!xml.includes('<c:cat>'), 'and states no category names rather than an empty range')
			assertEqual((xml.match(/<c:dPt>/g) || []).length, 3, 'one data point per slice')
		},
	},
	{
		name: 'a pie with neither labels nor values plots nothing, and says so',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart([{ name: 'S', values: [] }], { ...BASE, type: ChartType.pie })
				})
			)
			const xml = await chartXml(result.zip)
			assert(!xml.includes('<c:pieChart>'), 'no plot is emitted; got: ' + xml)
			assert(codes.includes('chart/point-count-mismatch'), 'the caller is told; got ' + JSON.stringify(codes))
		},
	},
])
