import { ChartType } from '../../../dist/node.js'
import { assert, assertEqual, build, captureDiagnostics, defineRegressionSuite } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// One chart part used to carry two readings of the same data-label font option. Two builders
// emit the run properties a `<c:dLbls>` wraps, and they defaulted differently -- `??` in one,
// `||` in the other -- so `dataLabelFontSize: 0` came out as `sz="100"` (clamped, warned) beside
// `sz="1200"` (silently defaulted) on labels of the same chart. Which reading a chart took
// depended on its type and its label format: a pie reached only the `||` builder and dropped the
// value with no diagnostic at all, an `XY` scatter reached both at once and emitted both, and a
// bar reached only the `??` one and warned. One input, three contracts.
//
// Each case in the first suite is shaped to reach the label builders its type actually uses --
// a plain scatter takes the chart-level block, and only `dataLabelFormatScatter` brings its own
// per-series ones into play. The `sz` values outside `<c:dLbls>` are a different option (axis
// tick labels default to 12pt), so those assertions look only inside the label blocks.

const BASE = { x: 1, y: 1, w: 4, h: 3 }

/**
 * Every `sz` on a run-properties tag inside a `<c:dLbls>` block, deduplicated. Both `<a:defRPr>`
 * and `<a:rPr>` are read: a scatter's rich label carries the same options on the run itself.
 */
function labelFontSizes(xml) {
	const blocks = [...xml.matchAll(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g)].map((m) => m[0])
	assert(blocks.length > 0, 'expected at least one <c:dLbls> block in the chart part')
	return [...new Set(blocks.flatMap((b) => [...b.matchAll(/<a:(?:defRPr|rPr)[^>]*\ssz="(\d+)"/g)].map((m) => m[1])))]
}

/** Build a one-chart deck and hand back its label sizes plus whatever it warned about. */
async function labelSizesFor(fixture, opts) {
	const { result, codes } = await captureDiagnostics(() =>
		build((p) => {
			p.addSlide().addChart(fixture.data, { ...BASE, ...fixture.opts, ...opts })
		})
	)
	return { sizes: labelFontSizes(await chartXml(result.zip)), codes }
}

const CAT_DATA = [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }]
const XY_DATA = [
	{ name: 'X', values: [1, 2] },
	{ name: 'A', values: [3, 4], sizes: [1, 2] },
]

const FIXTURES = [
	{ label: 'bar', data: CAT_DATA, opts: { type: ChartType.bar, showValue: true } },
	{ label: 'pie', data: CAT_DATA, opts: { type: ChartType.pie, showValue: true } },
	{ label: 'bubble', data: XY_DATA, opts: { type: ChartType.bubble, showValue: true } },
	{
		label: 'XY scatter',
		data: XY_DATA,
		opts: { type: ChartType.scatter, showValue: true, showLabel: true, dataLabelFormatScatter: 'XY' },
	},
]

/** The `<a:defRPr>` of the first `<c:dLbls>` block, whole. */
function firstLabelDefRPr(xml) {
	const blocks = [...xml.matchAll(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g)].map((m) => m[0]).join('')
	const match = /<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/.exec(blocks)
	assert(match, 'expected an <a:defRPr> inside a <c:dLbls> block')
	return match[0]
}

/** Build a one-chart deck and hand back its first label `<a:defRPr>`. */
async function labelDefRPrFor(fixture, opts) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(fixture.data, { ...BASE, ...fixture.opts, ...opts })
	})
	return firstLabelDefRPr(await chartXml(zip))
}

const BAR = FIXTURES[0]
const PIE = FIXTURES[1]

defineRegressionSuite(
	'Data-label font options have one reading per chart',
	FIXTURES.flatMap((fixture) => [
		{
			name: `a ${fixture.label} clamps dataLabelFontSize: 0 to the schema minimum, once, and says so`,
			fn: async () => {
				const { sizes, codes } = await labelSizesFor(fixture, { dataLabelFontSize: 0 })
				assertEqual(sizes.join(', '), '100', `${fixture.label} label sz values`)
				assert(
					codes.includes('font/size-out-of-range'),
					`${fixture.label} must warn that 0 is out of range; got ${JSON.stringify(codes)}`
				)
			},
		},
		{
			name: `a ${fixture.label} carries one dataLabelFontSize on every label block`,
			fn: async () => {
				const { sizes, codes } = await labelSizesFor(fixture, { dataLabelFontSize: 9 })
				assertEqual(sizes.join(', '), '900', `${fixture.label} label sz values`)
				assertEqual(codes.length, 0, `${fixture.label} must not warn about a size in range`)
			},
		},
	]).concat([
		// The other two spellings of "stated, but empty". `??` lets both through where `||`
		// silently substituted a default, so what they now mean has to be the same on both
		// builders -- and it is, because the colour goes through the one helper whose contract
		// says a value naming nothing to paint emits nothing.
		{
			name: "an empty dataLabelColor leaves the label's fill inherited, on either builder",
			fn: async () => {
				const bar = await labelDefRPrFor(BAR, { dataLabelColor: '' })
				const pie = await labelDefRPrFor(PIE, { dataLabelColor: '' })
				assert(!bar.includes('<a:solidFill>'), `bar emitted a fill for an empty colour: ${bar}`)
				assert(!pie.includes('<a:solidFill>'), `pie emitted a fill for an empty colour: ${pie}`)
			},
		},
		{
			name: 'an empty dataLabelFontFace emits an empty typeface rather than Arial, on either builder',
			fn: async () => {
				const bar = await labelDefRPrFor(BAR, { dataLabelFontFace: '' })
				const pie = await labelDefRPrFor(PIE, { dataLabelFontFace: '' })
				assert(bar.includes('<a:latin typeface=""/>'), `bar substituted a typeface: ${bar}`)
				assert(pie.includes('<a:latin typeface=""/>'), `pie substituted a typeface: ${pie}`)
			},
		},
	])
)

// The same defaulting shape reached past the data labels: every chart option that takes a font
// size, a colour or a typeface read its default through `||`, so a stated `0` or `''` meant
// "the caller said nothing" on options where those are values a caller can mean. Absence is now
// the only spelling of the default, and a stated falsy value is diagnosed rather than replaced.

/** Build a chart carrying a title, an axis title and a data table, and report what it warned. */
async function sweepChart(opts) {
	const { result, codes } = await captureDiagnostics(() =>
		build((p) => {
			p.addSlide().addChart(CAT_DATA, {
				...BASE,
				type: ChartType.bar,
				showTitle: true,
				title: 'T',
				showCatAxisTitle: true,
				catAxisTitle: 'C',
				showDataTable: true,
				...opts,
			})
		})
	)
	return { xml: await chartXml(result.zip), codes }
}

/**
 * Every `<a:defRPr>` inside the first element matching `tag`. A `<c:catAx>` holds two -- its
 * title's and its tick labels' -- and both read options this sweep touched, so both are checked.
 */
function defRPrsIn(xml, tag) {
	const start = xml.indexOf(`<${tag}>`)
	assert(start >= 0, `expected a <${tag}> in the chart part`)
	const block = xml.slice(start, xml.indexOf(`</${tag}>`, start))
	const found = [...block.matchAll(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g)].map((m) => m[0])
	assert(found.length > 0, `expected an <a:defRPr> inside <${tag}>`)
	return found
}

defineRegressionSuite('A stated falsy font option is a stated value', [
	...['catAxisLabelFontSize', 'valAxisLabelFontSize', 'titleFontSize', 'catAxisTitleFontSize', 'dataTableFontSize'].map(
		(option) => ({
			name: `${option}: 0 clamps to the schema minimum and warns`,
			fn: async () => {
				const { xml, codes } = await sweepChart({ [option]: 0 })
				assert(
					codes.includes('font/size-out-of-range'),
					`${option}: 0 must warn that it is out of range; got ${JSON.stringify(codes)}`
				)
				assert(xml.includes('sz="100"'), `${option}: 0 must reach the part as the 1pt minimum`)
			},
		})
	),
	{
		name: 'an empty colour leaves the fill inherited rather than painting a default',
		fn: async () => {
			const { xml } = await sweepChart({
				catAxisLabelColor: '',
				catAxisTitleColor: '',
				titleColor: '',
			})
			for (const tag of ['c:catAx', 'c:title'])
				for (const defRPr of defRPrsIn(xml, tag))
					assert(!defRPr.includes('<a:solidFill>'), `<${tag}> kept a fill for an empty colour: ${defRPr}`)
		},
	},
	{
		name: 'an empty typeface is emitted rather than substituted',
		fn: async () => {
			const { xml } = await sweepChart({
				catAxisLabelFontFace: '',
				catAxisTitleFontFace: '',
				titleFontFace: '',
			})
			for (const tag of ['c:catAx', 'c:title'])
				for (const defRPr of defRPrsIn(xml, tag))
					assert(defRPr.includes('<a:latin typeface=""/>'), `<${tag}> substituted a typeface: ${defRPr}`)
		},
	},
])
