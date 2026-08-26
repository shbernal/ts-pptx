import { ChartType } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	captureDiagnostics,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
} from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// The fragments several plot builders share, pinned at the byte.
//
// These are the emitters the byte-identity harness cannot speak for: its corpus is the two
// showcase decks, whose only charts are bar, doughnut and line. Scatter, bubble, stock and
// surface are unproven there, not proven unchanged — so a refactor that touched them was
// checked against a throwaway differential, and this file is what makes the next one cheaper.
//
// Two of the cases below pin something that looks like an accident and is not, which is the
// whole reason to write them down: bubble's `<c:f>` carries no indentation where every other
// numeric-reference block has four spaces, and scatter's `<c:dLbls>` omits the trailing
// `<c:showLeaderLines>` that the category-axis plots emit. Both are the current bytes. Change
// either deliberately, with a note, or not at all.

const XY = [
	{ name: 'X', labels: ['a', 'b', 'c'], values: [1, 2, 3] },
	{ name: 'Y', labels: ['a', 'b', 'c'], values: [4, 5, 6] },
]

async function chartFor(type, data, opts = {}) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(data, { type, x: 1, y: 1, w: 6, h: 4, ...opts })
	})
	return await chartXml(zip)
}

/** The `<c:xVal>` or `<c:yVal>` block, whole. */
function valBlock(xml, tag) {
	const block = new RegExp(`<${tag}>[\\s\\S]*?</${tag}>`).exec(xml)
	assert(block, `expected a <${tag}> block; got: ` + xml.slice(0, 400))
	return block[0]
}

defineRegressionSuite('Shared chart fragments', [
	{
		name: 'a scatter numeric-reference block indents its formula; the bubble y-block does not',
		fn: async () => {
			const scatter = await chartFor(ChartType.scatter, XY)
			assertIncludes(
				valBlock(scatter, 'c:xVal'),
				'<c:numRef>    <c:f>Sheet1!$A$2:$A$4</c:f>',
				'four spaces before the formula'
			)
			assertIncludes(valBlock(scatter, 'c:yVal'), '<c:numRef>    <c:f>Sheet1!$B$2:$B$4</c:f>', 'and on the y-block')

			const bubble = await chartFor(ChartType.bubble, [XY[0], { ...XY[1], sizes: [10, 20, 30] }])
			assertIncludes(
				valBlock(bubble, 'c:xVal'),
				'<c:numRef>    <c:f>Sheet1!$A$2:$A$4</c:f>',
				'bubble indents its x-block like everything else'
			)
			assertIncludes(
				valBlock(bubble, 'c:yVal'),
				'<c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>',
				'but its y-block never has — inert whitespace, and still the emitted bytes'
			)
		},
	},
	{
		// The Y series is cached against the X series' length, so a caller who supplies fewer Y
		// values than X gets gaps rather than a short cache: `<c:ptCount>` counts the X points and
		// the missing `<c:pt>` entries are simply absent (a sparse idx-keyed cache is valid).
		name: 'a short y series leaves cache gaps rather than shortening the point count',
		fn: async () => {
			const xml = await chartFor(ChartType.scatter, [
				{ name: 'X', labels: ['a', 'b', 'c', 'd'], values: [1, 2, 3, 4] },
				{ name: 'Y', labels: ['a', 'b'], values: [7, 8] },
			])
			const yVal = valBlock(xml, 'c:yVal')
			assertIncludes(yVal, '<c:ptCount val="4"/>', 'counted against the x series')
			assertEqual((yVal.match(/<c:pt idx="/g) || []).length, 2, 'only the two supplied y values are cached')
			assertIncludes(yVal, '<c:pt idx="1"><c:v>8</c:v></c:pt>', 'the last supplied value keeps its own index')
		},
	},
	{
		name: 'a category-axis plot emits showLeaderLines; a scatter plot omits it',
		fn: async () => {
			const line = await chartFor(ChartType.line, [{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
				showValue: true,
				showLeaderLines: true,
			})
			assertIncludes(line, '<c:showLeaderLines val="1"/>', 'cat-axis plots carry it')

			const scatter = await chartFor(ChartType.scatter, XY, { showValue: true, showLeaderLines: true })
			assertNotIncludes(
				scatter,
				'<c:showLeaderLines',
				'scatter never has — it has no <c:dLblPos> layout that moves a label away from its point'
			)
			// The rest of the block is the same on both, which is why they share a builder --
			// including the indentation, which is emitted and therefore part of the pin.
			for (const xml of [line, scatter]) {
				assertIncludes(
					xml,
					'    <c:showLegendKey val="0"/>    <c:showVal val="1"/>    <c:showCatName val="0"/>',
					'the flag run, in order, with its indentation'
				)
				assertIncludes(xml, '    <c:showSerName val="0"/>    <c:showPercent val="0"/>    <c:showBubbleSize val="0"/>')
			}
		},
	},
	{
		// The category-axis half of the time-unit validation. The series-axis half is pinned in
		// chart-bar3d-series-axis.test.js; this one had nothing.
		name: 'category-axis time units are lowercased, and an unrecognized one is dropped with a warning',
		fn: async () => {
			const ok = await chartFor(ChartType.line, [{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
				catLabelFormatCode: 'yyyy-mm',
				catAxisBaseTimeUnit: 'Days',
				catAxisMajorTimeUnit: 'MONTHS',
				catAxisMinorTimeUnit: 'years',
			})
			assertIncludes(ok, '<c:baseTimeUnit val="days"/>', 'lowercased on the way out')
			assertIncludes(ok, '<c:majorTimeUnit val="months"/>', 'and so is a shouted one')
			assertIncludes(ok, '<c:minorTimeUnit val="years"/>')

			const diag = await captureDiagnostics(async () => {
				const bad = await chartFor(ChartType.line, [{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
					catLabelFormatCode: 'yyyy',
					catAxisBaseTimeUnit: 'days',
					catAxisMajorTimeUnit: 'fortnights',
				})
				assertIncludes(bad, '<c:baseTimeUnit val="days"/>', 'the valid sibling still emits')
				assertNotIncludes(bad, '<c:majorTimeUnit', 'the whole element is omitted, not emitted with a bad value')
			})
			assert(
				diag.codes.includes('chart/invalid-axis-time-unit'),
				'expected the validation warning; got: ' + JSON.stringify(diag.codes)
			)
		},
	},
	{
		// Every plot builder resolves the palette the same way, through one function, and so
		// does `addChartDefinition`: an empty array names no colours, which is what omitting
		// the option means, so both land on the built-in default for the chart's own type.
		//
		// The pie case is the one that used to differ. `Array.isArray([])` is true, so an
		// explicit `[]` survived normalization untouched and met the plot builders' fallback
		// instead, which was the *bar* palette on every type.
		name: 'an empty chartColors means the same as omitting it',
		fn: async () => {
			const rows = Array.from({ length: 3 }, (_unused, idx) => ({
				name: `S${idx}`,
				labels: ['a', 'b'],
				values: [idx + 1, idx + 2],
			}))
			for (const type of [ChartType.bar, ChartType.line, ChartType.stock, ChartType.surface]) {
				const xml = await chartFor(type, rows)
				assert(/<a:srgbClr val="[0-9A-F]{6}"\/>/.test(xml), `${type}: a colour reaches the part`)
				assertEqual(await chartFor(type, rows, { chartColors: [] }), xml, `${type}: empty is the same as omitted`)
			}

			// The type whose default is not the bar palette, so it is the one that can tell an
			// empty array apart from an omitted one.
			const pie = await chartFor(ChartType.pie, rows)
			assertIncludes(pie, '<a:srgbClr val="5DA5DA"/>', 'omitting chartColors gives a pie the pie palette')
			const emptyPie = await chartFor(ChartType.pie, rows, { chartColors: [] })
			assertIncludes(emptyPie, '<a:srgbClr val="5DA5DA"/>', 'and so does an explicitly empty one')
			assertNotIncludes(emptyPie, '<a:srgbClr val="C0504D"/>', 'not the bar palette it used to fall back to')
			assertEqual(emptyPie, pie, 'empty is byte-identical to omitted on a pie')
		},
	},
])
