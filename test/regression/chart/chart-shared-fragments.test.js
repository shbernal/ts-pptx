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
// One case below pins something that looks like an accident and is not, which is the whole
// reason to write it down: scatter's `<c:dLbls>` omits the trailing `<c:showLeaderLines>` that
// the category-axis plots emit. That is the current bytes. Change it deliberately, with a note,
// or not at all.
//
// A second such pin used to live here — bubble's `<c:f>` carrying no indentation where every
// other numeric-reference block had four spaces. The chart emitters are flat now
// (docs/chart-whitespace-flatten.md), so that difference no longer exists to pin.

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
		name: 'every numeric-reference block emits the same shape, scatter and bubble alike',
		fn: async () => {
			// This case used to pin four different indentations, one of which was bubble's y-block
			// having none where the other three had four spaces. All four are flat now
			// (docs/chart-whitespace-flatten.md), so what is left to pin is the part that was
			// always the point: one builder, one shape, and the right formula in each block.
			const scatter = await chartFor(ChartType.scatter, XY)
			assertIncludes(valBlock(scatter, 'c:xVal'), '<c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>', 'scatter x-block')
			assertIncludes(valBlock(scatter, 'c:yVal'), '<c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>', 'scatter y-block')

			const bubble = await chartFor(ChartType.bubble, [XY[0], { ...XY[1], sizes: [10, 20, 30] }])
			assertIncludes(valBlock(bubble, 'c:xVal'), '<c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>', 'bubble x-block')
			assertIncludes(
				valBlock(bubble, 'c:yVal'),
				'<c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>',
				'bubble y-block, no longer the odd one out'
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
			// The rest of the block is the same on both, which is why they share a builder. The
			// run is pinned as one contiguous string because the ORDER is `CT_DLbls`'s and is not
			// negotiable: a flag in the wrong place is a repair prompt, not a wrong-looking chart.
			// (It used to pin the indentation too; that is gone, see
			// docs/chart-whitespace-flatten.md.)
			for (const xml of [line, scatter]) {
				assertIncludes(
					xml,
					'<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>',
					'the flag run, in schema order'
				)
				assertIncludes(xml, '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>')
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
	{
		// Pie was the last family building its own `<c:val>` instead of going through the shared
		// numeric-reference block, and it disagreed about a gap rather than merely spelling one
		// differently: a missing slice value kept its `<c:pt>` with an empty `<c:v>`, and an
		// infinity reached the deck as `<c:v>Infinity</c:v>`, which PowerPoint refuses to open
		// (0x80070570). Nothing covered a gap in a pie series, which is how it survived.
		//
		// The schema oracle cannot stand in for this test: `<c:v>` is `s:ST_Xstring`, so the
		// OpenXmlValidator passes the corrupt spelling as readily as the correct one.
		name: 'a pie caches a gap the same way every other family does',
		fn: async () => {
			const gap = await chartFor(ChartType.pie, [
				{ name: 'Status', labels: ['Red', 'Amber', 'Green', 'Unknown'], values: [10, null, 38, 2] },
			])
			const val = valBlock(gap, 'c:val')
			assertIncludes(val, '<c:ptCount val="4"/>', 'the range still spans every slice')
			assertNotIncludes(val, '<c:v></c:v>', 'the gap leaves the point out rather than emptying it')
			assertEqual((val.match(/<c:pt idx="/g) || []).length, 3, 'only the three supplied values are cached')
			assertIncludes(val, '<c:pt idx="2"><c:v>38</c:v></c:pt>', 'the slices after the gap keep their own indices')

			// A doughnut shares the builder, so it inherits the same treatment.
			const doughnut = await chartFor(ChartType.doughnut, [
				{ name: 'Status', labels: ['a', 'b', 'c'], values: [1, undefined, 3] },
			])
			assertNotIncludes(valBlock(doughnut, 'c:val'), '<c:v></c:v>', 'and so does a doughnut')

			// The half that was a repair prompt, not a cosmetic difference.
			const diag = await captureDiagnostics(async () => {
				const nonFinite = await chartFor(ChartType.pie, [
					{ name: 'Status', labels: ['a', 'b', 'c'], values: [1, Infinity, 3] },
				])
				assertNotIncludes(nonFinite, '<c:v>Infinity</c:v>', 'an infinity never reaches the cache')
				assertEqual(
					(valBlock(nonFinite, 'c:val').match(/<c:pt idx="/g) || []).length,
					2,
					'it is dropped like any other non-finite value'
				)
			})
			assert(
				diag.codes.includes('chart/non-finite-value'),
				'expected the non-finite warning; got: ' + JSON.stringify(diag.codes)
			)
		},
	},
])
