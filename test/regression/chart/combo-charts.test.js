import TsPptx, { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert, assertEqual, assertRejects } from '../../helpers.js'

defineRegressionSuite('Combo chart axes [legacy bug-06]', [
	{
		name: 'combo chart with secondary*Axis flags emits all referenced axIds as defs',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				const data = [{ name: 'a', labels: ['x', 'y'], values: [1, 2] }]
				const data2 = [{ name: 'b', labels: ['x', 'y'], values: [10, 20] }]
				s.addChart(
					[
						{ type: ChartType.bar, data: data, options: {} },
						{ type: ChartType.line, data: data2, options: { secondaryValAxis: true, secondaryCatAxis: true } },
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
			const axIdRefs = xml.match(/<c:axId\s+val="(\d+)"\/>/g) || []
			const distinctIds = new Set(axIdRefs.map((t) => t.match(/val="(\d+)"/)[1]))
			const catAx = xml.match(/<c:catAx>/g) || []
			const valAx = xml.match(/<c:valAx>/g) || []
			const definedIds = new Set()
			for (const m of xml.matchAll(/<c:(catAx|valAx)>\s*<c:axId\s+val="(\d+)"\/>/g)) {
				definedIds.add(m[2])
			}
			for (const id of distinctIds) {
				assert(
					definedIds.has(id),
					'axId ' +
						id +
						' referenced but not defined; defs:' +
						[...definedIds].join(',') +
						' refs:' +
						[...distinctIds].join(',')
				)
			}
			assert(
				catAx.length + valAx.length >= 4,
				'expected at least 4 axis defs (primary+secondary), got ' + (catAx.length + valAx.length)
			)
		},
	},
	{
		// A combo's series colours are decided by each series' position across the WHOLE chart, not by
		// its position inside its own subchart. Verified against desktop PowerPoint over COM: a
		// three-series clustered-column chart whose third series is switched to a line keeps that
		// series' third palette colour and merely moves it from `Format.Fill` to `Format.Line` --
		// `FullSeriesCollection(3)` reads back the same RGB before and after the type change, while
		// series 1 and 2 keep theirs. The palette does not restart per plot group.
		//
		// Before this, each subchart's own loop index drove the lookup, so the line below took entry 0
		// (`C0504D`, the first bar's colour) instead of entry 2. The default palette's first three
		// entries are C0504D / 4F81BD / 9BBB59, which is what makes the two readings distinguishable.
		name: "a combo's third series takes the third palette colour, not its subchart's first",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					[
						{
							type: ChartType.bar,
							data: [
								{ name: 'bar1', labels: ['x', 'y'], values: [1, 2] },
								{ name: 'bar2', labels: ['x', 'y'], values: [3, 4] },
							],
							options: {},
						},
						{
							type: ChartType.line,
							data: [{ name: 'line1', labels: ['x', 'y'], values: [5, 6] }],
							options: {},
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
			const lineSer = xml.slice(xml.indexOf('<c:lineChart>'))
			assert(
				lineSer.includes('<a:srgbClr val="9BBB59"/>'),
				'the line series should paint with palette entry 2 (9BBB59); got: ' + lineSer.slice(0, 900)
			)
			assert(
				!lineSer.includes('<a:srgbClr val="C0504D"/>'),
				'the line series must not reuse palette entry 0 (C0504D), which is the first bar'
			)
		},
	},
	{
		// `lineDashValues` reads off the same index, and says so ("the series order in the `data`
		// array"). It used to be handed each subchart's own loop index, so a combo applied entry 0 to
		// the first series of EVERY subchart and never reached the entries past the longest one.
		name: 'lineDashValues indexes a combo by overall series position',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					[
						{
							type: ChartType.line,
							data: [{ name: 'a', labels: ['x', 'y'], values: [1, 2] }],
							options: {},
						},
						{
							type: ChartType.line,
							data: [{ name: 'b', labels: ['x', 'y'], values: [3, 4] }],
							options: { secondaryValAxis: true, secondaryCatAxis: true },
						},
					],
					{ x: 1, y: 1, w: 6, h: 3, lineDashValues: ['solid', 'dash'] }
				)
			})
			const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
			const dashes = [...xml.matchAll(/<a:prstDash val="(\w+)"\/>/g)].map((m) => m[1])
			assert(
				dashes.includes('dash'),
				'the second series should take `lineDashValues[1]` ("dash"); got ' + dashes.join(',')
			)
		},
	},
	{
		// A scatter subchart counted its series from 0 inside its own loop, so behind a bar it took
		// the bar's `<c:idx>` and the bar's palette entry. It now counts its Y series by position
		// across the chart less the X row: the bar is series 0, the scatter's X row is not a series,
		// and its Y series is 1.
		name: 'a scatter behind a bar takes the next series index and colour',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ type: ChartType.bar, data: [{ name: 'A', labels: ['x', 'y', 'z'], values: [1, 2, 3] }], options: {} },
						{
							type: ChartType.scatter,
							data: [
								{ name: 'X', values: [10, 20, 30] },
								{ name: 'Y', values: [5, 6, 7] },
							],
							options: { secondaryValAxis: true, secondaryCatAxis: true },
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
			const scatter = xml.slice(xml.indexOf('<c:scatterChart>'), xml.indexOf('</c:scatterChart>'))
			assert(scatter.includes('<c:idx val="1"/><c:order val="1"/>'), 'the scatter series is series 1; got: ' + scatter)
			assert(scatter.includes('<a:srgbClr val="4F81BD"/>'), 'and paints with palette entry 1 (4F81BD)')
			assert(!scatter.includes('<a:srgbClr val="C0504D"/>'), 'not entry 0 (C0504D), which is the bar')
			// Its X and Y are the X row's and the Y series' own columns, after the label column.
			assert(scatter.includes('<c:xVal><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f>'), 'X is column C; got: ' + scatter)
			assert(scatter.includes('<c:yVal><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f>'), 'Y is column D; got: ' + scatter)
		},
	},
	{
		// PowerPoint refuses any combo holding a bubble chart as corrupt (0x80070570), another bubble
		// chart included, whatever its sizes reference; the same bubble chart alone opens.
		name: 'a combo refuses a bubble subchart',
		fn: async () => {
			const bubble = {
				type: ChartType.bubble,
				data: [
					{ name: 'X', values: [10, 20, 30] },
					{ name: 'Y', values: [5, 6, 7], sizes: [1, 2, 3] },
				],
				options: {},
			}
			for (const types of [
				[
					{ type: ChartType.bar, data: [{ name: 'A', labels: ['x', 'y', 'z'], values: [1, 2, 3] }], options: {} },
					bubble,
				],
				[bubble, { ...bubble, type: ChartType.bubble3d }],
			]) {
				const pres = new TsPptx()
				const slide = pres.addSlide()
				const error = await assertRejects(
					() => slide.addChart(types, { x: 1, y: 1, w: 6, h: 3 }),
					/bubble chart cannot be combined/,
					'addChart'
				)
				assertEqual(/** @type {any} */ (error).code, 'chart/bubble-in-combo', 'the refusal carries its code')
			}
		},
	},
])
