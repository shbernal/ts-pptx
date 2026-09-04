import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

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
])
