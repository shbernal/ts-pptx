import { build, readEntry, assert } from './helpers.js'

export default [
	{
		name: 'combo chart with secondary*Axis flags emits all referenced axIds as defs',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				const data = [{ name: 'a', labels: ['x', 'y'], values: [1, 2] }]
				const data2 = [{ name: 'b', labels: ['x', 'y'], values: [10, 20] }]
				s.addChart(
					[
						{ type: p.charts.BAR, data: data, options: {} },
						{ type: p.charts.LINE, data: data2, options: { secondaryValAxis: true, secondaryCatAxis: true } },
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
]
