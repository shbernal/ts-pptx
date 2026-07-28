/**
 * `addChart` treats its `data` and `options` arguments as read-only inputs.
 *
 * The series normalization (`labels` widened to `string[][]`, `_dataIndex` assigned) and the
 * option normalization (defaults, range clamps, deletion of invalid keys) used to be applied
 * in place, so the caller's own arrays and option objects changed under them: iterating the
 * same `labels` array after charting it yielded one nested array instead of the strings that
 * were passed in.
 */
import { ChartType } from '../../dist/node.js'
import { expect, vi } from 'vitest'
import { defineRegressionSuite, build, readEntry, listEntries, assert, assertIncludes } from '../helpers.js'

/** Every chart part in the package, in `chart{N}` order. */
function chartPartPaths(zip) {
	return listEntries(zip)
		.filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
		.sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))
}

async function firstChartPart(zip) {
	const [path] = chartPartPaths(zip)
	assert(path, 'no chart part found in package: ' + listEntries(zip).join(', '))
	return readEntry(zip, path)
}

const POS = { x: 1, y: 1, w: 6, h: 3 }

defineRegressionSuite('Chart input immutability', [
	{
		name: 'a series `labels` array is not rewritten to the nested form on the caller',
		fn: async () => {
			const labels = ['A', 'B', 'C']
			const data = [{ name: 'mix', labels, values: [1, 2, 3] }]
			const before = structuredClone(data)

			const { zip } = await build((p) => {
				p.addSlide().addChart(data, { ...POS, type: ChartType.doughnut })
			})

			expect(data).toEqual(before)
			assert(data[0].labels === labels, 'the caller`s labels array should still be the same object')
			assert(!('_dataIndex' in data[0]), '_dataIndex should not be stamped onto the caller`s series')
			// The normalization still reaches the emitter.
			assertIncludes(await firstChartPart(zip), '<c:v>A</c:v>', 'chart part')
		},
	},
	{
		name: 'the caller`s options object gains no defaults, no `_type`, and keeps invalid keys',
		fn: async () => {
			const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})
			try {
				const options = {
					...POS,
					type: ChartType.bar,
					// Invalid values the normalizer deletes on its own copy.
					layout: { x: 5, y: 0.2, w: 0.5, h: 0.5 },
					catGridLine: { size: -1, style: 'dotted' },
					dataLabelPosition: 'outEnd',
					plotArea: { border: { color: 'FF0000' } },
					dataBorder: { color: 'nope' },
				}
				const before = structuredClone(options)

				await build((p) => {
					p.addSlide().addChart([{ name: 'S', labels: ['A', 'B'], values: [1, 2] }], options)
				})

				expect(options).toEqual(before)
				assert(!('_type' in options), '`_type` should not be stamped onto the caller`s options')
				assert(!('chartColors' in options), 'defaults should not be written onto the caller`s options')
				assert(!('barGapWidthPct' in options), 'defaults should not be written onto the caller`s options')
			} finally {
				warnSpy.mockRestore()
			}
		},
	},
	{
		name: 'multi-level `string[][]` labels are neither double-wrapped nor copied onto the caller',
		fn: async () => {
			const data = [
				{
					name: 'S',
					labels: [
						['Q1', 'Q2', 'Q3', 'Q4'],
						['FY25', '', 'FY26', ''],
					],
					values: [1, 2, 3, 4],
				},
			]
			const before = structuredClone(data)

			const { zip } = await build((p) => {
				p.addSlide().addChart(data, { ...POS, type: ChartType.bar, catAxisMultiLevelLabels: true })
			})

			expect(data).toEqual(before)
			const xml = await firstChartPart(zip)
			assertIncludes(xml, '<c:multiLvlStrRef>', 'chart part')
			assertIncludes(xml, '<c:v>FY25</c:v>', 'chart part')
		},
	},
	{
		name: 'a combo chart leaves its ChartMulti[] untouched and still indexes every series',
		fn: async () => {
			const charts = [
				{ type: ChartType.bar, data: [{ name: 'Bar', labels: ['A', 'B'], values: [1, 2] }], options: {} },
				{ type: ChartType.line, data: [{ name: 'Line', labels: ['A', 'B'], values: [3, 4] }], options: {} },
			]
			const before = structuredClone(charts)

			const { zip } = await build((p) => {
				p.addSlide().addChart(charts, { ...POS })
			})

			expect(charts).toEqual(before)
			assert(!('_dataIndex' in charts[1].data[0]), 'subchart series should not be stamped')

			// The subchart series are plotted from `opts._type[i].data`, so the normalized copies
			// have to reach that path: without them the second series emits `<c:idx val="undefined"/>`.
			const xml = await firstChartPart(zip)
			assertIncludes(xml, '<c:barChart>', 'chart part')
			assertIncludes(xml, '<c:lineChart>', 'chart part')
			assertIncludes(xml, '<c:idx val="0"/>', 'chart part')
			assertIncludes(xml, '<c:idx val="1"/>', 'chart part')
			assert(!xml.includes('undefined'), 'chart part should carry no undefined series index: ' + xml.slice(0, 400))
		},
	},
	{
		name: 'reusing one data + options object for two charts emits two identical chart parts',
		fn: async () => {
			const data = [{ name: 'S', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
			const options = { ...POS, type: ChartType.bar }

			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addChart(data, options)
				slide.addChart(data, options)
			})

			const paths = chartPartPaths(zip)
			assert(paths.length === 2, 'expected two chart parts; got: ' + paths.join(', '))
			const [first, second] = await Promise.all(paths.map((path) => readEntry(zip, path)))
			// Second-chart normalization must not see anything the first chart left behind.
			expect(second).toEqual(first)
		},
	},
	{
		name: 'a chart declared on a slide master does not mutate the master`s object descriptors',
		fn: async () => {
			const chart = {
				type: ChartType.bar,
				data: [{ name: 'S', labels: ['A', 'B'], values: [1, 2] }],
				opts: { ...POS },
			}
			const before = structuredClone(chart)

			await build((p) => {
				p.defineSlideMaster({ title: 'CHART_MASTER', objects: [{ chart }] })
				p.addSlide({ masterName: 'CHART_MASTER' })
			})

			expect(chart).toEqual(before)
		},
	},
])
