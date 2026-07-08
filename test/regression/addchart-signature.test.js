import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'
import PptxGenJS from '../../dist/node.js'

const DATA = [{ name: 'Sales', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }]

// The chart-part filename uses a module-global counter that advances across builds in this file,
// so locate the (single) chart part rather than assuming `chart1.xml`.
function chartPart(zip) {
	const path = listEntries(zip).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
	assert(path, 'no chart part found in package: ' + listEntries(zip).join(', '))
	return readEntry(zip, path)
}

// Capture warnings emitted while running `fn` (console.warn is the library's warning sink, see
// src/log.ts). This is the one place that intentionally exercises the deprecated positional
// `addChart(type, data, options)` form, so it swaps out console.warn to keep the deprecation
// notice out of the test-run stderr and asserts on it instead.
async function withCapturedWarnings(fn) {
	const original = console.warn
	const messages = []
	console.warn = (msg) => messages.push(String(msg))
	try {
		await fn()
	} finally {
		console.warn = original
	}
	return messages
}

defineRegressionSuite('addChart signature', [
	{
		name: 'canonical form addChart(data, { type }) emits a bar chart',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, { type: p.ChartType.bar, x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartPart(zip)
			assert(xml.includes('<c:barChart>'), 'expected <c:barChart> from canonical form; got: ' + xml.slice(0, 200))
		},
	},
	{
		name: 'canonical form is byte-identical to the deprecated positional form',
		fn: async () => {
			const canonical = await build((p) => {
				p.addSlide().addChart(DATA, { type: p.ChartType.bar, x: 1, y: 1, w: 6, h: 3 })
			})
			/** @type {Awaited<ReturnType<typeof build>> | undefined} */
			let legacy
			const warnings = await withCapturedWarnings(async () => {
				legacy = await build((p) => {
					p.addSlide().addChart(p.ChartType.bar, DATA, { x: 1, y: 1, w: 6, h: 3 })
				})
			})
			// The deprecated positional form must still emit its one-time deprecation notice. (warnOnce
			// dedupes per process, and this is the suite's only positional call, so it fires here.)
			assert(
				warnings.some((m) => /addChart\(type, data, options\) is deprecated/.test(m)),
				'expected a deprecation warning from the positional form; got: ' + JSON.stringify(warnings)
			)
			assert(legacy, 'positional addChart build did not run')
			const a = await chartPart(canonical.zip)
			const b = await chartPart(legacy.zip)
			assert(a === b, 'canonical and legacy chart XML differ')
		},
	},
	{
		name: 'multi-type (combo) charts are unchanged: addChart(ChartMulti[], options)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ type: p.ChartType.bar, data: DATA, options: {} },
						{
							type: p.ChartType.line,
							data: [{ name: 'B', labels: ['Q1', 'Q2', 'Q3'], values: [1, 2, 3] }],
							options: {},
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			const xml = await chartPart(zip)
			assert(
				xml.includes('<c:barChart>') && xml.includes('<c:lineChart>'),
				'expected combo chart with bar+line; got: ' + xml.slice(0, 200)
			)
		},
	},
	{
		name: 'omitting the chart type on the options object throws',
		fn: async () => {
			const p = new PptxGenJS()
			let threw = false
			try {
				// Negative test: `type` is intentionally omitted; cast past the required-`type` overload.
				p.addSlide().addChart(DATA, /** @type {any} */ ({ x: 1, y: 1, w: 6, h: 3 }))
			} catch (err) {
				threw = true
				assert(/type/.test(String(err && err.message)), 'error should mention the missing `type`; got: ' + err)
			}
			assert(threw, 'expected addChart without a type to throw')
		},
	},
])
