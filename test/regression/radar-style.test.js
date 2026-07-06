import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

const DATA = [{ name: 'Region', labels: ['North', 'South', 'East'], values: [10, 20, 30] }]

function chartPart(zip) {
	const path = listEntries(zip).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
	assert(path, 'no chart part found in package: ' + listEntries(zip).join(', '))
	return readEntry(zip, path)
}

// Capture warnings emitted while running `fn` (console.warn is the library's warning sink, see src/log.ts).
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

async function radarXml(radarStyle) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(DATA, { type: p.ChartType.radar, radarStyle, x: 1, y: 1, w: 6, h: 4 })
	})
	return chartPart(zip)
}

defineRegressionSuite('radarStyle values', [
	{
		name: 'canonical \'radar\' emits ST_RadarStyle val="standard"',
		fn: async () => {
			const xml = await radarXml('radar')
			assert(
				xml.includes('<c:radarStyle val="standard"/>'),
				'expected val="standard" from radar; got: ' + xml.slice(0, 300)
			)
		},
	},
	{
		name: 'canonical \'markers\' emits ST_RadarStyle val="marker"',
		fn: async () => {
			const xml = await radarXml('markers')
			assert(
				xml.includes('<c:radarStyle val="marker"/>'),
				'expected val="marker" from markers; got: ' + xml.slice(0, 300)
			)
		},
	},
	{
		name: '\'filled\' emits ST_RadarStyle val="filled"',
		fn: async () => {
			const xml = await radarXml('filled')
			assert(xml.includes('<c:radarStyle val="filled"/>'), 'expected val="filled"; got: ' + xml.slice(0, 300))
		},
	},
	{
		name: 'default (no radarStyle) falls back to standard',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, { type: p.ChartType.radar, x: 1, y: 1, w: 6, h: 4 })
			})
			const xml = await chartPart(zip)
			assert(
				xml.includes('<c:radarStyle val="standard"/>'),
				'expected default val="standard"; got: ' + xml.slice(0, 300)
			)
		},
	},
	{
		name: 'deprecated \'standard\' still maps to val="standard" and warns',
		fn: async () => {
			let xml = ''
			const warnings = await withCapturedWarnings(async () => {
				xml = await radarXml('standard')
			})
			assert(
				xml.includes('<c:radarStyle val="standard"/>'),
				'expected val="standard" from deprecated standard; got: ' + xml.slice(0, 300)
			)
			assert(
				warnings.some((m) => /radarStyle: 'standard' is deprecated/.test(m)),
				'expected a deprecation warning for standard; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		name: 'deprecated \'marker\' still maps to val="marker" and warns',
		fn: async () => {
			let xml = ''
			const warnings = await withCapturedWarnings(async () => {
				xml = await radarXml('marker')
			})
			assert(
				xml.includes('<c:radarStyle val="marker"/>'),
				'expected val="marker" from deprecated marker; got: ' + xml.slice(0, 300)
			)
			assert(
				warnings.some((m) => /radarStyle: 'marker' is deprecated/.test(m)),
				'expected a deprecation warning for marker; got: ' + JSON.stringify(warnings)
			)
		},
	},
])
