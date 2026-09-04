import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, listEntries, assert, captureDiagnostics } from '../../helpers.js'

const DATA = [{ name: 'Region', labels: ['North', 'South', 'East'], values: [10, 20, 30] }]

function chartPart(zip) {
	const path = listEntries(zip).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
	assert(path, 'no chart part found in package: ' + listEntries(zip).join(', '))
	return readEntry(zip, path)
}

async function radarXml(radarStyle) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(DATA, { type: ChartType.radar, radarStyle, x: 1, y: 1, w: 6, h: 4 })
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
		name: "the wire spelling 'standard' is accepted as itself",
		fn: async () => {
			const xml = await radarXml('standard')
			assert(
				xml.includes('<c:radarStyle val="standard"/>'),
				'expected val="standard" from standard; got: ' + xml.slice(0, 300)
			)
		},
	},
	{
		// The case that made this an alias rather than a typo: `marker` is what the chart part
		// shows, so it is the obvious guess, and it used to warn and fall back to a plain radar.
		name: "the wire spelling 'marker' is accepted as itself",
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(() => radarXml('marker'))
			assert(
				xml.includes('<c:radarStyle val="marker"/>'),
				'expected val="marker" from marker; got: ' + xml.slice(0, 300)
			)
			assert(codes.length === 0, 'a wire spelling is not a diagnostic; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'an unknown radarStyle warns and falls back to standard',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(() => radarXml('spider'))
			assert(!xml.includes('spider'), 'the bad value must not reach the part; got: ' + xml.slice(0, 300))
			assert(
				xml.includes('<c:radarStyle val="standard"/>'),
				'expected the standard fallback; got: ' + xml.slice(0, 300)
			)
			assert(codes.includes('chart/invalid-option-value'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'default (no radarStyle) falls back to standard',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, { type: ChartType.radar, x: 1, y: 1, w: 6, h: 4 })
			})
			const xml = await chartPart(zip)
			assert(
				xml.includes('<c:radarStyle val="standard"/>'),
				'expected default val="standard"; got: ' + xml.slice(0, 300)
			)
		},
	},
])
