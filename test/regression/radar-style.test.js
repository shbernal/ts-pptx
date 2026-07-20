import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

const DATA = [{ name: 'Region', labels: ['North', 'South', 'East'], values: [10, 20, 30] }]

function chartPart(zip) {
	const path = listEntries(zip).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
	assert(path, 'no chart part found in package: ' + listEntries(zip).join(', '))
	return readEntry(zip, path)
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
])
