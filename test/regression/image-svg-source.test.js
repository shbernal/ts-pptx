import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const SVG_MARKUP =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" data-marker="svg-source"><circle cx="12" cy="12" r="10"/></svg>'

const svgEntry = (zip) => listEntries(zip).find((name) => name.startsWith('ppt/media/') && name.endsWith('.svg'))

defineRegressionSuite('Image svg source', [
	{
		name: 'addImage({ svg }) embeds the markup as an svg media part',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ svg: SVG_MARKUP, x: 1, y: 1, w: 1, h: 1 })
			})
			const entry = svgEntry(zip)
			assert(entry, 'expected an svg media part; got entries: ' + listEntries(zip).join(', '))
			const svg = await readEntry(zip, entry)
			assert(svg.includes('data-marker="svg-source"'), 'expected original svg markup; got: ' + svg)
		},
	},
	{
		name: 'addImage({ svg }) is treated as an svg (consumes a png preview rel too)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ svg: SVG_MARKUP, x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/asvg:svgBlip/.test(xml), 'expected svgBlip referencing the svg part; got: ' + xml)
		},
	},
	{
		name: 'data wins over svg when both are supplied',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, svg: SVG_MARKUP, x: 1, y: 1, w: 1, h: 1 })
			})
			assert(!svgEntry(zip), 'svg should be ignored when data is provided; got an svg part')
		},
	},
])
