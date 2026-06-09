import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const SVG_MARKUP =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" data-marker="svg-source"><circle cx="12" cy="12" r="10"/></svg>'

const IMG_BROKEN_PREFIX = 'iVBORw0KGgoAAAANSUhEUgAAAGQ'

const svgEntry = (zip) => listEntries(zip).find((name) => name.startsWith('ppt/media/') && name.endsWith('.svg'))
const pngEntry = (zip) => listEntries(zip).find((name) => name.startsWith('ppt/media/') && name.endsWith('.png'))

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
		name: 'addImage({ svg }) PNG preview part is not the broken-image icon',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ svg: SVG_MARKUP, x: 1, y: 1, w: 1, h: 1 })
			})
			const entry = pngEntry(zip)
			assert(entry, 'expected a png preview part; got entries: ' + listEntries(zip).join(', '))
			const b64 = await zip.file(entry).async('base64')
			assert(b64.startsWith('iVBORw0KGgo'), 'expected valid PNG magic bytes; got: ' + b64.slice(0, 20))
			assert(!b64.startsWith(IMG_BROKEN_PREFIX), 'expected placeholder PNG, not the broken-image icon')
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
