import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

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
		// Regression: the same SVG *file* placed 2+ times on one slide. Each placement pushes its
		// own png-fallback rel, but only the first (path-unique) rel was converted to the PNG
		// placeholder preview; the path-duplicate fallbacks kept the raw SVG bytes in a `.png` part,
		// producing a content-type/magic mismatch that made PowerPoint repair (and drop) the deck.
		name: 'same SVG file used twice never leaves SVG bytes in a .png fallback part',
		fn: async () => {
			const dir = mkdtempSync(join(tmpdir(), 'pptx-svg-dupe-'))
			const svgPath = join(dir, 'mark.svg')
			writeFileSync(svgPath, SVG_MARKUP)
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ path: svgPath, x: 1, y: 1, w: 1, h: 1 })
				s.addImage({ path: svgPath, x: 3, y: 1, w: 1, h: 1 })
			})
			const pngParts = listEntries(zip).filter((n) => n.startsWith('ppt/media/') && n.endsWith('.png'))
			assert(pngParts.length > 0, 'expected png preview parts; got: ' + listEntries(zip).join(', '))
			for (const part of pngParts) {
				const b64 = await zip.file(part).async('base64')
				assert(
					b64.startsWith('iVBORw0KGgo'),
					`png fallback ${part} lacks PNG magic bytes (holds SVG?); got: ${b64.slice(0, 24)}`
				)
				const text = await readEntry(zip, part)
				assert(!/<svg|<\?xml/i.test(text), `png fallback ${part} contains SVG markup`)
			}
			// Both svgBlip pictures must survive.
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert((xml.match(/asvg:svgBlip/g) || []).length === 2, 'expected two svgBlip pictures; got: ' + xml)
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
