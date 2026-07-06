import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'
import { defineRegressionSuite, assert } from '../helpers.js'

async function buildSlide1(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// PowerPoint-aligned style props: BorderProps.width (points) and BorderProps.transparency
// (0-100), plus ShadowProps.transparency (0-100, the value the PPT UI actually shows) as a
// friendlier alias of the legacy 0.0-1.0 `opacity`.
defineRegressionSuite('PPT-aligned border/shadow props', 'border-shadow-ppt-props', [
	{
		name: 'table cell border `width` emits the correct line w=',
		fn: async () => {
			const pres = new PptxGenJS()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				border: { type: 'solid', width: 2, color: 'FF0000' },
			})
			const xml = await buildSlide1(pres)
			// 2pt -> 25400 EMU
			assert(xml.includes('<a:lnL w="25400"'), 'expected `width:2` to emit w="25400"; got:\n' + xml)
		},
	},
	{
		name: 'table cell border `transparency` emits <a:alpha> inside the line fill',
		fn: async () => {
			const pres = new PptxGenJS()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				border: { type: 'solid', width: 2, color: 'FF0000', transparency: 25 },
			})
			const xml = await buildSlide1(pres)
			const i = xml.indexOf('<a:lnL ')
			const line = xml.substring(i, xml.indexOf('</a:lnL>', i))
			// transparency 25 -> alpha 75000
			assert(line.includes('<a:alpha val="75000"/>'), 'expected transparency:25 -> alpha 75000; got:\n' + line)
		},
	},
	{
		name: 'shadow `transparency` (0-100) is equivalent to the legacy `opacity` (0.0-1.0)',
		fn: async () => {
			const presT = new PptxGenJS()
			presT.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				shadow: { type: 'outer', color: '000000', blur: 3, offset: 2, transparency: 25 },
			})
			const presO = new PptxGenJS()
			presO.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				shadow: { type: 'outer', color: '000000', blur: 3, offset: 2, opacity: 0.75 },
			})

			const xmlT = await buildSlide1(presT)
			const xmlO = await buildSlide1(presO)
			// transparency 25 == opacity 0.75 -> alpha 75000
			assert(xmlT.includes('<a:alpha val="75000"/>'), 'expected transparency:25 -> alpha 75000; got:\n' + xmlT)
			assert(
				xmlO.includes('<a:alpha val="75000"/>'),
				'legacy opacity:0.75 should also emit alpha 75000 (parity check); got:\n' + xmlO
			)
		},
	},
])
