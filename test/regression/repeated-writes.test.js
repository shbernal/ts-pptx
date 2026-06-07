import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'
import { defineRegressionSuite, assert } from '../helpers.js'

async function buildOnce(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

defineRegressionSuite('Repeated presentation writes', 'legacy bug-04', [
	{
		name: 'two writes on same Presentation produce identical slide1.xml (text/shape branch)',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 4, h: 2, shadow })

			const xml1 = await buildOnce(pres)
			const xml2 = await buildOnce(pres)
			assert(
				xml1 === xml2,
				'expected slide1.xml byte-equal across two writes; differ.\nfirst:\n' + xml1 + '\nsecond:\n' + xml2
			)
		},
	},
	{
		name: 'user shadow object is not mutated across two writes',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 4, h: 2, shadow })

			await buildOnce(pres)
			await buildOnce(pres)

			assert(shadow.blur === 6, 'expected shadow.blur to remain 6 (pt); got ' + shadow.blur)
			assert(shadow.offset === 2, 'expected shadow.offset to remain 2 (pt); got ' + shadow.offset)
			assert(shadow.opacity === 0.15, 'expected shadow.opacity to remain 0.15; got ' + shadow.opacity)
			assert(shadow.angle === undefined, 'expected shadow.angle to remain undefined; got ' + shadow.angle)
			assert(shadow.color === '000000', 'expected shadow.color to remain "000000"; got ' + shadow.color)
		},
	},
	{
		name: 'two writes on same Presentation produce identical slide1.xml (image branch)',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			// 1x1 transparent PNG, base64
			const png =
				'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
			slide.addImage({ data: png, x: 1, y: 1, w: 1, h: 1, shadow })

			const xml1 = await buildOnce(pres)
			const xml2 = await buildOnce(pres)
			assert(
				xml1 === xml2,
				'expected image-branch slide1.xml byte-equal across two writes; differ.\nfirst:\n' + xml1 + '\nsecond:\n' + xml2
			)

			assert(shadow.blur === 6, 'expected image-branch shadow.blur to remain 6 (pt); got ' + shadow.blur)
			assert(shadow.opacity === 0.15, 'expected image-branch shadow.opacity to remain 0.15; got ' + shadow.opacity)
		},
	},
	{
		name: 'regression - single write still emits sane shadow EMU values',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			slide.addShape(pres.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 4,
				h: 2,
				shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 },
			})
			const xml = await buildOnce(pres)
			assert(xml.indexOf('<a:effectLst>') !== -1, 'expected <a:effectLst> in single-write output')
			// blurRad=valToPts(6)=76200; dist=valToPts(2)=25400; dir=Math.round(270*60000)=16200000; alpha=Math.round(0.15*100000)=15000
			assert(xml.indexOf('blurRad="76200"') !== -1, 'expected blurRad="76200"; got: ' + xml)
			assert(xml.indexOf('dist="25400"') !== -1, 'expected dist="25400"; got: ' + xml)
			assert(xml.indexOf('dir="16200000"') !== -1, 'expected dir="16200000"; got: ' + xml)
			assert(xml.indexOf('<a:alpha val="15000"/>') !== -1, 'expected <a:alpha val="15000"/>; got: ' + xml)
			// Belt-and-braces: no absurdly large EMU values that signal double-conversion
			assert(xml.indexOf('blurRad="967740000"') === -1, 'unexpected double-converted blurRad in output: ' + xml)
		},
	},
])
