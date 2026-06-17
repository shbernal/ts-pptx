import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'
import { defineRegressionSuite, assert } from '../helpers.js'

async function buildSlideXml(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// `textWarp` maps to <a:prstTxWarp prst=".."> inside <a:bodyPr>. Per
// CT_TextBodyProperties the warp child must come BEFORE the autofit group
// (spAutoFit/normAutofit), so guard both presence and ordering.
defineRegressionSuite('Preset text warp (prstTxWarp)', 'usages-transverses-textArchUp', [
	{
		name: 'textWarp emits <a:prstTxWarp> with the requested preset',
		fn: async () => {
			const pres = new PptxGenJS()
			pres.addSlide().addText('R&D', { x: 1, y: 1, w: 4, h: 1, textWarp: 'textArchUp' })

			const xml = await buildSlideXml(pres)
			assert(
				xml.includes('<a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>'),
				`expected prstTxWarp textArchUp; got: ${xml}`
			)
		},
	},
	{
		name: 'prstTxWarp precedes the autofit element when both are set',
		fn: async () => {
			const pres = new PptxGenJS()
			pres.addSlide().addText('R&D', { x: 1, y: 1, w: 4, h: 1, textWarp: 'textArchUp', fit: 'resize' })

			const xml = await buildSlideXml(pres)
			const warpAt = xml.indexOf('<a:prstTxWarp')
			const fitAt = xml.indexOf('<a:spAutoFit')
			assert(warpAt !== -1 && fitAt !== -1, `expected both prstTxWarp and spAutoFit; got: ${xml}`)
			assert(warpAt < fitAt, `expected prstTxWarp (${warpAt}) before spAutoFit (${fitAt})`)
		},
	},
	{
		name: 'no textWarp → no prstTxWarp emitted',
		fn: async () => {
			const pres = new PptxGenJS()
			pres.addSlide().addText('plain', { x: 1, y: 1, w: 4, h: 1 })

			const xml = await buildSlideXml(pres)
			assert(!xml.includes('prstTxWarp'), `expected no prstTxWarp; got: ${xml}`)
		},
	},
])
