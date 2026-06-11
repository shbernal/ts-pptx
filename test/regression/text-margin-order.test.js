import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'
import { defineRegressionSuite, assert, xmlOpeningTags, xmlAttributes } from '../helpers.js'

async function buildSlideXml(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// margin is documented as [Top, Right, Bottom, Left] (CSS clockwise order); table cells and
// slide numbers already map arrays that way. This guards that textboxes agree, i.e. index 0
// lands on tIns and index 3 on lIns (upstream-pr-1248). valToPts(pt) = round(pt * 12700).
defineRegressionSuite('Text box margin array order', 'upstream-pr-1248', [
	{
		name: 'margin [T,R,B,L] maps each value to the correct bodyPr inset',
		fn: async () => {
			const pres = new PptxGenJS()
			// Four distinct values so any transposition (e.g. swapping Top/Left) is caught.
			pres.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1, margin: [4, 8, 12, 16] })

			const xml = await buildSlideXml(pres)
			const bodyPr = xmlOpeningTags(xml, 'a:bodyPr')[0]
			assert(bodyPr, `expected <a:bodyPr> in slide XML; got: ${xml}`)
			const attrs = xmlAttributes(bodyPr)

			assert(attrs.tIns === '50800', `expected tIns=50800 (Top=4pt); got tIns=${attrs.tIns} in ${bodyPr}`)
			assert(attrs.rIns === '101600', `expected rIns=101600 (Right=8pt); got rIns=${attrs.rIns} in ${bodyPr}`)
			assert(attrs.bIns === '152400', `expected bIns=152400 (Bottom=12pt); got bIns=${attrs.bIns} in ${bodyPr}`)
			assert(attrs.lIns === '203200', `expected lIns=203200 (Left=16pt); got lIns=${attrs.lIns} in ${bodyPr}`)
		},
	},
])
