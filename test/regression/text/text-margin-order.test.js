import JSZip from 'jszip'
import TsPptx from '../../../dist/node.js'
import { defineRegressionSuite, assert, xmlOpeningTags, xmlAttributes } from '../../helpers.js'

async function buildSlideXml(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// margin is documented as [Top, Right, Bottom, Left] (CSS clockwise order); table cells and
// slide numbers already map arrays that way. This guards that textboxes agree, i.e. index 0
// lands on tIns and index 3 on lIns (upstream-pr-1248). Margins are inches (marginToEmu),
// so inch2Emu(in) = round(in * 914400).
defineRegressionSuite('Text box margin array order', 'upstream-pr-1248', [
	{
		name: 'margin [T,R,B,L] maps each value to the correct bodyPr inset',
		fn: async () => {
			const pres = new TsPptx()
			// Four distinct values (inches, all < 1 to avoid the legacy-points warning) so any
			// transposition (e.g. swapping Top/Left) is caught.
			pres.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1, margin: [0.1, 0.2, 0.3, 0.4] })

			const xml = await buildSlideXml(pres)
			const bodyPr = xmlOpeningTags(xml, 'a:bodyPr')[0]
			assert(bodyPr, `expected <a:bodyPr> in slide XML; got: ${xml}`)
			const attrs = xmlAttributes(bodyPr)

			assert(attrs.tIns === '91440', `expected tIns=91440 (Top=0.1in); got tIns=${attrs.tIns} in ${bodyPr}`)
			assert(attrs.rIns === '182880', `expected rIns=182880 (Right=0.2in); got rIns=${attrs.rIns} in ${bodyPr}`)
			assert(attrs.bIns === '274320', `expected bIns=274320 (Bottom=0.3in); got bIns=${attrs.bIns} in ${bodyPr}`)
			assert(attrs.lIns === '365760', `expected lIns=365760 (Left=0.4in); got lIns=${attrs.lIns} in ${bodyPr}`)
		},
	},
])
