import { defineRegressionSuite, build, readEntry, assert, assertIncludes, assertNotIncludes } from '../helpers.js'

// gitbrent/PptxGenJS#1159: slide numbers defined on a master disappeared on slides that
// PowerPoint inserts from that master. Root cause: makeXmlMaster always emitted
// <p:hf sldNum="0" .../>, and CT_HeaderFooter/@sldNum defaults to true (ECMA-376), so the
// explicit "0" disabled the slide-number placeholder for inherited/new slides even though the
// master itself carried a sldNum placeholder shape. This suite pins the emitted master XML to be
// internally consistent: when a slide number is defined, the placeholder is present AND the
// header/footer element does not disable sldNum.
defineRegressionSuite('Master slide-number header/footer', [
	{
		name: 'master with slideNumber → sldNum placeholder present and <p:hf> does not disable sldNum',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'MASTER_WITH_SLDNUM',
					slideNumber: { x: 0.5, y: '90%' },
				})
			})
			const xml = await readEntry(zip, 'ppt/slideMasters/slideMaster1.xml')
			assertIncludes(xml, '<p:ph type="sldNum"', 'slide-number placeholder')
			const hf = (xml.match(/<p:hf\b[^>]*\/>/) || [])[0]
			assert(hf, 'expected a <p:hf .../> element; got: ' + xml)
			assertNotIncludes(hf, 'sldNum="0"', '<p:hf> must not disable sldNum when a slide number is defined')
		},
	},
	{
		name: 'master without slideNumber → <p:hf> keeps sldNum="0" (no inherited slide number)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'MASTER_NO_SLDNUM' })
			})
			const xml = await readEntry(zip, 'ppt/slideMasters/slideMaster1.xml')
			const hf = (xml.match(/<p:hf\b[^>]*\/>/) || [])[0]
			assert(hf, 'expected a <p:hf .../> element; got: ' + xml)
			assertIncludes(hf, 'sldNum="0"', '<p:hf> should disable sldNum when no slide number is defined')
		},
	},
])
