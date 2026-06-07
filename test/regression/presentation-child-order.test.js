import { defineRegressionSuite, build, readEntry, assertXmlOrder } from '../helpers.js'

defineRegressionSuite('Presentation child order', 'legacy bug-20', [
	{
		name: 'ppt/presentation.xml: notesMasterIdLst appears AFTER sldMasterIdLst',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			assertXmlOrder(xml, '<p:sldMasterIdLst>', '<p:notesMasterIdLst>', 'presentation.xml')
		},
	},
	{
		name: 'ppt/presentation.xml: notesMasterIdLst appears BEFORE sldIdLst (CT_Presentation order)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			assertXmlOrder(xml, '<p:notesMasterIdLst>', '<p:sldIdLst>', 'presentation.xml')
		},
	},
	{
		name: 'ppt/presentation.xml: sldIdLst appears BEFORE sldSz',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			assertXmlOrder(xml, '<p:sldIdLst>', '<p:sldSz ', 'presentation.xml')
		},
	},
	{
		name: 'ppt/presentation.xml: sldSz appears BEFORE notesSz',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			assertXmlOrder(xml, '<p:sldSz ', '<p:notesSz ', 'presentation.xml')
		},
	},
])
