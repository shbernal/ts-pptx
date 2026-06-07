import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertContentTypeOverride,
	assertIncludes,
	assertNotIncludes,
} from '../helpers.js'

defineRegressionSuite('Theme relationships', 'legacy bug-15', [
	{
		name: 'ppt/theme/theme2.xml is present in archive',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const entries = listEntries(zip)
			assert(
				entries.indexOf('ppt/theme/theme2.xml') !== -1,
				'expected ppt/theme/theme2.xml in archive; got: ' + entries.join(',')
			)
		},
	},
	{
		name: 'notesMaster1.xml.rels references theme2.xml (not theme1)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/notesMasters/_rels/notesMaster1.xml.rels')
			assertIncludes(xml, 'Target="../theme/theme2.xml"', 'notesMaster rels')
			assertNotIncludes(xml, 'theme1.xml', 'notesMaster rels')
		},
	},
	{
		name: '[Content_Types].xml has Override for theme2.xml',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeOverride(xml, '/ppt/theme/theme2.xml')
		},
	},
	{
		name: 'regression - theme1.xml still present and slideMaster rel still resolves',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const entries = listEntries(zip)
			assert(
				entries.indexOf('ppt/theme/theme1.xml') !== -1,
				'expected ppt/theme/theme1.xml still present; got: ' + entries.join(',')
			)
			const masterRels = await readEntry(zip, 'ppt/slideMasters/_rels/slideMaster1.xml.rels')
			assertIncludes(masterRels, 'Target="../theme/theme1.xml"', 'slideMaster rels')
			const ct = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeOverride(ct, '/ppt/theme/theme1.xml')
		},
	},
])
