'use strict'

const { build, readEntry, listEntries, assert } = require('./helpers')

module.exports = [
	{
		name: 'ppt/theme/theme2.xml is present in archive',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const entries = listEntries(zip)
			assert(entries.indexOf('ppt/theme/theme2.xml') !== -1,
				'expected ppt/theme/theme2.xml in archive; got: ' + entries.join(','))
		}
	},
	{
		name: 'notesMaster1.xml.rels references theme2.xml (not theme1)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/notesMasters/_rels/notesMaster1.xml.rels')
			assert(xml.indexOf('Target="../theme/theme2.xml"') !== -1,
				'expected notesMaster rels to point to theme2.xml; got: ' + xml)
			assert(xml.indexOf('theme1.xml') === -1,
				'expected notesMaster rels NOT to reference theme1.xml; got: ' + xml)
		}
	},
	{
		name: '[Content_Types].xml has Override for theme2.xml',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assert(/Override\s+PartName="\/ppt\/theme\/theme2\.xml"\s+ContentType="application\/vnd\.openxmlformats-officedocument\.theme\+xml"/.test(xml),
				'expected Override entry for /ppt/theme/theme2.xml; got: ' + xml)
		}
	},
	{
		name: 'regression - theme1.xml still present and slideMaster rel still resolves',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const entries = listEntries(zip)
			assert(entries.indexOf('ppt/theme/theme1.xml') !== -1,
				'expected ppt/theme/theme1.xml still present; got: ' + entries.join(','))
			const masterRels = await readEntry(zip, 'ppt/slideMasters/_rels/slideMaster1.xml.rels')
			assert(masterRels.indexOf('Target="../theme/theme1.xml"') !== -1,
				'expected slideMaster rels to still point to theme1.xml; got: ' + masterRels)
			const ct = await readEntry(zip, '[Content_Types].xml')
			assert(/Override\s+PartName="\/ppt\/theme\/theme1\.xml"/.test(ct),
				'expected Override entry for theme1.xml still present; got: ' + ct)
		}
	}
]
