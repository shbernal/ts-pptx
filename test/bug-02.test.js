'use strict'

const { build, readEntry, listEntries, assert } = require('./helpers')

module.exports = [
	{
		name: '[Content_Types].xml emits exactly one slideMaster Override per existing master part',
		fn: async () => {
			const { zip } = await build(p => {
				for (let i = 0; i < 5; i++) {
					const s = p.addSlide()
					s.addText('Slide ' + (i + 1), { x: 1, y: 1, w: 6, h: 1 })
				}
			})
			const ct = await readEntry(zip, '[Content_Types].xml')
			const overrideMatches = ct.match(/<Override\s+PartName="\/ppt\/slideMasters\/[^"]+"/g) || []
			const masterFiles = listEntries(zip).filter(f => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f))
			assert(overrideMatches.length === masterFiles.length,
				'mismatch: Content_Types has ' + overrideMatches.length + ' slideMaster Overrides but archive has ' + masterFiles.length + ' master parts')
			for (const m of overrideMatches) {
				const target = m.match(/PartName="([^"]+)"/)[1].replace(/^\//, '')
				assert(masterFiles.includes(target),
					'phantom Override: ' + target + ' is in Content_Types but not in archive')
			}
		}
	}
]
