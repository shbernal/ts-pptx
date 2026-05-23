'use strict'

const { build, readEntry, assert } = require('./helpers')

module.exports = [
	{
		name: 'B12: solid-color slide.background <p:bgPr> contains <a:effectLst/>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.background = { color: '0088CC' }
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:bg><p:bgPr><a:solidFill>[\s\S]*?<\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg>/.test(xml),
				'expected <p:bgPr> to contain <a:solidFill>...</a:solidFill><a:effectLst/>; got: ' + xml)
		}
	},
	{
		name: 'B12: image-background still emits <a:effectLst/> (regression guard)',
		fn: async () => {
			// 1x1 transparent PNG
			const b64png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.background = { data: 'image/png;base64,' + b64png }
				s.addText('hi', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:bg><p:bgPr>[\s\S]*<a:effectLst\/><\/p:bgPr><\/p:bg>/.test(xml),
				'expected image bgPr to keep <a:effectLst/>; got: ' + xml)
		}
	}
]
