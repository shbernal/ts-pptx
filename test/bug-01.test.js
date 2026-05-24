'use strict'

const { build, readEntry, assert } = require('./helpers')

module.exports = [
	{
		name: 'mixed-formatting paragraph emits at most one <a:pPr>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText([
					{ text: 'Bold ', options: { bold: true } },
					{ text: 'and regular' }
				], { x: 1, y: 1, w: 6, h: 1, fontSize: 24 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paragraphs = xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			assert(paragraphs.length > 0, 'expected at least one <a:p>')
			for (const p of paragraphs) {
				const pPrCount = (p.match(/<a:pPr[\s>]/g) || []).length
				assert(pPrCount <= 1, 'paragraph has ' + pPrCount + ' <a:pPr> tags but OOXML allows at most 1: ' + p)
			}
		}
	},
	{
		name: 'paragraph with mixed-align runs still emits 1 <a:pPr> per <a:p>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText([
					{ text: 'left', options: { align: 'left' } },
					{ text: 'right', options: { align: 'right' } }
				], { x: 1, y: 1, w: 6, h: 1, fontSize: 18 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const paragraphs = xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []
			for (const p of paragraphs) {
				const pPrCount = (p.match(/<a:pPr[\s>]/g) || []).length
				assert(pPrCount <= 1, 'paragraph has ' + pPrCount + ' <a:pPr> tags: ' + p)
			}
		}
	}
]
