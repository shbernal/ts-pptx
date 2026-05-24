'use strict'

const { PptxGenJS, build, readEntry, assert } = require('./helpers')

module.exports = [
	{
		name: 'textless addShape emits <p:sp> containing <p:txBody>',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				// no text passed — this is the failing case for #1441
				s.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// The shape must appear and it must contain a <p:txBody>.
			const spMatch = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)
			assert(spMatch, 'expected a <p:sp>...</p:sp> block in slide1.xml; got: ' + xml)
			const sp = spMatch[0]
			assert(sp.indexOf('<p:txBody>') !== -1,
				'expected <p:txBody> inside <p:sp> for textless shape; got: ' + sp)
			// Empty-txBody fallback must produce at least one <a:p> with endParaRPr.
			assert(/<p:txBody>[\s\S]*?<a:p>[\s\S]*?<a:endParaRPr[^>]*\/>[\s\S]*?<\/a:p>[\s\S]*?<\/p:txBody>/.test(sp),
				'expected <p:txBody> to contain at least <a:p><a:endParaRPr/></a:p>; got: ' + sp)
		}
	},
	{
		name: 'textful addShape still emits text run (regression guard)',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' } })
				s.addText('hello world', { shape: p.shapes.RECTANGLE, x: 4, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('<a:t>hello world</a:t>') !== -1,
				'expected text run <a:t>hello world</a:t> to still appear; got: ' + xml)
		}
	}
]
