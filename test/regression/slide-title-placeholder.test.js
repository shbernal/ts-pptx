import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// #1298: PowerPoint's accessibility checker reports "Missing Slide Title" unless a
// slide has a shape with a title placeholder (<p:ph type="title"/>). A standalone
// `addText(..., { placeholder: 'title' })` on a blank/default layout must therefore
// still emit a real <p:ph type="title"/> on its slide shape.
defineRegressionSuite('Slide title placeholder', 'upstream-issue-1298', [
	{
		name: 'standalone title text (no matching layout placeholder) emits <p:ph type="title">',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('My Slide Title', { x: 0.5, y: 0.3, w: 9, h: 1, fontSize: 32, placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			const titleSp = spBlocks.find((sp) => /<p:ph[^>]*type="title"/.test(sp))
			assert(titleSp, 'expected a <p:sp> with <p:ph type="title" .../>; got: ' + xml)
			assert(/<a:t>My Slide Title<\/a:t>/.test(titleSp), 'expected populated title text run; got: ' + titleSp)
			// Populated slide-level title must NOT be flagged with hasCustomPrompt (that marks
			// placeholder *definitions* with prompt text; here the visible text is real content).
			assert(
				!/hasCustomPrompt/.test(titleSp),
				'standalone title placeholder must not set hasCustomPrompt; got: ' + titleSp
			)
		},
	},
	{
		name: 'standalone body text emits <p:ph type="body">',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('Body content', { x: 0.5, y: 2, w: 9, h: 4, placeholder: 'body' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(
				spBlocks.some((sp) => /<p:ph[^>]*type="body"/.test(sp)),
				'expected a <p:sp> with <p:ph type="body" .../>; got: ' + xml
			)
		},
	},
	{
		name: 'plain text (no placeholder) emits no <p:ph>',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('Just a text box', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/<p:ph/.test(xml), 'plain text box must not emit <p:ph>; got: ' + xml)
		},
	},
])
