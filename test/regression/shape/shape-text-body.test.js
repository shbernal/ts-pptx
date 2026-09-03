import { ShapeType } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

defineRegressionSuite('Shape text bodies [legacy bug-13]', [
	{
		name: 'textless addShape emits <p:sp> containing <p:txBody>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// no text passed — this is the failing case
				s.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// The shape must appear and it must contain a <p:txBody>.
			const spMatch = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)
			assert(spMatch, 'expected a <p:sp>...</p:sp> block in slide1.xml; got: ' + xml)
			const sp = spMatch[0]
			assert(sp.indexOf('<p:txBody>') !== -1, 'expected <p:txBody> inside <p:sp> for textless shape; got: ' + sp)
			// Empty-txBody fallback must produce at least one <a:p> with endParaRPr.
			assert(
				/<p:txBody>[\s\S]*?<a:p>[\s\S]*?<a:endParaRPr[^>]*\/>[\s\S]*?<\/a:p>[\s\S]*?<\/p:txBody>/.test(sp),
				'expected <p:txBody> to contain at least <a:p><a:endParaRPr/></a:p>; got: ' + sp
			)
		},
	},
	{
		// `addShape` builds a `_type === text` object and never writes `_bodyProp`, where
		// `addTextDefinition` always does. The body-property builder used to read that absent bag's
		// missing `wrap` as `false` and emit `wrap="none"`, so text inside every autoshape ran off
		// the shape on one line. `square` is PowerPoint's own default and the schema's, and it is
		// what an object nobody made a wrap decision about now gets.
		name: 'a shape with no body properties wraps its text',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1 })
				s.addShape(ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, text: 'a line long enough to need wrapping' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const bodyPrs = xml.match(/<a:bodyPr[^>]*>/g) || []
			assert(bodyPrs.length === 2, 'expected one <a:bodyPr> per shape; got: ' + bodyPrs.join(' | '))
			assert(
				bodyPrs.every((b) => b.includes('wrap="square"')),
				'expected every shape body to wrap; got: ' + bodyPrs.join(' | ')
			)
		},
	},
	{
		// The other half of that default: `wrap` is authorable on a text box, and an explicit
		// `false` still has to reach the attribute — the fix must not have made `square` absolute.
		name: 'an explicit `wrap: false` still turns wrapping off',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('no wrapping here', { x: 1, y: 1, w: 2, h: 1, wrap: false })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('<a:bodyPr wrap="none"'), 'expected wrap="none" for an authored `wrap: false`; got: ' + xml)
		},
	},
	{
		name: 'textful addShape still emits text run (regression guard)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' } })
				s.addText('hello world', { shape: ShapeType.rect, x: 4, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.indexOf('<a:t>hello world</a:t>') !== -1,
				'expected text run <a:t>hello world</a:t> to still appear; got: ' + xml
			)
		},
	},
])
