// `addImage({ line })` — the picture border.
//
// The option is documented on `ImageBaseProps` with two worked examples, and the picture emitter
// reads it. What sat between them was the definer's `pickDefined` allow-list, which named every
// other option the emitter reads and not this one, so the option was dropped on the way in and the
// picture came out with no border at all. Nothing said so: a dropped option is indistinguishable
// from one never passed.
//
// The read side was never the problem — a picture's `p:spPr/a:ln` has always read back through
// `lineColor`/`lineWidthPt`/`lineDash` — so the last case here is the round trip that proves the
// two halves now meet.

import { PNG_1X1, defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'
import { Presentation } from '../../../dist/read.js'

/** The `<a:ln …>` opening tag inside the slide's one `p:pic`, or `null`. */
function pictureLine(xml) {
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(xml)
	return pic ? (/<a:ln[^>]*>/.exec(pic[0])?.[0] ?? null) : null
}

defineRegressionSuite('Image border', [
	{
		name: 'addImage({ line }) writes the outline the emitter has always known how to draw',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addImage({
					data: PNG_1X1,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					line: { color: '0088CC', width: 2, dashType: 'dash' },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const ln = pictureLine(xml)
			assert(ln, 'the picture carries an a:ln; got: ' + xml.slice(0, 600))
			// 2pt is 25400 EMU.
			assert(ln.includes('w="25400"'), `the stated width reaches @w; got ${ln}`)
			const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(xml)[0]
			assert(pic.includes('<a:srgbClr val="0088CC"/>'), 'the stated colour reaches the line fill')
			assert(pic.includes('<a:prstDash val="dash"/>'), 'and so does the dash')
		},
	},
	{
		name: 'a picture with no line option carries no a:ln',
		fn: async () => {
			// The guard: it is easy to fix a dropped option by always emitting the element.
			const { zip } = await build((p) => {
				p.addSlide().addImage({ data: PNG_1X1, x: 1, y: 1, w: 2, h: 2 })
			})
			assertEqual(pictureLine(await readEntry(zip, 'ppt/slides/slide1.xml')), null, 'no outline is written')
		},
	},
	{
		name: 'the border a picture is written with is the border it reads back',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: PNG_1X1,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					line: { color: '0088CC', width: 2, dashType: 'dash' },
				})
			})
			const pres = await Presentation.load(buf)
			const picture = pres.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
			assert(picture, 'the deck has a picture')
			assertEqual(picture.lineColor, '0088CC', 'lineColor')
			assertEqual(picture.lineWidthPt, 2, 'lineWidthPt')
			assertEqual(picture.lineDash, 'dash', 'lineDash')
		},
	},
])
