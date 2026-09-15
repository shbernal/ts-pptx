// Image geometry: the extent a picture is drawn at, the crop that fits it into a box, the size an
// SVG states, what an unset background leaves behind, and the frame a layout placeholder lends.
//
// `fitSrcRectPercents` and `getImageSizeFromBytes` are imported from `src/`, as the format registry
// suite imports its tables: they are internal, and each case also asserts what they decide through
// the package.

import { describe, test } from 'vitest'
import { fitSrcRectPercents, getImageSizeFromBytes } from '../../../src/media/image-size.ts'
import { Presentation } from '../../../dist/read.js'
import { EMU_PER_INCH } from '../../../dist/node.js'
import { assert, assertEqual, build, captureDiagnostics, readEntry } from '../../helpers.js'

/**
 * A PNG header carrying an arbitrary intrinsic size. The size reader looks at the IHDR dimensions
 * only, so 24 bytes are enough.
 * @param {number} w
 * @param {number} h
 */
function pngBytes(w, h) {
	const b = Buffer.alloc(24)
	b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0)
	b.writeUInt32BE(13, 8)
	b.write('IHDR', 12, 'ascii')
	b.writeUInt32BE(w, 16)
	b.writeUInt32BE(h, 20)
	return b
}
const pngData = (w, h) => 'image/png;base64,' + pngBytes(w, h).toString('base64')

/** The picture's own `<a:off>`/`<a:ext>`, scoped to `<p:pic>` past the slide's group frame. */
function pictureFrame(xml) {
	const pic = xml.split('<p:pic>')[1]
	assert(pic, 'expected a <p:pic>; got: ' + xml)
	const m = /<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/.exec(pic)
	assert(m, 'expected the picture frame; got: ' + pic)
	return { x: Number(m[1]), y: Number(m[2]), cx: Number(m[3]), cy: Number(m[4]) }
}

async function slideXml(author, n = 1) {
	const { zip } = await build(author)
	return readEntry(zip, `ppt/slides/slide${n}.xml`)
}

describe('image geometry', () => {
	test('a data image given one side as a string derives the other from its natural ratio', async () => {
		// It fell to the 1in default, so a square image given `w: '20%'` came out 2in by 1in, while the
		// same image by `path` came out square.
		const wide = pictureFrame(
			await slideXml((p) => p.addSlide().addImage({ data: pngData(32, 32), x: 0, y: 0, w: '20%' }))
		)
		assertEqual(wide.cx, 2 * EMU_PER_INCH, '20% of the 10in slide')
		assertEqual(wide.cy, wide.cx, 'a square image stays square')
		const tall = pictureFrame(
			await slideXml((p) => p.addSlide().addImage({ data: pngData(64, 32), x: 0, y: 0, h: '1in' }))
		)
		assertEqual(tall.cy, EMU_PER_INCH, 'the stated height')
		assertEqual(tall.cx, 2 * EMU_PER_INCH, 'a 2:1 image is twice as wide')
	})

	test('a fit has no answer for a side that is not greater than zero', () => {
		assertEqual(fitSrcRectPercents('cover', { w: 10, h: 10 }, { w: 0, h: 0 }), null, 'a zero box')
		assertEqual(fitSrcRectPercents('contain', { w: 10, h: 10 }, { w: 0, h: 5 }), null, 'a zero-width box')
		assertEqual(fitSrcRectPercents('cover', { w: 0, h: 10 }, { w: 5, h: 5 }), null, 'a zero-width image')
		assertEqual(
			JSON.stringify(fitSrcRectPercents('cover', { w: 10, h: 10 }, { w: 20, h: 10 })),
			JSON.stringify({ l: 0, r: 0, t: 25000, b: 25000 }),
			'a box with extent still fits'
		)
	})

	test('cover and contain on a zero-size box write a plain stretch rather than NaN', async () => {
		for (const type of ['cover', 'contain']) {
			const { result } = await captureDiagnostics(() =>
				slideXml((p) => p.addSlide().addImage({ data: pngData(32, 32), x: 1, y: 1, w: 0, h: 0, sizing: { type } }))
			)
			assert(!result.includes('NaN'), `${type}: no NaN in the slide`)
			assert(!result.includes('<a:srcRect'), `${type}: no crop rectangle`)
			assert(result.includes('<a:fillRect'), `${type}: the fill directive is kept`)
		}
	})

	test('an SVG size is read from width and height, not from an attribute whose name ends in them', async () => {
		const size = (svg) => JSON.stringify(getImageSizeFromBytes(new TextEncoder().encode(svg)))
		const expected = JSON.stringify({ w: 200, h: 100 })
		assertEqual(
			size('<svg xmlns="http://www.w3.org/2000/svg" stroke-width="2" width="200" height="100"/>'),
			expected,
			'stroke-width first'
		)
		assertEqual(
			size('<svg xmlns="http://www.w3.org/2000/svg" data-height="7" width="200" height="100"/>'),
			expected,
			'data-height first'
		)
		assertEqual(size('<svg\n\twidth=\'200\'\theight="100"/>'), expected, 'other whitespace and quotes')
		const frame = pictureFrame(
			await slideXml((p) =>
				p.addSlide().addImage({
					svg: '<svg xmlns="http://www.w3.org/2000/svg" stroke-width="2" width="200" height="100"/>',
					x: 0,
					y: 0,
					w: 2,
				})
			)
		)
		assertEqual(frame.cy, EMU_PER_INCH, 'a 2:1 SVG at 2in is 1in tall')
	})

	test('assigning undefined to a slide background drops the image it held', async () => {
		let after
		const xml = await slideXml((p) => {
			const slide = p.addSlide()
			slide.background = { data: pngData(4, 4) }
			slide.background = undefined
			after = slide.background
		})
		assertEqual(after, undefined, 'the getter reads no background')
		assert(!/<p:bg>[\s\S]*?<a:blip/.test(xml), 'and the slide paints no background image')
	})

	test('a table and an image with a null axis both take it from the placeholder', async () => {
		// The image treated `null` as unstated and the table did not: a table with `x: null` landed at
		// the half-inch default rather than the placeholder's 3in.
		const withPlaceholder = (p) =>
			p.defineSlideMaster({
				title: 'M',
				objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 3, y: 1, w: 4, h: 2 }, text: '' } }],
			})
		const unset = /** @type {any} */ (null)
		const tableXml = await slideXml((p) => {
			withPlaceholder(p)
			p.addSlide({ masterTitle: 'M' }).addTable([[{ text: 't' }]], { placeholder: 'body', x: unset })
		})
		const tableX = /<p:graphicFrame>[\s\S]*?<a:off x="(-?\d+)"/.exec(tableXml)?.[1]
		assertEqual(Number(tableX), 3 * EMU_PER_INCH, 'the table')
		const imageXml = await slideXml((p) => {
			withPlaceholder(p)
			p.addSlide({ masterTitle: 'M' }).addImage({ data: pngData(32, 32), placeholder: 'body', x: unset })
		})
		assertEqual(pictureFrame(imageXml).x, 3 * EMU_PER_INCH, 'the image')
	})

	test('Picture.setImage refuses a fit on a zero-extent picture, and leaves the picture as it was', async () => {
		const { buf } = await build((p) => p.addSlide().addImage({ data: pngData(32, 32), x: 1, y: 1, w: 0, h: 0 }))
		const deck = await Presentation.load(buf)
		const picture = deck.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		const embed = () => picture.element_.getElementsByTagNameNS('*', 'blip')[0]?.getAttribute('r:embed')
		const before = embed()
		let code
		try {
			picture.setImage(pngBytes(16, 16), { contentType: 'image/png', fit: 'cover' })
		} catch (err) {
			code = /** @type {any} */ (err).code
		}
		assertEqual(code, 'image/fit-needs-extent', 'the refusal')
		assertEqual(embed(), before, 'the picture still shows its own image')
	})
})
