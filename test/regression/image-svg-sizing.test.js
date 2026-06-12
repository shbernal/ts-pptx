import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// A square SVG (1:1) — its intrinsic aspect must come from width/height or viewBox,
// never from the displayed box. Placed in a wide box with sizing:'contain' it should
// letterbox (non-zero srcRect), not stretch.
const SQUARE_SVG_VIEWBOX =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" fill="#123456"/></svg>'
const SQUARE_SVG_WH =
	'<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100" fill="#123456"/></svg>'

function srcRectAttrs(xml) {
	const m = /<a:srcRect\b([^/]*)\/>/.exec(xml)
	if (!m) return null
	const attrs = {}
	for (const a of m[1].matchAll(/(\w+)="(-?\d+)"/g)) attrs[a[1]] = parseInt(a[2], 10)
	return attrs
}

defineRegressionSuite('Image SVG sizing', [
	{
		name: "sizing:'contain' reads a square SVG's viewBox and letterboxes (non-zero srcRect) in a wide box",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ svg: SQUARE_SVG_VIEWBOX, x: 1, y: 1, w: 4, h: 1, sizing: { type: 'contain', w: 4, h: 1 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const rect = srcRectAttrs(xml)
			assert(rect, 'expected an <a:srcRect> from sizing:contain; got: ' + xml)
			// 1:1 image in a 4:1 box → horizontal letterbox: l/r diverge from 0, t/b stay 0.
			assert(
				rect.l !== 0 && rect.l === rect.r,
				`expected symmetric non-zero horizontal inset; got ${JSON.stringify(rect)}`
			)
			assert((rect.t || 0) === 0 && (rect.b || 0) === 0, `expected zero vertical inset; got ${JSON.stringify(rect)}`)
		},
	},
	{
		name: "sizing:'contain' falls back to a square SVG's width/height when no viewBox",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ svg: SQUARE_SVG_WH, x: 1, y: 1, w: 4, h: 1, sizing: { type: 'contain', w: 4, h: 1 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const rect = srcRectAttrs(xml)
			assert(rect, 'expected an <a:srcRect>; got: ' + xml)
			assert(
				rect.l !== 0 && rect.l === rect.r,
				`expected non-zero horizontal inset from width/height; got ${JSON.stringify(rect)}`
			)
		},
	},
	{
		name: "sizing:'cover' reads a square SVG aspect and crops vertically in a wide box",
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ svg: SQUARE_SVG_VIEWBOX, x: 1, y: 1, w: 4, h: 1, sizing: { type: 'cover', w: 4, h: 1 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const rect = srcRectAttrs(xml)
			assert(rect, 'expected an <a:srcRect> from sizing:cover; got: ' + xml)
			// 1:1 image filling a 4:1 box → crop top/bottom: t/b diverge, l/r stay 0.
			assert(
				rect.t !== 0 && rect.t === rect.b,
				`expected symmetric non-zero vertical crop; got ${JSON.stringify(rect)}`
			)
			assert((rect.l || 0) === 0 && (rect.r || 0) === 0, `expected zero horizontal crop; got ${JSON.stringify(rect)}`)
		},
	},
])
