import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../../helpers.js'
import { EMU_PER_INCH } from '../../../dist/node.js'

// A square SVG (1:1) — its intrinsic aspect must come from width/height or viewBox,
// never from the displayed box. Placed in a wide box with sizing:'contain' it should
// letterbox (non-zero srcRect), not stretch.
const SQUARE_SVG_VIEWBOX =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" fill="#123456"/></svg>'
const SQUARE_SVG_WH =
	'<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100" fill="#123456"/></svg>'
// 2:1 — an intrinsic ratio that is neither square nor the box's, so a derived dimension and a
// letterbox are both visible in the output.
const WIDE_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><rect width="200" height="100"/></svg>'
// Neither width/height nor viewBox: nothing to measure, so nothing to place aspect-correctly.
const UNMEASURABLE_SVG = '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>'
// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

function srcRectAttrs(xml) {
	const m = /<a:srcRect\b([^/]*)\/>/.exec(xml)
	if (!m) return null
	const attrs = {}
	for (const a of m[1].matchAll(/(\w+)="(-?\d+)"/g)) attrs[a[1]] = parseInt(a[2], 10)
	return attrs
}

/** The picture's displayed extent (`<a:ext>`), in EMU. Scoped to `<p:pic>`: the slide's own
 *  `<p:grpSpPr>` carries an `<a:ext cx="0" cy="0"/>` that would otherwise match first. */
function pictureExtent(xml) {
	const pic = xml.split('<p:pic>')[1]
	assert(pic, 'expected a <p:pic>; got: ' + xml)
	const m = /<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/.exec(pic)
	assert(m, 'expected an <a:ext> for the picture; got: ' + xml)
	return { cx: parseInt(m[1], 10), cy: parseInt(m[2], 10) }
}

async function slideXmlFor(opts) {
	const { zip } = await build((p) => {
		const s = p.addSlide()
		s.addImage(opts)
	})
	return readEntry(zip, 'ppt/slides/slide1.xml')
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

	// --- Aspect-correct placement is the DEFAULT for a vector source, not an opt-in. ---
	// A glyph squashed into a box whose ratio disagrees with its viewBox is a defect, and the
	// caller who wrote `{ svg, w, h }` did not ask for it. Rasters keep filling their box.
	{
		name: 'an SVG whose ratio differs from its box letterboxes with no sizing asked for',
		fn: async () => {
			const xml = await slideXmlFor({ svg: SQUARE_SVG_VIEWBOX, x: 1, y: 1, w: 4, h: 1 })
			const rect = srcRectAttrs(xml)
			assert(rect, 'expected an implicit <a:srcRect> letterbox; got: ' + xml)
			assert(
				rect.l !== 0 && rect.l === rect.r,
				`expected symmetric non-zero horizontal inset; got ${JSON.stringify(rect)}`
			)
			assert((rect.t || 0) === 0 && (rect.b || 0) === 0, `expected zero vertical inset; got ${JSON.stringify(rect)}`)
		},
	},
	{
		name: "the implicit letterbox is byte-identical to asking for sizing:'contain'",
		fn: async () => {
			const implicit = await slideXmlFor({ svg: WIDE_SVG, x: 1, y: 1, w: 3, h: 3 })
			const explicit = await slideXmlFor({ svg: WIDE_SVG, x: 1, y: 1, w: 3, h: 3, sizing: { type: 'contain' } })
			assertEqual(implicit, explicit, 'the default must be the same request, not a near-miss of it')
		},
	},
	{
		name: 'an SVG already matching its box still emits a plain stretch — no srcRect, no new bytes',
		fn: async () => {
			const xml = await slideXmlFor({ svg: SQUARE_SVG_VIEWBOX, x: 1, y: 1, w: 2, h: 2 })
			assertEqual(srcRectAttrs(xml), null, 'a matching ratio must not gain a zero-inset srcRect')
			assert(/<a:stretch>/.test(xml), 'expected the plain stretch fill; got: ' + xml)
		},
	},
	{
		name: 'a raster still fills its box — the default changed for vectors only',
		fn: async () => {
			const xml = await slideXmlFor({ data: PNG_DATA, x: 1, y: 1, w: 4, h: 1 })
			assertEqual(srcRectAttrs(xml), null, 'a raster must keep stretching to its box')
		},
	},
	{
		name: "sizing:'stretch' opts a vector back out of the aspect-correct default",
		fn: async () => {
			const xml = await slideXmlFor({ svg: SQUARE_SVG_VIEWBOX, x: 1, y: 1, w: 4, h: 1, sizing: { type: 'stretch' } })
			assertEqual(srcRectAttrs(xml), null, "sizing:'stretch' must emit no source rectangle")
			assert(/<a:stretch>/.test(xml), 'expected the plain stretch fill; got: ' + xml)
		},
	},
	{
		name: 'an unmeasurable SVG falls through to stretch silently — nothing was asked for',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(() =>
				slideXmlFor({ svg: UNMEASURABLE_SVG, x: 1, y: 1, w: 4, h: 1 })
			)
			assertEqual(srcRectAttrs(xml), null, 'an unmeasurable vector cannot be letterboxed')
			assert(
				!codes.includes('image/unmeasurable-natural-size'),
				`the implicit path must not warn about a size the caller never asked to use; got ${JSON.stringify(codes)}`
			)
		},
	},

	// --- Intrinsic ratio fills in a missing dimension; intrinsic magnitude never does. ---
	{
		name: 'an SVG given only w derives h from its viewBox instead of the 1in fallback',
		fn: async () => {
			const xml = await slideXmlFor({ svg: WIDE_SVG, x: 1, y: 1, w: 4 })
			const ext = pictureExtent(xml)
			assertEqual(ext.cx, 4 * EMU_PER_INCH, 'the supplied width must be kept exactly')
			assertEqual(ext.cy, 2 * EMU_PER_INCH, 'a 2:1 viewBox in a 4in width is 2in tall')
		},
	},
	{
		name: 'an SVG given only h derives w from its viewBox',
		fn: async () => {
			const ext = pictureExtent(await slideXmlFor({ svg: WIDE_SVG, x: 1, y: 1, h: 1.5 }))
			assertEqual(ext.cy, 1.5 * EMU_PER_INCH, 'the supplied height must be kept exactly')
			assertEqual(ext.cx, 3 * EMU_PER_INCH, 'a 2:1 viewBox at 1.5in tall is 3in wide')
		},
	},
	{
		name: 'an SVG given neither dimension keeps the 1in fallback — user units are not pixels',
		fn: async () => {
			// A 200x100 viewBox at 96 DPI would be 2.08in x 1.04in, and a 24-unit icon a quarter
			// inch. SVG user units are dependable relative to each other and conventional in
			// absolute terms, so only the ratio is trusted and an unanchored vector stays 1in.
			const ext = pictureExtent(await slideXmlFor({ svg: WIDE_SVG, x: 1, y: 1 }))
			assertEqual(ext.cx, EMU_PER_INCH, 'expected the 1in fallback width')
			assertEqual(ext.cy, EMU_PER_INCH, 'expected the 1in fallback height')
		},
	},

	// --- sizing.w / sizing.h default to the picture's own box. ---
	{
		name: "sizing w/h are optional and default to the picture's box",
		fn: async () => {
			const short = await slideXmlFor({ data: PNG_DATA, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'cover' } })
			const long = await slideXmlFor({ data: PNG_DATA, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'cover', w: 4, h: 3 } })
			assertEqual(short, long, 'omitting sizing w/h must mean the picture box, not a different crop')
		},
	},
])
