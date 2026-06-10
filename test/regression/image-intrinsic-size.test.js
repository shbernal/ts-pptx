import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// addImage() previously fell back to a 1in x 1in square whenever `w`/`h` were omitted, which
// squished every dimensionless image into the wrong aspect ratio (issue #1351). For base64
// `data` images the bytes are in hand, so we read the natural pixel size synchronously and
// default the missing dimension(s) from it. PowerPoint inserts raster images at 96 DPI, so
// natural pixels / 96 == inches, and inches * 914400 == EMU (the units in the emitted <a:ext>).

// 1x1 transparent PNG → natural size 1x1 px → 1/96in → 9525 EMU per side.
const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

// Minimal 4x2 (ratio 2:1) raster headers, one per supported format. Only the dimension header
// bytes are meaningful — the parser reads headers, never pixel data. 4px/96 → 38100 EMU,
// 2px/96 → 19050 EMU.
const RASTER_4x2 = {
	png: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAAAAAA=',
	gif: 'image/gif;base64,R0lGODlhBAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=',
	bmp: 'image/bmp;base64,Qk0AAAAAAAAAAAAAAAAAAAAABAAAAAIAAAAAAAAAAAA=',
	webp: 'image/webp;base64,UklGRgAAAABXRUJQVlA4WAAAAAAAAAAAAwAAAQAAAAA=',
	jpeg: 'image/jpeg;base64,/9j/wAARCAACAAQDAAAAAAAAAAAAAAAAAAAAAAAAAAA=',
}

// Tiny inline SVG (vector → no intrinsic pixel size, must keep the 1in fallback).
const SVG_DATA = 'image/svg+xml;base64,' + Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>').toString('base64')

// A PNG header declaring a 10000x1 image. Natural width 10000px / 96 = 104.1667in — above the old
// "a number >= 100 must be EMU" threshold, so the previous code collapsed it to ~104 EMU (a dot).
// With magnitude guessing gone, 10000/96 inches resolves correctly to 95,250,000 EMU.
const PNG_10000x1 =
	'image/png;base64,' +
	Buffer.concat([
		Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
		Buffer.from([0, 0, 0, 0x0d]),
		Buffer.from('IHDR'),
		Buffer.from([0, 0, 0x27, 0x10]), // width = 10000
		Buffer.from([0, 0, 0, 0x01]), // height = 1
	]).toString('base64')

async function extFor(opts) {
	const { zip } = await build((p) => {
		const s = p.addSlide()
		s.addImage(opts)
	})
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	// Scope to the picture block — the spTree's <p:grpSpPr> also carries an <a:ext cx="0" cy="0"/>.
	const pic = xml.slice(xml.indexOf('<p:pic'))
	const m = /<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(pic)
	assert(m, 'expected an <a:ext> element in the picture block; got: ' + xml)
	return { cx: +m[1], cy: +m[2] }
}

defineRegressionSuite('Image intrinsic-size defaults', [
	{
		// Neither w nor h: square 1x1 image must emit its natural 9525x9525 EMU box, not 1in.
		name: 'no w/h: data image defaults to its natural pixel size (square)',
		fn: async () => {
			const r = await extFor({ data: PNG_1x1, x: 1, y: 1 })
			assert(r.cx === 9525 && r.cy === 9525, `expected cx=9525 cy=9525; got ${JSON.stringify(r)}`)
		},
	},
	// Cross-format coverage: a 4x2 image of each format with no w/h must emit 38100x19050 EMU.
	// This passes only if the header parser read the natural 4x2 size for that format.
	...Object.entries(RASTER_4x2).map(([fmt, data]) => ({
		name: `no w/h: ${fmt} 4x2 image defaults to 38100x19050 EMU (parser reads ${fmt} header)`,
		fn: async () => {
			const r = await extFor({ data, x: 1, y: 1 })
			assert(r.cx === 38100 && r.cy === 19050, `${fmt}: expected cx=38100 cy=19050; got ${JSON.stringify(r)}`)
		},
	})),
	{
		// Only width given (2in): height is derived from the 2:1 natural ratio → 1in.
		name: 'w only: height is derived from natural aspect ratio',
		fn: async () => {
			const r = await extFor({ data: RASTER_4x2.png, x: 1, y: 1, w: 2 })
			assert(r.cx === 1828800 && r.cy === 914400, `expected cx=1828800 cy=914400; got ${JSON.stringify(r)}`)
		},
	},
	{
		// Only height given (2in): width is derived from the 2:1 natural ratio → 4in.
		name: 'h only: width is derived from natural aspect ratio',
		fn: async () => {
			const r = await extFor({ data: RASTER_4x2.png, x: 1, y: 1, h: 2 })
			assert(r.cx === 3657600 && r.cy === 1828800, `expected cx=3657600 cy=1828800; got ${JSON.stringify(r)}`)
		},
	},
	{
		// Both given: explicit dimensions always win, intrinsic size is ignored.
		name: 'explicit w and h are never overridden by intrinsic size',
		fn: async () => {
			const r = await extFor({ data: RASTER_4x2.png, x: 1, y: 1, w: 3, h: 3 })
			assert(r.cx === 2743200 && r.cy === 2743200, `expected cx=2743200 cy=2743200; got ${JSON.stringify(r)}`)
		},
	},
	{
		// A >9600px image used to collapse under the EMU-passthrough heuristic; now it sizes correctly.
		name: 'very large image (10000px wide) no longer collapses under unit guessing',
		fn: async () => {
			const r = await extFor({ data: PNG_10000x1, x: 1, y: 1 })
			assert(r.cx === 95250000 && r.cy === 9525, `expected cx=95250000 cy=9525; got ${JSON.stringify(r)}`)
		},
	},
	{
		// SVG (and any path image, which can't be measured synchronously) has no intrinsic
		// pixel size here → keep the legacy 1in (914400 EMU) fallback.
		name: 'svg data keeps the 1in fallback (no intrinsic pixel size)',
		fn: async () => {
			const r = await extFor({ data: SVG_DATA, x: 1, y: 1 })
			assert(r.cx === 914400 && r.cy === 914400, `expected cx=914400 cy=914400; got ${JSON.stringify(r)}`)
		},
	},
])
