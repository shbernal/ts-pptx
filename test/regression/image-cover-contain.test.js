import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// `cover`/`contain` crop the *source* bitmap, so the emitted `<a:srcRect>` must be derived
// from the image's NATURAL pixel ratio — not the displayed box (options.w/h). Previously the
// displayed EMU size was passed as the image size, so whenever display ratio == box ratio the
// crop collapsed to all-zeros (no crop), squashing the image instead of cover-fitting it.

// 1x1 transparent PNG → natural ratio is 1:1 (square), independent of how it is displayed.
const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

// Minimal 4x2 (ratio 0.5) raster headers, one per supported format. Only the dimension
// header bytes are meaningful — the parser reads headers, never pixel data.
const RASTER_4x2 = {
	png: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAAAAAA=',
	gif: 'image/gif;base64,R0lGODlhBAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=',
	bmp: 'image/bmp;base64,Qk0AAAAAAAAAAAAAAAAAAAAABAAAAAIAAAAAAAAAAAA=',
	webp: 'image/webp;base64,UklGRgAAAABXRUJQVlA4WAAAAAAAAAAAAwAAAQAAAAA=',
	jpeg: 'image/jpeg;base64,/9j/wAARCAACAAQDAAAAAAAAAAAAAAAAAAAAAAAAAAA=',
}

async function srcRectFor(opts) {
	const { zip } = await build((p) => {
		const s = p.addSlide()
		s.addImage(opts)
	})
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const m = /<a:srcRect l="(-?\d+)" r="(-?\d+)" t="(-?\d+)" b="(-?\d+)"\/>/.exec(xml)
	assert(m, 'expected a srcRect element; got: ' + xml)
	return { l: +m[1], r: +m[2], t: +m[3], b: +m[4] }
}

defineRegressionSuite('Image cover/contain natural-ratio crop', [
	{
		// The exact bug: a square image displayed at 4:3 and cover-fit into a 4:3 box.
		// Display ratio == box ratio, which previously yielded all-zeros (no crop).
		// Natural ratio is 1:1, so cover must crop 12.5% off top and bottom.
		name: 'cover: square image into wide (4:3) box crops top/bottom from natural ratio',
		fn: async () => {
			const r = await srcRectFor({ data: PNG_1x1, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'cover', w: 4, h: 3 } })
			assert(
				r.l === 0 && r.r === 0 && r.t === 12500 && r.b === 12500,
				`expected l=0 r=0 t=12500 b=12500; got ${JSON.stringify(r)}`
			)
		},
	},
	{
		// Square image cover-fit into a tall (3:4) box → crop left/right by 12.5%.
		name: 'cover: square image into tall (3:4) box crops left/right from natural ratio',
		fn: async () => {
			const r = await srcRectFor({ data: PNG_1x1, x: 1, y: 1, w: 3, h: 4, sizing: { type: 'cover', w: 3, h: 4 } })
			assert(
				r.l === 12500 && r.r === 12500 && r.t === 0 && r.b === 0,
				`expected l=12500 r=12500 t=0 b=0; got ${JSON.stringify(r)}`
			)
		},
	},
	{
		// contain letterboxes instead of cropping: square image inside a 4:3 box pads
		// left/right (negative srcRect insets) and leaves top/bottom flush.
		name: 'contain: square image into wide (4:3) box pads left/right (negative inset)',
		fn: async () => {
			const r = await srcRectFor({ data: PNG_1x1, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'contain', w: 4, h: 3 } })
			assert(
				r.l === -16667 && r.r === -16667 && r.t === 0 && r.b === 0,
				`expected l=-16667 r=-16667 t=0 b=0; got ${JSON.stringify(r)}`
			)
		},
	},
	// Cross-format coverage: a 4x2 (ratio 0.5) image of each format, cover-fit into a 2x2
	// square box, must crop 25% off left and right (and nothing top/bottom). This passes
	// only if the header parser read the natural 4x2 ratio for that format.
	...Object.entries(RASTER_4x2).map(([fmt, data]) => ({
		name: `cover: ${fmt} 4x2 image into square box crops left/right (parser reads ${fmt} header)`,
		fn: async () => {
			const r = await srcRectFor({ data, x: 1, y: 1, w: 2, h: 2, sizing: { type: 'cover', w: 2, h: 2 } })
			assert(
				r.l === 25000 && r.r === 25000 && r.t === 0 && r.b === 0,
				`${fmt}: expected l=25000 r=25000 t=0 b=0; got ${JSON.stringify(r)}`
			)
		},
	})),
])
