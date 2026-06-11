import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// `cover`/`contain` crop the *source* bitmap, so the emitted `<a:srcRect>` must be derived
// from the image's NATURAL pixel ratio — not the displayed box (options.w/h). Previously the
// displayed EMU size was passed as the image size, so whenever display ratio == box ratio the
// crop collapsed to all-zeros (no crop), squashing the image instead of cover-fitting it.

// 1x1 transparent PNG → natural ratio is 1:1 (square), independent of how it is displayed.
const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

// Synthesize a PNG header carrying an arbitrary intrinsic w×h. getImageSizeFromBase64 reads only
// the IHDR dimension bytes (width@16 / height@20, big-endian), so a 24-byte header is enough to
// exercise the natural-ratio crop math with exact, self-documenting dimensions.
function pngOf(w, h) {
	const b = Buffer.alloc(24)
	b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0)
	b.writeUInt32BE(13, 8)
	b.write('IHDR', 12, 'ascii')
	b.writeUInt32BE(w, 16)
	b.writeUInt32BE(h, 20)
	return 'image/png;base64,' + b.toString('base64')
}

// A srcRect is PowerPoint-valid only if the cropped source keeps positive area: l+r and t+b must
// each stay below 100% (100000). Negative edges (outset/letterbox) are legal. Out-of-range values
// are exactly what triggered the #1286 repair dialog, so every case asserts this invariant.
function assertValidSrcRect(r, label) {
	for (const k of ['l', 'r', 't', 'b']) {
		assert(Number.isInteger(r[k]), `${label}: srcRect ${k} must be a finite integer; got ${r[k]}`)
	}
	assert(r.l + r.r < 1e5, `${label}: l+r must keep positive source width (<100000); got ${r.l}+${r.r}`)
	assert(r.t + r.b < 1e5, `${label}: t+b must keep positive source height (<100000); got ${r.t}+${r.b}`)
}

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
	// upstream #1286: mixed "pixel-like" dimensions straddling 100 (some <100, some ≥100) once
	// produced invalid out-of-range srcRect crop values and a PowerPoint repair. The old converter
	// guessed units by magnitude (a number ≥100 was treated as already-EMU, <100 as inches), so a
	// single object's two dimensions could be resolved in *different* units, wrecking the box ratio.
	// That magnitude guess is gone — a bare number is always inches — and cover/contain now measure
	// the natural pixel ratio from the image bytes. These cases lock in a valid, in-bounds srcRect
	// for the exact straddling-100 shape that used to trigger repair.
	{
		// contain is the original repair surface: with mixed units the letterbox inset blew up to a
		// wildly out-of-range negative percentage. natural 200×80 (ratio 0.4) into a 120×80 box must
		// pad top/bottom by a modest, in-bounds inset and leave left/right flush.
		name: 'contain: dimensions straddling 100 stay in-bounds (no #1286 repair)',
		fn: async () => {
			const r = await srcRectFor({
				data: pngOf(200, 80),
				x: 1,
				y: 1,
				w: 120,
				h: 80,
				sizing: { type: 'contain', w: 120, h: 80 },
			})
			assertValidSrcRect(r, 'contain 200x80 into 120x80')
			assert(
				r.l === 0 && r.r === 0 && r.t === -33333 && r.b === -33333,
				`expected l=0 r=0 t=-33333 b=-33333; got ${JSON.stringify(r)}`
			)
		},
	},
	{
		// cover counterpart: same straddling-100 shape crops left/right from the natural 0.4 ratio.
		name: 'cover: dimensions straddling 100 crop from natural ratio (no #1286 repair)',
		fn: async () => {
			const r = await srcRectFor({
				data: pngOf(200, 80),
				x: 1,
				y: 1,
				w: 120,
				h: 80,
				sizing: { type: 'cover', w: 120, h: 80 },
			})
			assertValidSrcRect(r, 'cover 200x80 into 120x80')
			assert(
				r.l === 20000 && r.r === 20000 && r.t === 0 && r.b === 0,
				`expected l=20000 r=20000 t=0 b=0; got ${JSON.stringify(r)}`
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
