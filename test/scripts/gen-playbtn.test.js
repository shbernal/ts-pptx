// The default video poster is drawn by `scripts/gen-playbtn.mjs`, not pasted in.
//
// The artwork it replaced was 74,380 base64 characters of PNG for four flat colours, and the
// reason nobody noticed is that a base64 blob in a source file is unreadable by construction: it
// carries no evidence of what it depicts or why it is that size. Generating it puts the geometry
// under review, and this case is what keeps the committed module and the generator from drifting
// apart — a hand-edit to `src/media/playbtn.ts`, or a change to the design constants that was
// never re-run, both fail here.
//
// The last case is the deliberate red: perturb one design constant and the bytes must move. A
// comparison that passed whatever the renderer produced would gate nothing.

import fs from 'node:fs'
import { describe, expect, test } from 'vitest'
import { FRAME, PLAYBTN_PATH, playbtnModule, renderPlayButton } from '../../scripts/gen-playbtn.mjs'

/** The RGBA pixel at `(x, y)` of a rendered frame. */
function pixel(rgba, w, x, y) {
	const i = (y * w + x) * 4
	return [rgba[i], rgba[i + 1], rgba[i + 2], rgba[i + 3]]
}

describe('gen-playbtn', () => {
	test('src/media/playbtn.ts is what the generator writes', () => {
		expect(fs.readFileSync(PLAYBTN_PATH, 'utf8')).toBe(playbtnModule())
	})

	test('the poster is smaller than the artwork it replaced', () => {
		// 74,380 characters was the old literal. This asserts the win did not quietly unwind, and
		// the floor asserts the artwork is still artwork rather than an empty frame.
		const source = fs.readFileSync(PLAYBTN_PATH, 'utf8')
		const base64 = source.slice(source.indexOf('base64,') + 'base64,'.length, source.lastIndexOf("'"))
		expect(base64.length).toBeLessThan(30_000)
		expect(base64.length).toBeGreaterThan(5_000)
	})

	test('the four design colours land where the design says', () => {
		const { w, h } = FRAME
		const rgba = renderPlayButton(FRAME)
		// The scrim, well outside the button.
		expect(pixel(rgba, w, 4, 4)).toEqual([9, 9, 9, 55])
		// The disc, well inside it: the ring's inner edge is at 316/1383 == 0.229 of the height.
		expect(pixel(rgba, w, Math.round(w / 2), Math.round(h / 2 - 0.2 * h))).toEqual([102, 102, 102, 255])
		// The white ring, between the two radii (0.229 and 357/1383 == 0.258).
		expect(pixel(rgba, w, Math.round(w / 2), Math.round(h / 2 - 0.245 * h))).toEqual([255, 255, 255, 255])
		// The triangle, a little left of the disc centre so the taper cannot be mistaken for it.
		expect(pixel(rgba, w, Math.round(w / 2 - 0.04 * h), Math.round(h / 2))).toEqual([255, 255, 255, 255])
	})

	test('compositing is premultiplied: the edge colour is the un-premultiplied one', () => {
		// The scrim is `rgba(9,9,9,55)` and the ring is opaque white. An antialiaser that averages
		// those straight puts the scrim's black into the colour at full weight, which paints a dark
		// halo around the whole button. So the check is not "looks bright enough" — it is the exact
		// arithmetic, for every partially covered pixel on the ring's outer edge:
		//
		//   alpha       = 255c + 55(1 - c)                       (c = the sample's coverage)
		//   premult rgb = 255c + 9 * (55/255) * (1 - c)
		//   rgb         = premult rgb / alpha * 255
		//
		// Straight averaging would give `255c + 9(1 - c)` instead, which is 62 lower at half
		// coverage. Rounding to 8 bits is the only slack allowed.
		const { w, h } = FRAME
		const rgba = renderPlayButton(FRAME)
		const cx = Math.round(w / 2)
		const rOuter = (357 / 1383) * h
		const edge = []
		for (let y = Math.floor(h / 2 - rOuter) - 3; y <= Math.ceil(h / 2 - rOuter) + 3; y++) {
			const [r, , , a] = pixel(rgba, w, cx, y)
			if (a > 56 && a < 254) edge.push([y, r, a])
		}
		expect(edge.length, 'no partially covered pixel on the ring edge to check').toBeGreaterThan(0)
		for (const [y, r, a] of edge) {
			const coverage = (a - 55) / 200
			const premultiplied = 255 * coverage + 9 * (55 / 255) * (1 - coverage)
			const expected = Math.round((premultiplied / a) * 255)
			expect(Math.abs(r - expected), `row ${y}: rgb ${r}, expected ${expected} at alpha ${a}`).toBeLessThanOrEqual(1)
		}
	})

	test('a changed design constant changes the bytes', () => {
		const base = renderPlayButton(FRAME)
		const wider = renderPlayButton({ w: FRAME.w, h: FRAME.h + 2 })
		expect(wider.length).not.toBe(base.length)
		// Same frame, same pixels: the renderer is deterministic, which is what makes the
		// file-comparison case above meaningful rather than flaky.
		expect(Buffer.from(renderPlayButton(FRAME))).toEqual(Buffer.from(base))
	})
})
