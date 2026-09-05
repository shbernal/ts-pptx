#!/usr/bin/env node
/**
 * Draw the default video poster into `src/media/playbtn.ts`.
 *
 * The artwork `addMedia` falls back to when the caller passes no `cover`: a translucent scrim
 * with a white-ringed disc and a play triangle. It used to be a pasted blob — 74,380 base64
 * characters, some 54 kB of PNG, for four flat colours — and nobody could tell from reading the
 * module that the size was noise rather than detail. Drawn from these constants it is a quarter
 * of that, at a resolution no realistic media frame outruns.
 *
 * **The geometry is measured, not invented.** Every fraction below was read off the original
 * artwork's own pixels (centre, outer radius, ring thickness, triangle bbox) so the poster keeps
 * the shape decks have always shown. Two things did change and both are deliberate: the button is
 * centred, where the original sat 22 px left of centre in a 1920-wide frame, and the stray 6-pixel
 * olive border the original carried around all four edges is gone.
 *
 * **Composite in premultiplied alpha.** The scrim is `rgba(9,9,9,55)` and the ring is opaque
 * white, and averaging those two straight — as an antialiaser that ignores alpha does — puts a
 * dark halo around every edge of the button, because a nearly transparent black contributes its
 * black to the colour average at full weight. Sample coverage in premultiplied space and
 * un-premultiply once at the end.
 *
 *   node scripts/gen-playbtn.mjs            # rewrite src/media/playbtn.ts
 *   node scripts/gen-playbtn.mjs --check    # exit 1 if the file is not what this would write
 */

import fs from 'node:fs'
import path from 'node:path'
import { encodePngRgba } from './png-utils.mjs'
import { ROOT, isMain, parseCliOrExit } from './script-utils.mjs'

/** The module this writes. */
export const PLAYBTN_PATH = path.join(ROOT, 'src', 'media', 'playbtn.ts')

/**
 * The frame, half the original artwork's 1920x1383 in each dimension.
 *
 * The poster is stretched to the media object's box by an `<a:stretch><a:fillRect/>`, so its
 * aspect ratio reaches nothing and its resolution only has to survive the largest box a caller
 * plausibly gives it. A full-slide 16:9 video is 1280 device pixels wide at 96 DPI; the disc is
 * 316 px across here, which is still upscaling by less than 2x on that worst case, and the whole
 * shape is four flat colours where softness would show least.
 */
export const FRAME = { w: 960, h: 692 }

/**
 * The design, in fractions of the frame height — the dimension the original's circle was sized
 * against (its horizontal and vertical radii differed by 3.5 px in 1383, which is measurement
 * noise, not an ellipse).
 */
const DESIGN = {
	/** White ring, outer edge. */
	rOuter: 357 / 1383,
	/** Grey disc, i.e. the ring's inner edge. */
	rDisc: 316 / 1383,
	/** Play triangle: base at the left, apex at the right. */
	triangleW: 274 / 1383,
	triangleH: 281 / 1383,
	/** The triangle's centroid sits a shade left of the disc's centre, as it did originally. */
	triangleDx: -4 / 1383,
}

/** The four colours, RGBA, non-premultiplied. */
const SCRIM = [9, 9, 9, 55]
const RING = [255, 255, 255, 255]
const DISC = [102, 102, 102, 255]

/** Samples per pixel per axis. 4 gives 17 coverage levels, which is past where an edge reads smooth. */
const SUPERSAMPLE = 4

/**
 * Render the poster as non-premultiplied 8-bit RGBA.
 * @param {{w: number, h: number}} frame - the pixel dimensions to draw into
 * @returns {Uint8Array} `w * h * 4` bytes, row-major
 */
export function renderPlayButton(frame) {
	const { w, h } = frame
	const out = new Uint8Array(w * h * 4)
	const cx = w / 2
	const cy = h / 2
	const rOuter = DESIGN.rOuter * h
	const rDisc = DESIGN.rDisc * h
	const triW = DESIGN.triangleW * h
	const triH = DESIGN.triangleH * h
	const triLeft = cx - triW / 2 + DESIGN.triangleDx * h

	/**
	 * Whether a sample falls inside the play triangle: base at `triLeft`, apex at `triLeft + triW`,
	 * half-height shrinking linearly from `triH / 2` to zero across that span.
	 * @param {number} x - sample x
	 * @param {number} y - sample y
	 */
	const inTriangle = (x, y) => {
		if (x < triLeft || x > triLeft + triW) return false
		return Math.abs(y - cy) <= (triH / 2) * (1 - (x - triLeft) / triW)
	}

	const samples = SUPERSAMPLE * SUPERSAMPLE
	for (let py = 0; py < h; py++) {
		for (let px = 0; px < w; px++) {
			// Accumulate in premultiplied space; see the header.
			let r = 0
			let g = 0
			let b = 0
			let a = 0
			for (let sy = 0; sy < SUPERSAMPLE; sy++) {
				for (let sx = 0; sx < SUPERSAMPLE; sx++) {
					const x = px + (sx + 0.5) / SUPERSAMPLE
					const y = py + (sy + 0.5) / SUPERSAMPLE
					const d = Math.hypot(x - cx, y - cy)
					const colour = d > rOuter ? SCRIM : d > rDisc || inTriangle(x, y) ? RING : DISC
					const alpha = (colour[3] ?? 0) / 255
					r += (colour[0] ?? 0) * alpha
					g += (colour[1] ?? 0) * alpha
					b += (colour[2] ?? 0) * alpha
					a += colour[3] ?? 0
				}
			}
			const i = (py * w + px) * 4
			const meanAlpha = a / samples
			// Un-premultiply. A fully transparent pixel cannot arise here (the scrim is the floor),
			// but the guard keeps the division honest if the design ever grows one.
			const scale = meanAlpha > 0 ? 255 / meanAlpha / samples : 0
			out[i] = Math.round(r * scale)
			out[i + 1] = Math.round(g * scale)
			out[i + 2] = Math.round(b * scale)
			out[i + 3] = Math.round(meanAlpha)
		}
	}
	return out
}

/**
 * The whole `src/media/playbtn.ts` source, artwork included.
 * @returns {string} the module text, ending in a newline
 */
export function playbtnModule() {
	const png = encodePngRgba(FRAME.w, FRAME.h, renderPlayButton(FRAME))
	const base64 = png.toString('base64')
	return `/**
 * The default poster frame for an embedded video: PowerPoint's play-button overlay.
 *
 * Alone in a module because of its size. ${base64.length.toLocaleString('en-US')} base64 characters is about ${Math.round(png.length / 1024)} kB of PNG, and
 * base64 of an already-compressed image barely deflates, so it is still the largest thing the
 * write path can reach. The chunker splits by module, so every chunk reaching this module is
 * charged the whole payload; keeping it alone is what lets \`gen/media.ts\` reach it through a
 * dynamic import instead, and leave it out of the entry chunk.
 *
 * It is resolved during the async media pass, for a deck that actually carries media and only
 * where the caller supplied no \`cover\` of their own — see \`SlideRelMedia.isDefaultCover\`.
 *
 * **Generated — do not hand-edit.** \`scripts/gen-playbtn.mjs\` draws it from measured geometry and
 * writes this file; that script's header is where the design and its two deliberate departures
 * from the original artwork are written down. \`pnpm run media:playbtn\` rewrites it.
 */
export const IMG_PLAYBTN =
\t'data:image/png;base64,${base64}'
`
}

if (isMain(import.meta.url)) {
	const { values } = parseCliOrExit(process.argv.slice(2), {
		options: { check: { type: 'boolean', default: false } },
		usage: 'usage: node scripts/gen-playbtn.mjs [--check]',
	})
	const source = playbtnModule()
	if (values.check) {
		const onDisk = fs.readFileSync(PLAYBTN_PATH, 'utf8')
		if (onDisk === source) console.log(`${path.relative(ROOT, PLAYBTN_PATH)} is up to date`)
		else {
			console.error(`${path.relative(ROOT, PLAYBTN_PATH)} differs from what gen-playbtn.mjs writes`)
			process.exit(1)
		}
	} else {
		fs.writeFileSync(PLAYBTN_PATH, source)
		console.log(`wrote ${path.relative(ROOT, PLAYBTN_PATH)}: ${FRAME.w}x${FRAME.h}, ${source.length} bytes of module`)
	}
}
