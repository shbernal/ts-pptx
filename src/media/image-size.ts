/**
 * Intrinsic image geometry: reading a natural pixel size out of header bytes, and the
 * `<a:srcRect>` crop math that fits that natural size into a display box.
 *
 * Shared by both sides of the library — the write path (`addImage({ sizing })`) and the read
 * API's `Picture.setImage({ fit })` — which is why it lives outside `gen/` and `read/`.
 * Everything here is synchronous and pure: headers only, never pixel decoding.
 */

import { decodeBase64ToBytes } from './base64.js'

/**
 * Read the intrinsic dimensions of an image from its header bytes.
 * - synchronous: parses only file-format headers, never decodes pixels
 * - raster: PNG, JPEG, GIF, BMP, and WebP (VP8 / VP8L / VP8X) — natural pixels
 * - vector: SVG — intrinsic size from the root `<svg>` width/height or viewBox
 * - unrecognized formats return `null` (no measurable intrinsic size)
 *
 * Used by image `sizing: 'cover' | 'contain'` to compute an aspect-correct
 * `<a:srcRect>` crop from the *natural* image ratio rather than the displayed box.
 * @param {string} dataB64 - base64 image payload or `data:` URI
 * @returns {{ w: number, h: number } | null} natural size, or `null` when unmeasurable
 */
export function getImageSizeFromBase64(dataB64: string): { w: number; h: number } | null {
	const b = decodeBase64ToBytes(dataB64)
	return b ? getImageSizeFromBytes(b) : null
}

/**
 * Read the intrinsic dimensions of an image from raw header bytes — the
 * byte-level core shared by {@link getImageSizeFromBase64} and the read API's
 * `Picture.setImage({ fit })`, which already holds the media bytes.
 * @param {Uint8Array} b - image bytes
 * @returns {{ w: number, h: number } | null} natural size, or `null` when unmeasurable
 */
export function getImageSizeFromBytes(b: Uint8Array): { w: number; h: number } | null {
	if (!b || b.length < 24) return null

	// Bounds-checked byte read: every access below is already guarded by an
	// explicit length check, so the `?? 0` fallback is unreachable in practice.
	const u = (n: number): number => b[n] ?? 0

	// PNG: 8-byte signature, then IHDR with width@16 / height@20 (big-endian uint32)
	if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
		const w = (u(16) << 24) | (u(17) << 16) | (u(18) << 8) | u(19)
		const h = (u(20) << 24) | (u(21) << 16) | (u(22) << 8) | u(23)
		return w > 0 && h > 0 ? { w, h } : null
	}

	// GIF: "GIF87a"/"GIF89a", width@6 / height@8 (little-endian uint16)
	if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
		const w = u(6) | (u(7) << 8)
		const h = u(8) | (u(9) << 8)
		return w > 0 && h > 0 ? { w, h } : null
	}

	// BMP: "BM", width@18 / height@22 (little-endian int32; height may be negative for top-down)
	if (b[0] === 0x42 && b[1] === 0x4d) {
		const w = u(18) | (u(19) << 8) | (u(20) << 16) | (u(21) << 24)
		const h = u(22) | (u(23) << 8) | (u(24) << 16) | (u(25) << 24)
		const aw = Math.abs(w)
		const ah = Math.abs(h)
		return aw > 0 && ah > 0 ? { w: aw, h: ah } : null
	}

	// WebP: "RIFF"...."WEBP" then a VP8 / VP8L / VP8X chunk
	if (
		b[0] === 0x52 &&
		b[1] === 0x49 &&
		b[2] === 0x46 &&
		b[3] === 0x46 &&
		b[8] === 0x57 &&
		b[9] === 0x45 &&
		b[10] === 0x42 &&
		b[11] === 0x50
	) {
		const fourCC = String.fromCharCode(u(12), u(13), u(14), u(15))
		if (fourCC === 'VP8 ' && b.length >= 30) {
			// Lossy: 14-bit width/height at offset 26/28 (little-endian, mask off scale bits)
			const w = (u(26) | (u(27) << 8)) & 0x3fff
			const h = (u(28) | (u(29) << 8)) & 0x3fff
			return w > 0 && h > 0 ? { w, h } : null
		}
		if (fourCC === 'VP8L' && b.length >= 25) {
			// Lossless: 14-bit width/height packed starting at bit 0 of offset 21
			const bits = u(21) | (u(22) << 8) | (u(23) << 16) | (u(24) << 24)
			const w = (bits & 0x3fff) + 1
			const h = ((bits >> 14) & 0x3fff) + 1
			return w > 0 && h > 0 ? { w, h } : null
		}
		if (fourCC === 'VP8X' && b.length >= 30) {
			// Extended: 24-bit canvas width/height minus one at offset 24/27 (little-endian)
			const w = (u(24) | (u(25) << 8) | (u(26) << 16)) + 1
			const h = (u(27) | (u(28) << 8) | (u(29) << 16)) + 1
			return w > 0 && h > 0 ? { w, h } : null
		}
		return null
	}

	// JPEG: "FFD8", scan segment markers for a Start-Of-Frame (SOFn) and read height@5 / width@7
	if (b[0] === 0xff && b[1] === 0xd8) {
		let i = 2
		while (i + 9 < b.length) {
			if (b[i] !== 0xff) {
				i++
				continue
			}
			const marker = u(i + 1)
			// SOF0..SOF15 carry frame dimensions, excluding DHT(C4)/JPG(C8)/DAC(CC)
			if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
				const h = (u(i + 5) << 8) | u(i + 6)
				const w = (u(i + 7) << 8) | u(i + 8)
				return w > 0 && h > 0 ? { w, h } : null
			}
			// Standalone markers (RSTn / SOI / EOI / TEM) have no length payload
			if ((marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) {
				i += 2
				continue
			}
			// Otherwise skip this segment using its 2-byte big-endian length
			const segLen = (u(i + 2) << 8) | u(i + 3)
			if (segLen < 2) break
			i += 2 + segLen
		}
		return null
	}

	// SVG: text-based vector with no binary signature. When the payload is an
	// `<svg>` document, read its intrinsic size from the root element so that
	// `sizing: 'cover' | 'contain'` is aspect-correct for SVG, not just rasters.
	const text = utf8Decode(b)
	if (/<svg[\s>]/i.test(text)) return getSvgSizeFromMarkup(text)

	return null
}

/**
 * Compute the `<a:srcRect>` crop percentages (each in 1/1000 of a percent, the
 * OOXML unit) for fitting an image of natural size `img` into a display `box`,
 * assuming the cropped region is then stretched to fill the box (`<a:stretch>`).
 *
 * - `cover`: fill the box, cropping the overflowing axis (positive l/r or t/b)
 * - `contain`: fit inside the box, letterboxing the short axis (negative l/r or t/b)
 *
 * Single source of truth for the crop math shared by the write side
 * (`ImageSizingXml`) and the read API's `Picture.setImage({ fit })`. `l`/`r` and
 * `t`/`b` are symmetric (centered crop).
 * @param {'cover' | 'contain'} type - fit mode
 * @param {{ w: number, h: number }} img - natural image pixel size
 * @param {{ w: number, h: number }} box - displayed frame size (any consistent unit)
 * @returns {{ l: number, r: number, t: number, b: number }} srcRect percentages
 */
export function fitSrcRectPercents(
	type: 'cover' | 'contain',
	img: { w: number; h: number },
	box: { w: number; h: number }
): { l: number; r: number; t: number; b: number } {
	const imgRatio = img.h / img.w
	const boxRatio = box.h / box.w
	let width: number
	let height: number
	if (type === 'cover') {
		const isBoxBased = boxRatio > imgRatio
		width = isBoxBased ? box.h / imgRatio : box.w
		height = isBoxBased ? box.h : box.w * imgRatio
	} else {
		const widthBased = boxRatio > imgRatio
		width = widthBased ? box.w : box.h / imgRatio
		height = widthBased ? box.w * imgRatio : box.h
	}
	const hz = Math.round(1e5 * 0.5 * (1 - box.w / width))
	const vz = Math.round(1e5 * 0.5 * (1 - box.h / height))
	return { l: hz, r: hz, t: vz, b: vz }
}

/**
 * Read the intrinsic size of an SVG document from its root `<svg>` element.
 * Follows the SVG sizing model: an explicit absolute `width`/`height` pair wins;
 * otherwise the `viewBox` width/height defines the size (and thus aspect ratio).
 * Percentage or missing `width`/`height` fall through to `viewBox`.
 * @param {string} svg - SVG markup
 * @returns {{ w: number, h: number } | null} intrinsic size, or `null` when undeterminable
 */
function getSvgSizeFromMarkup(svg: string): { w: number; h: number } | null {
	const openTag = /<svg\b[^>]*>/i.exec(svg)?.[0]
	if (!openTag) return null
	const attr = (name: string): string | null =>
		new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`, 'i').exec(openTag)?.[1] ?? null
	// Leading number with an optional absolute unit; a percentage is not an intrinsic length.
	const absLength = (val: string | null): number => {
		if (val == null || /%\s*$/.test(val)) return NaN
		const m = /^\s*\+?(\d*\.?\d+)/.exec(val)
		return m ? parseFloat(m[1] ?? '') : NaN
	}
	let w = absLength(attr('width'))
	let h = absLength(attr('height'))
	if (!(w > 0 && h > 0)) {
		const vb = attr('viewBox')
		const p = vb
			? vb
					.trim()
					.split(/[\s,]+/)
					.map(Number)
			: []
		const vw = p[2]
		const vh = p[3]
		if (p.length === 4 && vw != null && vh != null && vw > 0 && vh > 0) {
			w = vw
			h = vh
		}
	}
	return w > 0 && h > 0 ? { w, h } : null
}

/** Decode UTF-8 bytes to a string, isomorphic across Node and browsers. */
function utf8Decode(bytes: Uint8Array): string {
	return new TextDecoder().decode(bytes)
}
