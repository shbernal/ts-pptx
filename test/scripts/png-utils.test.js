// The PNG encoder and decoder the pixel-reading gates use.
//
// `decodePng` is the load-bearing half: `powerpoint-com-smoke.mjs` exports a rendered slide to
// PNG and reads the pixels back to prove PowerPoint *drew* the 3D model rather than falling back
// to its preview. A decoder is the kind of code that fails by returning plausible pixels — a
// scanline unfilter with the wrong predictor still produces an image, just not the one on disk —
// and it had no test at all while it lived inside an 848-line COM driver that only runs by hand,
// on Windows, with PowerPoint installed.
//
// So the cases here are built round-trip (encode a known image, decode it, compare every pixel)
// and the last one is the deliberate red: corrupt one filter byte and the decode must come back
// wrong. A decoder that ignored filters would pass everything above it.

import zlib from 'node:zlib'
import { describe, expect, test } from 'vitest'
import { decodePng, solidPngBase64 } from '../../scripts/png-utils.mjs'

/** CRC-32, as PNG defines it. */
function crc32(buf) {
	const table = Int32Array.from({ length: 256 }, (_, n) => {
		let c = n
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		return c
	})
	let c = -1
	for (const byte of buf) c = table[(c ^ byte) & 0xff] ^ (c >>> 8)
	return (c ^ -1) >>> 0
}

/** One PNG chunk: length, type, data, CRC. */
function chunk(type, data) {
	const len = Buffer.alloc(4)
	len.writeUInt32BE(data.length)
	const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
	const crc = Buffer.alloc(4)
	crc.writeUInt32BE(crc32(body))
	return Buffer.concat([len, body, crc])
}

const SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

/**
 * An 8-bit truecolour PNG from `rows` of `[r, g, b]`, with one filter byte per scanline.
 *
 * `filters` lets a case choose the predictor per row, which is the whole point: PowerPoint's own
 * exporter picks per row, so a decoder that handled only `None` would pass a naive test and fail
 * on a real export.
 */
function encodeRgb(rows, filters = []) {
	const h = rows.length
	const w = rows[0].length
	const stride = w * 3
	const flat = []
	for (let y = 0; y < h; y++) {
		const filter = filters[y] ?? 0
		const line = Buffer.alloc(stride)
		for (let x = 0; x < w; x++) {
			line[x * 3] = rows[y][x][0]
			line[x * 3 + 1] = rows[y][x][1]
			line[x * 3 + 2] = rows[y][x][2]
		}
		const encoded = Buffer.alloc(stride)
		for (let i = 0; i < stride; i++) {
			const a = i >= 3 ? line[i - 3] : 0
			const b = y > 0 ? rows[y - 1][Math.floor(i / 3)][i % 3] : 0
			const c = y > 0 && i >= 3 ? rows[y - 1][Math.floor((i - 3) / 3)][(i - 3) % 3] : 0
			let v = line[i]
			if (filter === 1) v -= a
			else if (filter === 2) v -= b
			else if (filter === 3) v -= (a + b) >> 1
			else if (filter === 4) {
				const p = a + b - c
				const pa = Math.abs(p - a)
				const pb = Math.abs(p - b)
				const pc = Math.abs(p - c)
				v -= pa <= pb && pa <= pc ? a : pb <= pc ? b : c
			}
			encoded[i] = v & 0xff
		}
		flat.push(Buffer.from([filter]), encoded)
	}
	const ihdr = Buffer.alloc(13)
	ihdr.writeUInt32BE(w, 0)
	ihdr.writeUInt32BE(h, 4)
	ihdr[8] = 8
	ihdr[9] = 2
	return Buffer.concat([
		SIGNATURE,
		chunk('IHDR', ihdr),
		chunk('IDAT', zlib.deflateSync(Buffer.concat(flat))),
		chunk('IEND', Buffer.alloc(0)),
	])
}

/** An 8-bit indexed PNG: PowerPoint writes one of these for a blank slide. */
function encodeIndexed(rows, palette) {
	const h = rows.length
	const w = rows[0].length
	const flat = []
	for (const row of rows) flat.push(Buffer.from([0]), Buffer.from(row))
	const ihdr = Buffer.alloc(13)
	ihdr.writeUInt32BE(w, 0)
	ihdr.writeUInt32BE(h, 4)
	ihdr[8] = 8
	ihdr[9] = 3
	return Buffer.concat([
		SIGNATURE,
		chunk('IHDR', ihdr),
		chunk('PLTE', Buffer.from(palette.flat())),
		chunk('IDAT', zlib.deflateSync(Buffer.concat(flat))),
		chunk('IEND', Buffer.alloc(0)),
	])
}

const CHECKER = [
	[
		[255, 0, 0],
		[0, 255, 0],
	],
	[
		[0, 0, 255],
		[16, 32, 48],
	],
]

/** Every pixel of `png`, as rows of `[r, g, b]`. */
function pixels(png) {
	const image = decodePng(png)
	return Array.from({ length: image.h }, (_, y) => Array.from({ length: image.w }, (_, x) => image.rgb(x, y)))
}

describe('decodePng', () => {
	test('reads a 2x2 truecolour image back pixel for pixel', () => {
		const image = decodePng(encodeRgb(CHECKER))
		expect([image.w, image.h]).toEqual([2, 2])
		expect(pixels(encodeRgb(CHECKER))).toEqual(CHECKER)
	})

	test('undoes every scanline filter, not just None', () => {
		// Sub, Up, Average and Paeth, one per row of a four-row image. Each predictor reads a
		// different neighbour, so a decoder that confused two of them comes back wrong here and
		// nowhere else.
		const rows = [
			[
				[10, 20, 30],
				[40, 50, 60],
			],
			[
				[11, 22, 33],
				[44, 55, 66],
			],
			[
				[12, 24, 36],
				[48, 60, 72],
			],
			[
				[13, 26, 39],
				[52, 65, 78],
			],
		]
		for (const filter of [1, 2, 3, 4]) {
			expect(
				pixels(
					encodeRgb(
						rows,
						rows.map(() => filter)
					)
				)
			).toEqual(rows)
		}
		expect(pixels(encodeRgb(rows, [0, 1, 3, 4]))).toEqual(rows)
	})

	test('resolves an indexed image through its palette', () => {
		// PowerPoint exports a blank slide as an indexed PNG, so this is not a hypothetical
		// colour type: it is what the 3D-model check sees when the model did not draw.
		const palette = [
			[255, 0, 255],
			[255, 255, 255],
		]
		expect(
			pixels(
				encodeIndexed(
					[
						[0, 1],
						[1, 0],
					],
					palette
				)
			)
		).toEqual([
			[
				[255, 0, 255],
				[255, 255, 255],
			],
			[
				[255, 255, 255],
				[255, 0, 255],
			],
		])
	})

	test('a corrupted filter byte changes the pixels it decodes', () => {
		// The sensitivity check for everything above. `Up` on row 1 predicts from row 0; relabel
		// it `None` and the stored deltas are read as absolute values, so the row comes back
		// different. A decoder that ignored the filter byte would return the same pixels either
		// way and pass every case above this one.
		const good = encodeRgb(CHECKER, [0, 2])
		const bad = Buffer.from(good)

		// The filter bytes live inside the deflated IDAT, so the payload is inflated, edited and
		// re-deflated rather than patched in place.
		const idatStart = bad.indexOf(Buffer.from('IDAT', 'ascii')) + 4
		const idatLen = bad.readUInt32BE(idatStart - 8)
		const raw = zlib.inflateSync(bad.subarray(idatStart, idatStart + idatLen))
		raw[1 + CHECKER[0].length * 3] = 0 // row 1's filter: Up -> None

		const patched = Buffer.concat([
			SIGNATURE,
			bad.subarray(SIGNATURE.length, idatStart - 8),
			chunk('IDAT', zlib.deflateSync(raw)),
			chunk('IEND', Buffer.alloc(0)),
		])
		expect(pixels(patched)).not.toEqual(CHECKER)
	})

	test('refuses a colour type or bit depth it cannot read, rather than guessing', () => {
		const png = encodeRgb(CHECKER)
		const depth16 = Buffer.from(png)
		depth16[SIGNATURE.length + 8 + 8] = 16
		expect(() => decodePng(depth16)).toThrow(/bit depth/)
	})
})

describe('solidPngBase64', () => {
	test('produces a 1x1 PNG of exactly the colour asked for', () => {
		// The preview under the 3D model is deliberately magenta so that a fallback render is
		// unmistakable in the pixels. That argument only holds if the encoder writes the colour
		// it was given.
		const png = Buffer.from(solidPngBase64([255, 0, 255]), 'base64')
		const image = decodePng(png)
		expect([image.w, image.h]).toEqual([1, 1])
		expect(image.rgb(0, 0)).toEqual([255, 0, 255])
	})
})
