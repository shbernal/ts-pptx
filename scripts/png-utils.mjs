#!/usr/bin/env node
/**
 * A minimal PNG encoder and decoder, for the gates that have to look at pixels.
 *
 * Both halves lived inside `powerpoint-com-smoke.mjs`, an 848-line COM driver, which is the last
 * place a reader would look for a zlib-inflating scanline unfilter -- and the last place a
 * reviewer would think to check one. `scripts/README.md`'s own rule is that a gate with parsing
 * logic exports it and guards its CLI behind `isMain`; this is that parsing logic, and it now has
 * a case in `test/scripts/` rather than being exercised only as a side effect of driving desktop
 * PowerPoint on Windows. It is the kind of code that fails by returning plausible pixels.
 *
 * Hand-rolled rather than pulled in as a dependency: two gate scripts are the only consumers,
 * PowerPoint writes whichever 8-bit colour type suits the slide (truecolour for a rendered model,
 * an indexed palette for a blank one), and `node:zlib` already does the hard part.
 *
 * Not in `script-utils.mjs` or `pack-utils.mjs`: both have deliberately narrow audiences stated
 * in their own headers.
 */

import zlib from 'node:zlib'

/**
 * A 1x1 solid-colour PNG, built here so nothing about the preview can be mistaken for model pixels.
 * @param {readonly number[]} rgb - the fill colour
 */
export function solidPngBase64(rgb) {
	const raw = Buffer.from([0, rgb[0] ?? 0, rgb[1] ?? 0, rgb[2] ?? 0])
	const crcTable = Int32Array.from({ length: 256 }, (_, n) => {
		let c = n
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		return c
	})
	/** @param {Uint8Array} buf */
	const crc = (buf) => {
		let c = -1
		for (const byte of buf) c = (crcTable[(c ^ byte) & 0xff] ?? 0) ^ (c >>> 8)
		return (c ^ -1) >>> 0
	}
	/**
	 * @param {string} type four-character PNG chunk type
	 * @param {Buffer} data
	 */
	const chunk = (type, data) => {
		const len = Buffer.alloc(4)
		len.writeUInt32BE(data.length)
		const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
		const crcBuf = Buffer.alloc(4)
		crcBuf.writeUInt32BE(crc(body))
		return Buffer.concat([len, body, crcBuf])
	}
	const ihdr = Buffer.alloc(13)
	ihdr.writeUInt32BE(1, 0)
	ihdr.writeUInt32BE(1, 4)
	ihdr[8] = 8 // bit depth
	ihdr[9] = 2 // colour type: truecolour
	return Buffer.concat([
		Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
		chunk('IHDR', ihdr),
		chunk('IDAT', zlib.deflateSync(raw)),
		chunk('IEND', Buffer.alloc(0)),
	]).toString('base64')
}

/**
 * Decode an 8-bit PNG (truecolour, palette, or grey, +/- alpha) into `{w, h, rgb(x,y)}`.
 *
 * Hand-rolled rather than pulled in as a dependency: this script is the only consumer, PowerPoint
 * writes whichever of those colour types suits the slide (truecolour for the rendered model, an
 * indexed palette for a blank one), and `node:zlib` already does the hard part.
 * @param {Buffer} bytes - the PNG file
 */
export function decodePng(bytes) {
	/** Indexed read that satisfies `noUncheckedIndexedAccess`; out-of-range is 0, as for a `Buffer`. */
	const at = (/** @type {Uint8Array} */ buf, /** @type {number} */ i) => buf[i] ?? 0
	let off = 8
	let w = 0
	let h = 0
	let colour = 0
	/** @type {Buffer | null} */
	let palette = null
	/** @type {Buffer[]} */
	const idat = []
	while (off + 8 <= bytes.length) {
		const len = bytes.readUInt32BE(off)
		const type = bytes.toString('ascii', off + 4, off + 8)
		const data = bytes.subarray(off + 8, off + 8 + len)
		if (type === 'IHDR') {
			w = data.readUInt32BE(0)
			h = data.readUInt32BE(4)
			if (at(data, 8) !== 8) throw new Error(`unsupported PNG bit depth ${at(data, 8)}`)
			colour = at(data, 9)
			if (at(data, 12) !== 0) throw new Error('interlaced PNG not supported')
		} else if (type === 'PLTE') palette = Buffer.from(data)
		else if (type === 'IDAT') idat.push(data)
		else if (type === 'IEND') break
		off += 12 + len
	}
	const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colour]
	if (!channels) throw new Error(`unsupported PNG colour type ${colour}`)
	if (colour === 3 && !palette) throw new Error('indexed PNG with no PLTE chunk')
	const pal = palette ?? Buffer.alloc(0)
	const raw = zlib.inflateSync(Buffer.concat(idat))
	const stride = w * channels
	const out = Buffer.alloc(h * stride)
	let pos = 0
	for (let y = 0; y < h; y++) {
		const filter = at(raw, pos++)
		const line = raw.subarray(pos, pos + stride)
		pos += stride
		const cur = out.subarray(y * stride, (y + 1) * stride)
		const prior = y > 0 ? out.subarray((y - 1) * stride, y * stride) : Buffer.alloc(stride)
		for (let i = 0; i < stride; i++) {
			const a = i >= channels ? at(cur, i - channels) : 0
			const b = at(prior, i)
			const c = i >= channels ? at(prior, i - channels) : 0
			let v = at(line, i)
			if (filter === 1) v += a
			else if (filter === 2) v += b
			else if (filter === 3) v += (a + b) >> 1
			else if (filter === 4) {
				const p = a + b - c
				const pa = Math.abs(p - a)
				const pb = Math.abs(p - b)
				const pc = Math.abs(p - c)
				v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c
			}
			cur[i] = v & 0xff
		}
	}
	/**
	 * @param {number} x
	 * @param {number} y
	 * @returns {[number, number, number]}
	 */
	const rgb = (x, y) => {
		const i = y * stride + x * channels
		if (colour === 3) {
			const p = at(out, i) * 3
			return [at(pal, p), at(pal, p + 1), at(pal, p + 2)]
		}
		if (colour === 0 || colour === 4) return [at(out, i), at(out, i), at(out, i)]
		return [at(out, i), at(out, i + 1), at(out, i + 2)]
	}
	return { w, h, rgb }
}
