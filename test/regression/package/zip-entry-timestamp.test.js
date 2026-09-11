// Zip entry timestamps are the same bytes in every timezone.
//
// Every entry is stamped with one fixed modification time so a written deck is reproducible. A
// zip's DOS date and time are wall-clock fields, and fflate fills them from the local getters, so
// a fixed UTC instant came out as different header bytes under different `TZ` values: the same
// deck written in Los Angeles and in Tokyo differed, and west of UTC its entries were dated
// 2000-12-31.
//
// `TZ` is read when a process starts, so each timezone gets its own child `node`. It is spawned
// through `process.execPath` directly rather than a package-manager shim, which on Windows would
// need a shell.

import { execFileSync } from 'node:child_process'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { describe, expect, test } from 'vitest'

const NODE_ENTRY = pathToFileURL(
	path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'dist', 'node.js')
).href

/**
 * Write a one-slide deck under `tz` and return the first entry's local-header time and date bytes
 * (offsets 10 to 13), as hex.
 * @param {string} tz
 * @returns {string}
 */
function headerTimeAndDate(tz) {
	const script = [
		`const { default: TsPptx } = await import(${JSON.stringify(NODE_ENTRY)})`,
		'const pres = new TsPptx()',
		"pres.addSlide().addText('hi', { x: 1, y: 1, w: 2, h: 1 })",
		'const bytes = await pres.toBytes()',
		"process.stdout.write(Buffer.from(bytes.subarray(10, 14)).toString('hex'))",
	].join('\n')
	return execFileSync(process.execPath, ['--input-type=module', '-e', script], {
		env: { ...process.env, TZ: tz },
		encoding: 'utf8',
	})
}

describe('zip entry timestamps', () => {
	test('are identical in two timezones either side of UTC', () => {
		const losAngeles = headerTimeAndDate('America/Los_Angeles')
		const tokyo = headerTimeAndDate('Asia/Tokyo')
		expect(losAngeles).toBe(tokyo)
		// 00:00:00 on 2001-01-01: DOS time 0x0000, DOS date (21 << 9) | (1 << 5) | 1 = 0x2a21,
		// both little-endian.
		expect(tokyo).toBe('0000212a')
	})
})
