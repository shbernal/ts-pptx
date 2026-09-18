// Probe reader; produces no committed fixture. Reads the legend order out of the PNGs
// `probe-combo-legend-order.ps1` exports, so the ranking is measured rather than eyeballed.
//
// Each subchart's one series is painted a distinct flat colour by
// `probe-combo-legend-order.mjs` -- the one given FIRST is red, the second blue -- so the legend's
// swatches name their subcharts without anything having to read the labels. Both colours appear in
// the plot too, so only the legend band is scanned: `legendPos: 'b'` puts the legend in the bottom
// strip, and the subcharts are drawn above it.
//
// The answer is which colour's swatch sits further left. That is the whole measurement: if blue
// leads, the second-given subchart was listed first, which means the legend reordered by type.

import fs from 'node:fs'
import path from 'node:path'
import { decodePng } from '../../../../scripts/png-utils.mjs'

const outDir = new URL('../../../../.tmp/combo-legend-order/', import.meta.url)
const dir = path.normalize(outDir.pathname.replace(/^\//, ''))

/** The two colours `probe-combo-legend-order.mjs` paints the first- and second-given subcharts. */
const FIRST = [0xff, 0x00, 0x00]
const SECOND = [0x00, 0x00, 0xff]

/** A tolerance, because PowerPoint antialiases a swatch's edges but not its middle. */
const NEAR = 24
const matches = (rgb, want) =>
	Math.abs(rgb[0] - want[0]) <= NEAR && Math.abs(rgb[1] - want[1]) <= NEAR && Math.abs(rgb[2] - want[2]) <= NEAR

/**
 * The leftmost x at which `want` appears in the image's bottom `band` fraction, or `null`.
 *
 * Counted rather than found on the first hit: a stray antialiased pixel is not a swatch, so a
 * column has to hold a few of them to count. Without that the answer came back as whichever
 * colour's plot line happened to dip lowest.
 */
function leftmostInLegend(img, want, band = 0.18, minRun = 3) {
	const from = Math.floor(img.h * (1 - band))
	for (let x = 0; x < img.w; x++) {
		let hits = 0
		for (let y = from; y < img.h; y++) {
			if (matches(img.rgb(x, y), want)) hits++
		}
		if (hits >= minRun) return x
	}
	return null
}

const rows = []
for (const file of fs
	.readdirSync(dir)
	.filter((name) => name.endsWith('.png'))
	.sort()) {
	const img = decodePng(fs.readFileSync(path.join(dir, file)))
	const first = leftmostInLegend(img, FIRST)
	const second = leftmostInLegend(img, SECOND)
	const name = file.replace(/\.png$/, '')
	if (first === null || second === null) {
		rows.push({ name, order: `UNREADABLE (first=${first}, second=${second})` })
		continue
	}
	rows.push({ name, order: first < second ? 'given order' : 'REORDERED: second listed first', first, second })
}

const width = Math.max(...rows.map((row) => row.name.length))
for (const row of rows) {
	console.log(
		`${row.name.padEnd(width)}  ${row.order}${row.first === undefined ? '' : `  (x ${row.first} vs ${row.second})`}`
	)
}

const refused = path.join(dir, 'refused.txt')
if (fs.existsSync(refused)) {
	console.log('\nrefused by the write path (recorded, not dropped):')
	console.log(fs.readFileSync(refused, 'utf8').trimEnd())
}
