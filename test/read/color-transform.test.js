// Read-model coverage for DrawingML colour-transform application
// (applyColorTransforms): base token hex + an ordered transform list → the
// effective rendered hex. This is the step that previously forced the NYCO
// replication agent to hand-compute every lightened/darkened colour.
//
// The oracle table below is the set of source→effective mappings documented in
// the replication plan and verified against PowerPoint/LibreOffice output. The
// implementation is correct iff it reproduces them within ±1 per channel (the
// small slack absorbs PowerPoint's luminance-rounding quirks).

import { describe, test } from 'vitest'
import { applyColorTransforms } from '../../dist/read.js'
import { assert } from '../helpers.js'

/** Parse `RRGGBB` → [r,g,b] 0–255. */
function channels(hex) {
	const h = hex.startsWith('#') ? hex.slice(1) : hex
	return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}

/** Assert two 6-hex colours match within ±tol per channel. */
function assertHexClose(actual, expected, tol, label) {
	const a = channels(actual)
	const e = channels(expected)
	const ok = a.every((v, i) => Math.abs(v - e[i]) <= tol)
	assert(ok, `${label}: expected ~${expected} (±${tol}/channel), got ${actual}`)
}

/** `[name, valuePerMille]` pairs → the read-model transform shape. */
function tf(...pairs) {
	return pairs.map(([name, value]) => ({ name, value: String(value) }))
}

// base | transforms | effective — straight from the replication oracle.
const ORACLE = [
	['accent1 lumMod60/lumOff40', '04F06A', tf(['lumMod', 60000], ['lumOff', 40000]), '62FCA5'],
	['accent3 lumMod40/lumOff60', '250F6B', tf(['lumMod', 40000], ['lumOff', 60000]), '9377EC'],
	['bg2 lumMod20/lumOff80', '451DC7', tf(['lumMod', 20000], ['lumOff', 80000]), 'D8CEF8'],
	['bg2 lumMod40/lumOff60', '451DC7', tf(['lumMod', 40000], ['lumOff', 60000]), 'B09DF1'],
	['bg2 lumMod60/lumOff40', '451DC7', tf(['lumMod', 60000], ['lumOff', 40000]), '896BEA'],
	['bg2 lumMod75', '451DC7', tf(['lumMod', 75000]), '341695'],
	['bg2 shade30/satMod115', '451DC7', tf(['shade', 30000], ['satMod', 115000]), '20047B'],
	['bg2 shade67.5/satMod115', '451DC7', tf(['shade', 67500], ['satMod', 115000]), '330CB2'],
	['bg2 shade100/satMod115', '451DC7', tf(['shade', 100000], ['satMod', 115000]), '3E10D4'],
	['bg1 lumMod50', 'FFFFFF', tf(['lumMod', 50000]), '7F7F7F'],
]

describe('applyColorTransforms — replication oracle', () => {
	for (const [label, base, transforms, expected] of ORACLE) {
		test(label, () => {
			const { hex } = applyColorTransforms(base, transforms)
			assertHexClose(hex, expected, 1, label)
		})
	}

	test('no transforms returns the base hex unchanged', () => {
		assert(applyColorTransforms('451DC7', []).hex === '451DC7', 'empty transform list is identity')
	})

	test('an unparseable base hex is returned unchanged', () => {
		assert(applyColorTransforms('phClr', tf(['lumMod', 50000])).hex === 'phClr', 'non-hex base passes through')
	})

	test('alpha is reported alongside the hex, not folded into RGB', () => {
		const { hex, alpha } = applyColorTransforms('451DC7', tf(['alpha', 50000]))
		assert(hex === '451DC7', 'alpha leaves RGB untouched')
		assert(alpha === 0.5, `alpha 50000 → 0.5, got ${alpha}`)
	})

	test('alphaMod multiplies the running opacity', () => {
		const { alpha } = applyColorTransforms('451DC7', tf(['alpha', 50000], ['alphaMod', 50000]))
		assert(alpha === 0.25, `alpha 0.5 then alphaMod 0.5 → 0.25, got ${alpha}`)
	})

	test('no alpha transform leaves alpha undefined', () => {
		assert(applyColorTransforms('451DC7', tf(['lumMod', 75000])).alpha === undefined, 'alpha absent unless set')
	})

	test('null/empty transform values are skipped', () => {
		const { hex } = applyColorTransforms('451DC7', [{ name: 'lumMod', value: null }])
		assert(hex === '451DC7', 'a valueless modifier is a no-op')
	})

	test('lumOff clamps luminance at 1 (white)', () => {
		const { hex } = applyColorTransforms('451DC7', tf(['lumOff', 100000]))
		assertHexClose(hex, 'FFFFFF', 1, 'lumOff 100% saturates to white')
	})
})
