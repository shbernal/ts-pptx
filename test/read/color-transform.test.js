// Read-model coverage for DrawingML colour-transform application
// (applyColorTransforms): base token hex + an ordered transform list → the
// effective rendered hex. This is the step that previously forced the NYCO
// replication agent to hand-compute every lightened/darkened colour.
//
// The main oracle below is read from theme-colors.pptx, a desktop
// PowerPoint-authored deck. Expected effective colours were read back from
// PowerPoint COM (`Shape.Fill.ForeColor.RGB`) after opening the saved fixture.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { applyColorTransforms, Presentation } from '../../dist/read.js'
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

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function open(name) {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', `${name}.pptx`)))
}

function shapeNamed(slide, name) {
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape, `expected shape named ${name}`)
	return shape
}

// shape name | PowerPoint COM effective RGB.
const POWERPOINT_ORACLE = [
	['accent1-plain', 'B01513'],
	['accent1-lm60-lo40', 'ED5654'],
	['accent2-lm75', 'B04A0E'],
	['accent3-shade50', '795F0E'],
	['accent4-tint40', 'A6CDBC'],
	['lt2-lm20-lo80', 'FBFBFB'],
]

describe('applyColorTransforms — PowerPoint fixture oracle', () => {
	for (const [shapeName, expected] of POWERPOINT_ORACLE) {
		test(shapeName, async () => {
			const shape = shapeNamed((await open('theme-colors')).slides[0], shapeName)
			const { hex: base, transforms } = shape.resolvedFill
			const { hex } = applyColorTransforms(base, transforms)
			assertHexClose(hex, expected, 1, shapeName)
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
