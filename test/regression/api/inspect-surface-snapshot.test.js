import { defineRegressionSuite, assert, assertEqual } from '../../helpers.js'
import { buildSnapshot } from '../../../scripts/gen-inspect-snapshot.mjs'
import { readFileSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

// A characterization test: it pins what `inspectPptx()` reports for every fixture
// deck, field by field, without claiming any of it is *right*. The recorded file
// is `inspect-surface.snapshot.json`; regenerate it with
// `node scripts/gen-inspect-snapshot.mjs` and read the diff.
//
// Its job is to make a change to the inspect surface visible. The suite beside it
// (pptx-inspection.test.js) asserts the handful of behaviours that were once bugs
// and explains why each matters; this one covers the other several hundred fields
// nobody would write an assertion for, which is exactly where an "equivalent"
// reimplementation quietly stops being equivalent.

const SNAPSHOT = join(
	dirname(fileURLToPath(import.meta.url)),
	'..',
	'..',
	'read',
	'fixtures',
	'inspect-surface.snapshot.json'
)

/**
 * Inches, ~0.9 EMU. Every box in the snapshot is EMU converted to inches, and EMU
 * is the finest unit PowerPoint itself authors — so a difference this small cannot
 * come from reading a different number out of the file, only from rounding one
 * differently on the way out. Wide enough to permit that, ~1000x tighter than the
 * geometry tolerance the assertions in pptx-inspection.test.js use.
 */
const EPSILON = 1e-6

/** Deep-compare, returning the JSON path of the first difference or null. */
function firstDifference(actual, expected, path = '$') {
	if (typeof expected === 'number' && typeof actual === 'number') {
		if (actual === expected) return null
		if (Number.isFinite(actual) && Number.isFinite(expected) && Math.abs(actual - expected) <= EPSILON) return null
		return `${path}: expected ${expected}, got ${actual}`
	}
	if (Array.isArray(expected) || Array.isArray(actual)) {
		if (!Array.isArray(expected) || !Array.isArray(actual))
			return `${path}: expected ${kind(expected)}, got ${kind(actual)}`
		if (actual.length !== expected.length) return `${path}: expected ${expected.length} entries, got ${actual.length}`
		for (let i = 0; i < expected.length; i++) {
			const diff = firstDifference(actual[i], expected[i], `${path}[${i}]`)
			if (diff) return diff
		}
		return null
	}
	if (expected !== null && typeof expected === 'object') {
		if (actual === null || typeof actual !== 'object') return `${path}: expected an object, got ${kind(actual)}`
		const keys = [...new Set([...Object.keys(expected), ...Object.keys(actual)])]
		for (const key of keys) {
			if (!(key in expected)) return `${path}.${key}: unexpected field (value ${JSON.stringify(actual[key])})`
			if (!(key in actual)) return `${path}.${key}: missing field (expected ${JSON.stringify(expected[key])})`
			const diff = firstDifference(actual[key], expected[key], `${path}.${key}`)
			if (diff) return diff
		}
		return null
	}
	return actual === expected ? null : `${path}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
}

function kind(value) {
	if (value === null) return 'null'
	return Array.isArray(value) ? 'an array' : typeof value
}

defineRegressionSuite('Inspect surface snapshot', [
	{
		name: 'inspectPptx reports the recorded result for every fixture deck',
		fn: async () => {
			const expected = JSON.parse(readFileSync(SNAPSHOT, 'utf8'))
			const actual = await buildSnapshot()

			// Deck-by-deck first: one deck's diff is legible, a 43-deck diff is not.
			for (const deck of Object.keys(expected)) {
				const diff = firstDifference(actual[deck], expected[deck], `${deck}`)
				assert(
					!diff,
					`${deck} no longer inspects the way the snapshot records.\n  ${diff}\n` +
						'  If the change was intended, regenerate: node scripts/gen-inspect-snapshot.mjs'
				)
			}
			// The snapshot only gates the decks it holds; a fixture added without
			// re-running the generator would sit outside the net entirely.
			assertEqual(
				Object.keys(actual).join(' '),
				Object.keys(expected).join(' '),
				'the snapshot covers exactly the fixture decks on disk'
			)
		},
	},
])
