// The round trip's aligner, and the rule the two copies disagreed about.
//
// Both the slide calls and the chrome masters have to be paired before anything is diffed, and
// both sides wrote the pairing out. They differed in one respect, and `diffChrome`'s own doc
// records why its version is the correct one: a mutation that removed layout-title
// deduplication left two layouts sharing a title, and the round trip came back clean. So the
// key has to identify exactly one item on BOTH sides. The call aligner counted the actual side
// only, and that reasoning was never carried back.
//
// The whole fixture corpus pairs identically either way -- no fixture has a name that is
// unique among the expected items and duplicated among the actual ones -- so the third case
// below is the evidence that the tightened rule is a rule rather than a no-op.

import { describe, test, expect } from 'vitest'
import { alignByKey } from '../../../src/script/verify/align.ts'

/** `[key, key]` pairs, with `null` where one side had nothing. */
const keys = (pairs) => pairs.map(([before, after]) => [before?.k ?? null, after?.k ?? null])

const item = (k) => ({ k })

describe('alignByKey', () => {
	test('a key pairs its item wherever it moved to', () => {
		// The whole point: position alone would report two mismatches for a reorder.
		const expected = [item('a'), item('b')]
		const actual = [item('b'), item('a')]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			['a', 'a'],
			['b', 'b'],
		])
	})

	test('an item with no counterpart takes the next unclaimed slot, and the tail runs out', () => {
		// The positional fallback is a cursor, not a gap-filler: an expected item whose key names
		// nothing on the other side consumes the next unclaimed actual item, so the shift lands at
		// the end rather than beside the item that went missing. Pre-existing, and the same in both
		// aligners this replaced.
		const expected = [item('a'), item('b'), item('c')]
		const actual = [item('a'), item('c')]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			['a', 'a'],
			['b', 'c'],
			['c', null],
		])
	})

	test('a key duplicated on EITHER side is not usable', () => {
		// `x` is unique among the expected items and duplicated among the actual ones. Keying on
		// it would pair the first of the duplicates and let the second pass unexamined -- which is
		// exactly the shape of the layout-title case that came back clean when it should not have.
		// Only `y`, unique on both sides, aligns by key.
		const expected = [item('x'), item('y')]
		const actual = [item('y'), item('x'), item('x')]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			['x', 'y'],
			['y', 'x'],
			[null, 'x'],
		])
	})

	test('an unkeyed item aligns positionally', () => {
		const expected = [item('a'), item(null)]
		const actual = [item('a'), item(null)]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			['a', 'a'],
			[null, null],
		])
	})

	test('an extra actual item is reported as added, after every expected pairing', () => {
		expect(keys(alignByKey([item('a')], [item('a'), item('b')], (i) => i.k))).toEqual([
			['a', 'a'],
			[null, 'b'],
		])
	})
})
