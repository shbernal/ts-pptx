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

	test('an item with no counterpart is reported lost, and does not consume a keyed slot', () => {
		// The positional fallback used to be a cursor with no reservation, so `b` consumed `c`
		// positionally and `c` then found its own match already claimed -- two cross-matched pairs
		// where one item was simply missing. Reserving the keyed matches first puts the loss where
		// it happened.
		const expected = [item('a'), item('b'), item('c')]
		const actual = [item('a'), item('c')]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			['a', 'a'],
			['b', null],
			['c', 'c'],
		])
	})

	test('an unkeyed expected item does not steal a slot a later keyed one will claim', () => {
		// The two-element case: `expected[0]` has no key, `expected[1]` is named "A", and the
		// actual side has them the other way round. Without a reservation the unnamed item took
		// "A" positionally, "A" fell through to the unnamed one, and the report described two
		// shapes that each had an exact counterpart as a full set of differences.
		const expected = [item(null), item('A')]
		const actual = [item('A'), item(null)]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			[null, null],
			['A', 'A'],
		])
	})

	test('a key duplicated on EITHER side is not usable', () => {
		// `x` is unique among the expected items and duplicated among the actual ones. Keying on
		// it would pair the first of the duplicates and let the second pass unexamined -- which is
		// exactly the shape of the layout-title case that came back clean when it should not have.
		// `x` therefore aligns positionally, past the `y` reserved for the expected `y`, and the
		// surplus `x` is still reported as added: the mutation this guards still fails.
		const expected = [item('x'), item('y')]
		const actual = [item('y'), item('x'), item('x')]
		expect(keys(alignByKey(expected, actual, (i) => i.k))).toEqual([
			['x', 'x'],
			['y', 'y'],
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
