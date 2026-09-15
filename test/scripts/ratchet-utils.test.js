// The ratchet mechanics both size gates share, which can be wrong without a build: the verdict a
// gate draws from a number, the budget a freeze writes, and which keys a budget file and a
// measurement disagree about.

import { describe, expect, test } from 'vitest'
import {
	HEADROOM_PCT,
	SLACK_MIN_BYTES,
	SLACK_PCT,
	budgetKeyDrift,
	frozenBudget,
	verdictFor,
} from '../../scripts/ratchet-utils.mjs'

describe('verdictFor', () => {
	test('a measurement on its budget passes', () => {
		expect(verdictFor(1000, 1000)).toBe('ok')
	})

	test('one byte over fails', () => {
		expect(verdictFor(1001, 1000)).toBe('over')
	})

	// Both conditions have to hold, which is what keeps the nag off the small figures: a
	// percentage of a tiny number is noise, and `--freeze` rounds up to a whole KiB anyway.
	test('a win is only worth banking when it clears both the percentage and the floor', () => {
		const budget = 100 * 1024
		expect(verdictFor(budget * (1 - SLACK_PCT / 100) - 1, budget)).toBe('under')
		expect(verdictFor(budget - SLACK_MIN_BYTES + 1, budget)).toBe('ok')

		const small = SLACK_MIN_BYTES * 2
		expect(verdictFor(small * (1 - SLACK_PCT / 100) - 1, small)).toBe('ok')
	})
})

describe('frozenBudget', () => {
	test('leaves headroom above the measurement and rounds to a whole KiB', () => {
		expect(frozenBudget(100 * 1024)).toBe(Math.ceil(100 * (1 + HEADROOM_PCT / 100)) * 1024)
		expect(frozenBudget(1)).toBe(1024)
	})

	// The property the pair has to have between them, or `--freeze` writes a budget its own
	// `check` immediately fails.
	test('a freshly frozen budget passes its own check', () => {
		for (const bytes of [1, 1024, 40_000, 144_500, 211_000])
			expect(verdictFor(bytes, frozenBudget(bytes))).not.toBe('over')
	})
})

describe('budgetKeyDrift', () => {
	test('a measured key with no budget is missing', () => {
		expect(budgetKeyDrift(['a.js', 'b.js'], { 'a.js': 1024 })).toEqual({ missing: ['b.js'], stale: [] })
	})

	// The half the size gates used to leave unchecked: a row whose entry or tier the gate stopped
	// measuring sat in the budget file with nothing reporting it.
	test('a budgeted key the gate no longer measures is stale', () => {
		expect(budgetKeyDrift(['a.js'], { 'a.js': 1024, 'gone.js': 2048 })).toEqual({ missing: [], stale: ['gone.js'] })
	})

	test('a value that is not a complete budget counts as missing, by the rule the gate passes', () => {
		expect(budgetKeyDrift(['a.js'], { 'a.js': '1024' })).toEqual({ missing: ['a.js'], stale: [] })

		/** @param {{ initial?: number, total?: number } | undefined} figures */
		const bothFigures = (figures) => typeof figures?.initial === 'number' && typeof figures?.total === 'number'
		const tiers = { text: { initial: 1, total: 2 }, full: { initial: 1 } }
		expect(budgetKeyDrift(['text', 'full'], tiers, bothFigures)).toEqual({ missing: ['full'], stale: [] })
	})
})
