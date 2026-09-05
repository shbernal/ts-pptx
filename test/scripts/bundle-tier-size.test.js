// The tier gate's two halves that can be wrong without esbuild: the programs it bundles
// and the verdict it draws from a number.
//
// Bundling is not exercised here — it costs seconds per tier and what it proves is
// esbuild's, not ours. What is ours is that each tier really is a superset of the one
// before it (a `full` that lost its chart call would measure a cheaper program and read as
// a win) and that a measurement sitting exactly on its budget is not a failure.

import { describe, expect, test } from 'vitest'
import { HEADROOM_PCT, SLACK_MIN_BYTES, SLACK_PCT } from '../../scripts/bundle-size-ratchet.mjs'
import { frozenBudget, programFor, verdictFor } from '../../scripts/bundle-tier-size.mjs'

describe('programFor', () => {
	test('each tier keeps every call of the tiers below it', () => {
		const text = programFor('text')
		const middle = programFor('text-shape-image')
		const full = programFor('full')

		expect(text).toContain('addText(')
		expect(middle).toContain('addText(')
		expect(middle).toContain('addShape(')
		expect(middle).toContain('addImage(')
		for (const call of ['addText(', 'addShape(', 'addImage(', 'addChart(', 'addTable(', 'addMedia('])
			expect(full).toContain(call)
	})

	test('the cheap tier reaches no construct family the split is meant to defer', () => {
		const text = programFor('text')
		for (const call of ['addChart(', 'addTable(', 'addMedia(', 'addImage(']) expect(text).not.toContain(call)
	})

	// The export is parked on a global on purpose: a discarded result is dead code a
	// minifier may drop the whole call graph behind, which would turn every row into a
	// measurement of side-effect annotations.
	test('every program consumes what it writes', () => {
		for (const tier of ['text', 'text-shape-image', 'full'])
			expect(programFor(tier)).toMatch(/globalThis\.\w+ = await pres\.write\(/)
	})

	// esbuild would take a Windows absolute path; Node refuses `import('C:/...')`, reading
	// `C:` as a URL scheme. The gate runs each program before measuring it, so the one
	// specifier has to satisfy both.
	test('names the built entry relatively, so Node can run the program too', () => {
		const specifier = /import TsPptx from "([^"]+)"/.exec(programFor('text'))?.[1]
		expect(specifier).toMatch(/^\.\.?\//)
		expect(specifier).toContain('dist/browser.js')
		expect(specifier).not.toMatch(/^[A-Za-z]:/)
	})

	test('an unknown tier names itself rather than bundling nothing', () => {
		expect(() => programFor('charts-only')).toThrow(/no such tier/)
	})
})

describe('verdictFor', () => {
	test('a measurement on its budget passes', () => {
		expect(verdictFor(1000, 1000)).toBe('ok')
	})

	test('one byte over fails', () => {
		expect(verdictFor(1001, 1000)).toBe('over')
	})

	// Both conditions have to hold, which is what keeps the nag off the small figures: a
	// percentage of a tiny number is noise, and `--freeze` rounds up to a whole kB anyway.
	test('a win is only worth banking when it clears both the percentage and the floor', () => {
		const budget = 100 * 1024
		expect(verdictFor(budget * (1 - SLACK_PCT / 100) - 1, budget)).toBe('under')
		expect(verdictFor(budget - SLACK_MIN_BYTES + 1, budget)).toBe('ok')

		const small = SLACK_MIN_BYTES * 2
		expect(verdictFor(small * (1 - SLACK_PCT / 100) - 1, small)).toBe('ok')
	})
})

describe('frozenBudget', () => {
	test('leaves headroom above the measurement and rounds to a whole kB', () => {
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
