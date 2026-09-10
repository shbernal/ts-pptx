// The merged coverage gate's inputs, and the ways they used to pass without comparing anything.
//
// Every comparison in the gate is a `<`. `x < undefined` and `NaN < y` are both false, so a
// misspelled axis key in `coverage-gates.json` or a missing `minimumSlack` passed that axis at
// any coverage, and istanbul's `'Unknown'` for an empty axis crashed on `.toFixed` instead of
// saying what was wrong.

import fs from 'node:fs'
import path from 'node:path'
import { describe, expect, test } from 'vitest'
import { AXES, evaluate } from '../../scripts/coverage-gate.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'

const GATES = { minimumSlack: 1, thresholds: { statements: 93, branches: 83, functions: 97, lines: 95 } }

/** A merged summary with every axis at `pct`, or at the value `overrides` names. */
const summary = (pct, overrides = {}) => ({
	total: Object.fromEntries(
		AXES.map((axis) => [axis, { total: 100, covered: 50, skipped: 0, pct: overrides[axis] ?? pct }])
	),
})

describe('coverage gate', () => {
	test('coverage a full point over every notch passes', () => {
		expect(evaluate(summary(99), GATES).failures).toEqual([])
	})

	test('below a notch fails', () => {
		expect(evaluate(summary(99, { statements: 50 }), GATES).failures).toEqual([
			expect.stringMatching(/^statements: 50\.00 is below its gate of 93/),
		])
	})

	test('above a notch by less than the slack fails', () => {
		expect(evaluate(summary(99, { lines: 95.5 }), GATES).failures).toEqual([
			expect.stringMatching(/^lines: 95\.50 clears its gate of 95 by only 0\.50/),
		])
	})

	test('a misspelled axis key fails, however low that axis is', () => {
		const { statements, ...rest } = GATES.thresholds
		const gates = { ...GATES, thresholds: { ...rest, statement: statements } }
		expect(evaluate(summary(99, { statements: 50 }), gates).failures).toEqual([
			expect.stringContaining('thresholds.statements must be a finite number, got undefined'),
		])
	})

	test('a threshold that is not a number fails', () => {
		const gates = { ...GATES, thresholds: { ...GATES.thresholds, branches: '83' } }
		expect(evaluate(summary(99), gates).failures).toEqual([expect.stringContaining('thresholds.branches')])
	})

	test('a missing minimumSlack fails, even inside the point of slack', () => {
		const { minimumSlack: _, ...gates } = GATES
		expect(evaluate(summary(99, { lines: 95.5 }), gates).failures).toEqual([
			expect.stringContaining('minimumSlack must be a finite number, got undefined'),
		])
	})

	test("istanbul's 'Unknown' for an empty axis fails with a message, not a crash", () => {
		expect(evaluate(summary(99, { functions: 'Unknown' }), GATES).failures).toEqual([
			expect.stringContaining('total.functions.pct must be a number, got "Unknown"'),
		])
	})

	test('an axis missing from the summary fails', () => {
		const { total } = summary(99)
		delete total.branches
		expect(evaluate({ total }, GATES).failures).toEqual([expect.stringContaining('total.branches.pct')])
	})

	test('the committed gates file is well formed', () => {
		const gates = JSON.parse(fs.readFileSync(path.join(ROOT, 'scripts', 'coverage-gates.json'), 'utf8'))
		expect(evaluate(summary(100), gates).failures).toEqual([])
	})
})
