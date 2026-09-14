import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
	coverageGroups,
	formatRatio,
	formatTick,
	ratioDomain,
	ratioPosition,
	ratioSpan,
	ratioTicks,
	shapeComparison,
	timingRatios,
	validitySegments,
	validitySummary,
} from '../../../www/comparison/comparison.ts'

/**
 * The comparison charts' arithmetic, against the committed snapshot.
 *
 * The charts draw what `render.mjs` also writes as text on the same page, from the same snapshot,
 * through code that shares none of its helpers. So beyond the geometry, the one assertion worth
 * the most here is that the two agree: every ratio the timing chart plots appears, formatted the
 * same way, in the table the page prints under it.
 */

const read = (relative) => JSON.parse(readFileSync(new URL(relative, import.meta.url), 'utf8'))
const snapshot = read('../../../scripts/comparison/snapshot.json')
const labels = read('../../../scripts/comparison/groups.json')
const data = shapeComparison(snapshot, labels)
const SUBJECTS = ['ts-pptx', 'pptxgenjs']

describe('coverage matrix', () => {
	it('keeps every intent, in snapshot order, under the labels the tables use', () => {
		expect(data.groups.flatMap((group) => group.rows.map((row) => row.id))).toEqual(
			snapshot.coverage.map((row) => row.id)
		)
		for (const group of data.groups) expect(group.label).toBe(labels[group.id])
	})

	it('counts what each library emits per group, which the collapsed baseline row shows', () => {
		for (const subject of SUBJECTS) {
			const perGroup = data.groups.reduce((sum, group) => sum + group.emitted[subject], 0)
			expect(perGroup).toBe(snapshot.coverage.filter((row) => row.results[subject] === 'emitted').length)
		}
	})

	it('titles a group it has no label for, and marks an outcome it does not know as unmeasured', () => {
		const [group] = coverageGroups(
			[{ id: 'probe', label: 'Probe', group: 'widgets', results: { 'ts-pptx': 'emitted', pptxgenjs: 'sideways' } }],
			{}
		)
		expect(group.label).toBe('Widgets')
		expect(group.rows[0].outcomes.pptxgenjs).toBeNull()
		expect(group.emitted['ts-pptx']).toBe(1)
	})
})

describe('validity bars', () => {
	it('spans every intent for both libraries, so a library that built fewer decks shows the gap', () => {
		expect(data.validity.map((bar) => bar.subject)).toEqual(SUBJECTS)
		for (const bar of data.validity) expect(bar.clean + bar.withErrors + bar.notBuilt).toBe(snapshot.coverage.length)
	})

	it('leaves out an empty segment and names what the bar holds', () => {
		/** @type {import('../../../www/comparison/comparison.ts').ValidityBar} */
		const bar = { subject: 'ts-pptx', clean: 21, withErrors: 0, notBuilt: 1, total: 22 }
		expect(validitySegments(bar).map((segment) => segment.key)).toEqual(['clean', 'none'])
		expect(validitySummary(bar)).toBe(
			'ts-pptx: 21 decks with no schema error, 0 with errors, and 1 of 22 intents with no deck'
		)
	})
})

describe('timing ratios', () => {
	const ratios = timingRatios(snapshot.timing.cases)
	const domain = ratioDomain(ratios)

	it('plots every timing case, and prints the same ratios the page table does', () => {
		expect(ratios.map((row) => row.id)).toEqual(snapshot.timing.cases.map((row) => row.id))
		const page = readFileSync(new URL('../../../docs/comparison.md', import.meta.url), 'utf8')
		for (const row of ratios)
			expect(page).toContain(`| ${row.label} | ${formatRatio(row.compressed)} | ${formatRatio(row.stored)} |`)
	})

	it('centres 1× on a log scale wide enough for every ratio', () => {
		expect(ratioPosition(1, domain)).toBeCloseTo(50)
		expect(domain.min * domain.max).toBeCloseTo(1)
		for (const row of ratios)
			for (const value of [row.compressed, row.stored]) {
				expect(ratioPosition(value, domain)).toBeGreaterThan(0)
				expect(ratioPosition(value, domain)).toBeLessThan(100)
			}
	})

	it('ticks at 1× and the two outermost reciprocal pairs that fit', () => {
		expect(ratioTicks({ min: 0.45, max: 2.2 }).map(formatTick)).toEqual(['0.5×', '0.67×', '1×', '1.5×', '2×'])
		expect(ratioTicks({ min: 0.05, max: 20 }).map(formatTick)).toEqual(['0.1×', '0.2×', '1×', '5×', '10×'])
		for (const tick of ratioTicks(domain)) {
			expect(tick).toBeGreaterThanOrEqual(domain.min)
			expect(tick).toBeLessThanOrEqual(domain.max)
		}
	})

	it('joins a row’s two dots whichever side of 1× each one is on', () => {
		const row = { id: 'probe', label: 'Probe', compressed: 1.5, stored: 0.5 }
		const { left, width } = ratioSpan(row, domain)
		expect(left).toBeCloseTo(ratioPosition(0.5, domain))
		expect(left + width).toBeCloseTo(ratioPosition(1.5, domain))
	})
})
