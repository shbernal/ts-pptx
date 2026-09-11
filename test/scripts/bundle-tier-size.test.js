// The half of the tier gate that can be wrong without esbuild: the programs it bundles.
//
// Bundling is not exercised here — it costs seconds per tier and what it proves is
// esbuild's, not ours. What is ours is that each tier really is a superset of the one
// before it (a `full` that lost its chart call would measure a cheaper program and read as
// a win). The verdict it draws from a number is shared with the other size gate and tested
// in `ratchet-utils.test.js`.

import { describe, expect, test } from 'vitest'
import { programFor } from '../../scripts/bundle-tier-size.mjs'

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

	test('the composed rows are a different program, and name only the families they compose', () => {
		const core = programFor('composed')
		const withCharts = programFor('composed-charts')

		// Not `new TsPptx()`: the whole point of the row is that the class, and the families it names,
		// are never reached.
		expect(core).toContain('createPresentation({ use: [] })')
		expect(core).not.toContain('TsPptx')
		expect(core).not.toContain('dist/families.js')
		expect(core).not.toContain('addChart(')

		// One family added, and the same calls plus the one it enables -- so the difference between
		// the two rows is that family and nothing else.
		expect(withCharts).toContain('import { charts }')
		expect(withCharts).toContain('dist/families.js')
		expect(withCharts).toContain('createPresentation({ use: [charts] })')
		expect(withCharts).toContain('addChart(')
		expect(withCharts).toContain('addText(')
	})

	// The export is parked on a global on purpose: a discarded result is dead code a
	// minifier may drop the whole call graph behind, which would turn every row into a
	// measurement of side-effect annotations.
	test('every program consumes what it writes', () => {
		for (const tier of ['composed', 'composed-charts', 'text', 'text-shape-image', 'full'])
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
