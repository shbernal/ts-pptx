// Deterministic unit tests for the unregistered-font fallback (src/font-metrics.ts
// getHeuristicFontMetrics) and the shared table column-width resolver
// (src/gen-utils.ts resolveTableColWidthsEmu). Both are pure — no font files needed.
import { describe, test, expect } from 'vitest'
import { getHeuristicFontMetrics } from '../../src/font-metrics.ts'
import { resolveTableColWidthsEmu } from '../../src/gen-utils.ts'

const EMU_PER_IN = 914400

describe('font-metrics: heuristic fallback', () => {
	const h = getHeuristicFontMetrics()

	test('empty string is zero width', () => {
		expect(h.advanceWidthPt('', 18)).toBe(0)
	})

	test('width scales linearly with font size', () => {
		const at12 = h.advanceWidthPt('Hello World', 12)
		const at24 = h.advanceWidthPt('Hello World', 24)
		expect(at24).toBeCloseTo(at12 * 2, 6)
	})

	test('longer text is wider; wide glyphs beat narrow ones', () => {
		expect(h.advanceWidthPt('WWWW', 18)).toBeGreaterThan(h.advanceWidthPt('iiii', 18))
		expect(h.advanceWidthPt('aaaaa', 18)).toBeGreaterThan(h.advanceWidthPt('aaaa', 18))
	})

	test('charSpacing widens by spacing × char count', () => {
		const base = h.advanceWidthPt('abcd', 18)
		const spaced = h.advanceWidthPt('abcd', 18, 2)
		expect(spaced).toBeCloseTo(base + 4 * 2, 6)
	})

	test('over-estimates a typical sentence (conservative for shrink/resize)', () => {
		// Biased wide: average advance for mixed-case English should exceed ~0.45em/char.
		const text = 'The quick brown fox jumps over the lazy dog'
		const emPerChar = h.advanceWidthPt(text, 100) / 100 / text.length
		expect(emPerChar).toBeGreaterThan(0.45)
	})

	test('non-Latin code points are treated as full-em (safe for CJK)', () => {
		expect(h.advanceWidthPt('一二三', 10)).toBeCloseTo(3 * 1.0 * 10, 6)
	})
})

describe('gen-utils: resolveTableColWidthsEmu', () => {
	test('even distribution splits the EMU width across columns', () => {
		const cols = resolveTableColWidthsEmu(undefined, 9 * EMU_PER_IN, 3)
		expect(cols).toEqual([3 * EMU_PER_IN, 3 * EMU_PER_IN, 3 * EMU_PER_IN])
	})

	test('explicit colW array is per-column inches → EMU', () => {
		const cols = resolveTableColWidthsEmu([2, 3, 4], 9 * EMU_PER_IN, 3)
		expect(cols).toEqual([2 * EMU_PER_IN, 3 * EMU_PER_IN, 4 * EMU_PER_IN])
	})

	test('non-finite colW slot falls back to the even width', () => {
		const even = Math.round((6 * EMU_PER_IN) / 2)
		const cols = resolveTableColWidthsEmu([NaN, 4], 6 * EMU_PER_IN, 2)
		expect(cols[0]).toBe(even)
		expect(cols[1]).toBe(4 * EMU_PER_IN)
	})

	test('zero/absent width falls back to 1in columns rather than 0', () => {
		const cols = resolveTableColWidthsEmu(undefined, 0, 2)
		expect(cols).toEqual([EMU_PER_IN, EMU_PER_IN])
	})

	test('zero columns yields an empty array', () => {
		expect(resolveTableColWidthsEmu(undefined, EMU_PER_IN, 0)).toEqual([])
	})
})
