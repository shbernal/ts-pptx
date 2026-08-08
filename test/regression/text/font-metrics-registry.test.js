// Font-metrics provider + registry through the public `ts-pptx/measure` subpath
// (dist/measure.js). These primitives are the standalone surface a consumer uses to
// build its own resolver/registry (docs/measured-text-fit.md); the registry's
// coverage/fallback methods (`hasFace`, `hasCodepoint`, variant fallback in `get`)
// are not touched by the export-time bake, so they only get dist coverage here.
// Silkscreen (OFL, committed under test/read/fixtures/fonts) gives a real cmap:
// ASCII covered, emoji/CJK not — so `hasCodepoint` exercises both truth values.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, test, expect } from 'vitest'
import { parseFontMetrics, getHeuristicFontMetrics, FontMetricsRegistry } from '../../../dist/measure.js'

const REG_BYTES = new Uint8Array(
	readFileSync(fileURLToPath(new URL('../../read/fixtures/fonts/Silkscreen-Regular.ttf', import.meta.url)))
)
const BOLD_BYTES = new Uint8Array(
	readFileSync(fileURLToPath(new URL('../../read/fixtures/fonts/Silkscreen-Bold.ttf', import.meta.url)))
)

describe('parseFontMetrics → FontMetrics', () => {
	test('exposes unitsPerEm and sums raw advances (no shaping)', async () => {
		const fm = await parseFontMetrics(REG_BYTES)
		expect(fm.unitsPerEm).toBeGreaterThan(0)
		expect(fm.advanceWidthPt('', 12)).toBe(0) // empty string short-circuits to 0
		const one = fm.advanceWidthPt('A', 12)
		expect(one).toBeGreaterThan(0)
		expect(fm.advanceWidthPt('AA', 12)).toBeCloseTo(2 * one, 6) // raw sum, linear
	})

	test('charSpacingPt is added once per code point', async () => {
		const fm = await parseFontMetrics(REG_BYTES)
		const base = fm.advanceWidthPt('ABC', 12)
		expect(fm.advanceWidthPt('ABC', 12, 2)).toBeCloseTo(base + 3 * 2, 6)
	})

	test('hasCodepoint reads the cmap: ASCII covered, emoji/CJK not', async () => {
		const fm = await parseFontMetrics(REG_BYTES)
		expect(fm.hasCodepoint(0x41)).toBe(true) // 'A'
		expect(fm.hasCodepoint(0x1f600)).toBe(false) // 😀
		expect(fm.hasCodepoint(0x4e2d)).toBe(false) // 中
	})

	test('rejects a module without parse()', async () => {
		// Sanity: parse() is required; a garbage buffer throws from opentype, not silently.
		await expect(parseFontMetrics(new Uint8Array([0, 1, 2, 3]))).rejects.toThrow()
	})
})

describe('getHeuristicFontMetrics', () => {
	test('is a shared singleton', () => {
		expect(getHeuristicFontMetrics()).toBe(getHeuristicFontMetrics())
	})

	test('advanceWidthPt: empty is 0, non-empty scales with size and adds charSpacing', () => {
		const h = getHeuristicFontMetrics()
		expect(h.advanceWidthPt('', 18)).toBe(0)
		const w = h.advanceWidthPt('Mixed CASE 123', 18)
		expect(w).toBeGreaterThan(0)
		expect(h.advanceWidthPt('AB', 18, 5)).toBeCloseTo(h.advanceWidthPt('AB', 18) + 2 * 5, 6)
		// Wider glyphs advance more than narrow ones (heuristic ratio table).
		expect(h.advanceWidthPt('W', 18)).toBeGreaterThan(h.advanceWidthPt('i', 18))
	})

	test('hasCodepoint always reports covered (no cmap to consult)', () => {
		const h = getHeuristicFontMetrics()
		expect(h.hasCodepoint(0x41)).toBe(true)
		expect(h.hasCodepoint(0x1f600)).toBe(true)
	})
})

describe('FontMetricsRegistry', () => {
	test('set/get with exact → regular → any-variant fallback', async () => {
		const reg = new FontMetricsRegistry()
		const regular = await parseFontMetrics(REG_BYTES)
		const bold = await parseFontMetrics(BOLD_BYTES)
		expect(reg.size).toBe(0)
		reg.set('Silkscreen', regular)
		reg.set('Silkscreen', bold, { bold: true })
		expect(reg.size).toBe(2)
		expect(reg.get('Silkscreen')).toBe(regular) // exact regular
		expect(reg.get('Silkscreen', true)).toBe(bold) // exact bold
		expect(reg.get('Silkscreen', false, true)).toBe(regular) // italic unregistered → regular fallback
	})

	test('any-variant fallback when only a non-regular weight is registered', async () => {
		const reg = new FontMetricsRegistry()
		const bold = await parseFontMetrics(BOLD_BYTES)
		reg.set('Silkscreen', bold, { bold: true })
		// No regular registered → italic lookup falls through to the sole bold variant.
		expect(reg.get('Silkscreen', false, true)).toBe(bold)
	})

	test('get returns undefined for an unnamed or unregistered face', async () => {
		const reg = new FontMetricsRegistry()
		reg.set('Silkscreen', await parseFontMetrics(REG_BYTES))
		expect(reg.get(undefined)).toBeUndefined()
		expect(reg.get('')).toBeUndefined()
		expect(reg.get('Nope')).toBeUndefined()
	})

	test('hasFace is case-insensitive and per-face', async () => {
		const reg = new FontMetricsRegistry()
		reg.set('Silkscreen', await parseFontMetrics(REG_BYTES))
		expect(reg.hasFace('Silkscreen')).toBe(true)
		expect(reg.hasFace('SILKSCREEN')).toBe(true)
		expect(reg.hasFace('Other')).toBe(false)
	})

	test('hasCodepoint: registered → boolean from cmap, unregistered → undefined', async () => {
		const reg = new FontMetricsRegistry()
		reg.set('Silkscreen', await parseFontMetrics(REG_BYTES))
		expect(reg.hasCodepoint('Silkscreen', 0x41)).toBe(true)
		expect(reg.hasCodepoint('Silkscreen', 0x1f600)).toBe(false)
		expect(reg.hasCodepoint('Unregistered', 0x41)).toBeUndefined() // unknown, not "covered"
		expect(reg.hasCodepoint(null, 0x41)).toBeUndefined()
	})
})
