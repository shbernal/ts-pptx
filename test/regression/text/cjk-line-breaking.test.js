// CJK line breaking in the measured-fit tokenizer (src/measure/text-fit.ts).
//
// Latin text breaks only at whitespace, but Chinese and Japanese text breaks
// between any two characters, which is how PowerPoint lays those scripts out. The
// tokenizer used to treat one long CJK run as a single unbreakable "word": harmless
// while the run sits alone on its line (the over-long-word character-wrap fallback
// packs it the same way), but wrong once it follows other content — the whole run
// moved to the next line as a block, wasting the rest of the current line,
// over-reporting the wrapped line count and so shrinking text that actually fits
// (a false vertical overflow).
//
// Hangul is the counter-case, and the reason "CJK" is the wrong word for the rule:
// PowerPoint does not break Korean between syllables, because Korean is written
// with spaces between words. Breaking it per syllable would *under*-report the line
// count, which is the direction that overflows.
//
// What PowerPoint actually does is pinned by the authored deck in
// `test/read/cjk-line-breaking-oracle.test.js`; this file is the arithmetic. Synthetic
// monospace metrics keep it exact: at 18 pt every code point advances 9 raw pt,
// ×WIDTH_SAFETY_FACTOR (1.03) = 9.27 pt laid out, so a 2-inch inner box (144 pt)
// holds ⌊144 / 9.27⌋ = 15 characters per line.
import { describe, test, expect } from 'vitest'
import { measureText } from '../../../src/measure/fit.ts'
import { FontMetricsRegistry } from '../../../src/measure/font-metrics.ts'
import { isCjkBreakCharacter } from '../../../src/measure/text-fit.ts'

const mono = (emPerChar = 0.5) => ({
	unitsPerEm: 1000,
	advanceWidthPt(text, sizePt, charSpacingPt = 0) {
		const n = [...text].length
		return n * emPerChar * sizePt + n * charSpacingPt
	},
	hasCodepoint: () => true,
})

const reg = new FontMetricsRegistry()
reg.set('Mono', mono())

const OPTS = { wIn: 2, fontSize: 18, fontFace: 'Mono' }
// Laid-out advance of one code point: 18 pt × 0.5 em × WIDTH_SAFETY_FACTOR.
const CHAR_PT = 18 * 0.5 * 1.03

describe('CJK line breaking', () => {
	test('a long CJK run wraps at character boundaries', () => {
		// 20 Han characters: 15 fill the first line, 5 spill to the second.
		const m = measureText(reg, '一二三四五六七八九十一二三四五六七八九十', OPTS)
		expect(m.lineCount).toBe(2)
		expect(m.widestLineIn * 72).toBeCloseTo(15 * CHAR_PT, 6)
	})

	test('a CJK run between words fills the current line instead of moving wholesale', () => {
		// 10 Latin characters, then 5 Han characters, then 10 more Latin
		// characters, space-separated. Per-character breaking lays line 1 as
		// `aaaaaaaaaa 一二三四` (10 × 9.27 + space + 4 × 9.27 = 139.05 ≤ 144)
		// and line 2 as `五 bbbbbbbbbb` — 2 lines. Treating the Han run as one
		// unbreakable word pushes it wholly to line 2 (it does not fit beside
		// the first word), which pushes the last word to a third line and
		// over-reports the height — the false vertical overflow.
		const m = measureText(reg, 'aaaaaaaaaa 一二三四五 bbbbbbbbbb', OPTS)
		expect(m.lineCount).toBe(2)
		expect(m.widestLineIn * 72).toBeCloseTo(15 * CHAR_PT, 6)
	})

	test('Kana breaks like Han', () => {
		// Same shape as the mixed case above, in Hiragana.
		expect(measureText(reg, 'aaaaaaaaaa あいうえお bbbbbbbbbb', OPTS).lineCount).toBe(2)
	})

	test('astral CJK (Extension B) breaks per code point, not per UTF-16 unit', () => {
		// Five surrogate pairs — `.length` is 10, but only 5 advances are charged,
		// so this lays out exactly like the 5-character Han case above.
		const extB = '\u{20000}\u{20001}\u{20002}\u{20003}\u{20004}'
		expect(extB.length).toBe(10)
		const m = measureText(reg, 'aaaaaaaaaa ' + extB + ' bbbbbbbbbb', OPTS)
		expect(m.lineCount).toBe(2)
		expect(m.widestLineIn * 72).toBeCloseTo(15 * CHAR_PT, 6)
	})

	test('COUNTER-CASE: a Hangul run does NOT break between syllables', () => {
		// Same shape again, in Hangul. PowerPoint moves the run down whole and
		// spends a third line on it (autofit-cjk-wrap.pptx, cjk__hangul_between_words),
		// so the model must too. Breaking it per syllable would report 2 lines,
		// under-report the height, and let the text overflow.
		expect(measureText(reg, 'aaaaaaaaaa 가나다라마 bbbbbbbbbb', OPTS).lineCount).toBe(3)
	})

	test('a Hangul run longer than the line still breaks, via the over-long-token fallback', () => {
		// 20 syllables with no spaces: too wide for one line, so `countLines` packs
		// it character by character. That is the fallback every over-long token gets,
		// not a break class — which is exactly why Hangul needs no entry in one.
		const m = measureText(reg, '가나다라마바사아자차카타파하가나다라마바', OPTS)
		expect(m.lineCount).toBe(2)
		expect(m.widestLineIn * 72).toBeCloseTo(15 * CHAR_PT, 6)
	})

	test('an all-Latin paragraph is unaffected', () => {
		// Three 10-character words: only one word plus its trailing space fits
		// per line, so this wraps to 3 lines exactly as it did before CJK
		// handling existed.
		const m = measureText(reg, 'aaaaaaaaaa bbbbbbbbbb cccccccccc', OPTS)
		expect(m.lineCount).toBe(3)
		expect(m.widestLineIn * 72).toBeCloseTo(11 * CHAR_PT, 6)
	})
})

describe('isCjkBreakCharacter: range boundaries', () => {
	const breaks = (ch) => isCjkBreakCharacter(ch)

	test('breaks per character', () => {
		for (const cp of [
			0x2e80, // CJK Radicals Supplement, first
			0x3001, // ideographic comma
			0x3042, // Hiragana
			0x30ab, // Katakana
			0x4e00, // CJK Unified Ideographs, first
			0x9fff, // CJK Unified Ideographs, last
			0x3400, // Extension A
			0xf900, // Compatibility Ideographs
			0xfe30, // CJK Compatibility Forms
			0xff21, // fullwidth Latin A
			0xff76, // halfwidth Katakana
			0x1b000, // Kana Supplement
			0x20000, // Extension B
			0x323af, // Extension I, last
		])
			expect(breaks(String.fromCodePoint(cp)), cp.toString(16)).toBe(true)
	})

	test('does not break per character', () => {
		for (const cp of [
			0x0041, // 'A'
			0x0020, // space (whitespace, claimed before this predicate runs)
			0x1100, // Hangul Jamo
			0x3130, // Hangul Compatibility Jamo
			0xac00, // Hangul Syllables, first
			0xd7af, // Hangul Syllables, last
			0xffa0, // halfwidth Hangul jamo filler
			0xffdc, // halfwidth Hangul jamo, last
			0x2e7f, // just below CJK Radicals Supplement
			0xa000, // Yi Syllables, above the Hangul block
			0x1f600, // emoji
			0x10000, // Plane 1 (Linear B)
			0x323b0, // just above Extension I
		])
			expect(breaks(String.fromCodePoint(cp)), cp.toString(16)).toBe(false)
	})
})
