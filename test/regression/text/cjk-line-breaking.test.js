// CJK line breaking in the measured-fit tokenizer (src/measure/text-fit.ts).
//
// Latin text breaks only at whitespace, but Chinese/Japanese/Korean text breaks
// between any two characters (UAX #14 class ID for Han/Hangul/Kana), which is how
// PowerPoint lays these scripts out. The tokenizer used to treat one long CJK run
// as a single unbreakable "word": harmless while the run sits alone on its line
// (the over-long-word character-wrap fallback packs it the same way), but wrong
// once it follows other content — the whole run moved to the next line as a block,
// wasting the rest of the current line, over-reporting the wrapped line count and
// so shrinking text that actually fits (a false vertical overflow).
//
// Synthetic monospace metrics keep the arithmetic exact: at 18 pt every code point
// advances 9 raw pt, ×WIDTH_SAFETY_FACTOR (1.03) = 9.27 pt laid out, so a 2-inch
// inner box (144 pt) holds ⌊144 / 9.27⌋ = 15 characters per line.
import { describe, test, expect } from 'vitest'
import { measureText } from '../../../src/measure/fit.ts'
import { FontMetricsRegistry } from '../../../src/measure/font-metrics.ts'

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

	test('Korean and Kana runs break like Han', () => {
		// Same shape as the mixed case above, in Hangul syllables and Kana:
		// a 5-character run between two 10-character Latin words.
		const hangul = '가나다라마'
		const kana = 'あいうえお'
		expect(hangul.length).toBe(5)
		expect(kana.length).toBe(5)
		expect(measureText(reg, 'aaaaaaaaaa ' + hangul + ' bbbbbbbbbb', OPTS).lineCount).toBe(2)
		expect(measureText(reg, 'aaaaaaaaaa ' + kana + ' bbbbbbbbbb', OPTS).lineCount).toBe(2)
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
