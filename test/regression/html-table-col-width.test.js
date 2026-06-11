import { describe, test } from 'vitest'
import { assert } from '../helpers.js'
import { resolveHtmlColWidth } from '../../src/gen-tables.ts'

// Acceptance: HTML-table conversion must honor `data-pptx-width` (exact) and
// `data-pptx-min-width` (floor) overrides, and must keep working when the table is
// hidden. Hidden tables report offsetWidth 0 for every cell, so the proportional
// width is a 0/0 = NaN calc. Previously the `data-pptx-width` value was assigned into
// the min-width variable while the set-width variable stayed a const 0, so an explicit
// width never applied and hidden tables emitted NaN widths (upstream gitbrent/PptxGenJS#1157).

describe('HTML table column width resolution (upstream #1157)', () => {
	test('explicit data-pptx-width wins over the proportional calc', () => {
		assert(resolveHtmlColWidth(3, 5, 0) === 5, 'set width must override calc width')
	})

	test('data-pptx-width wins even over a larger min-width', () => {
		assert(resolveHtmlColWidth(3, 5, 9) === 5, 'set width is exact and beats min-width')
	})

	test('min-width raises the proportional calc when larger', () => {
		assert(resolveHtmlColWidth(2, 0, 4) === 4, 'min-width must act as a floor')
	})

	test('proportional calc is used when no override is present', () => {
		assert(resolveHtmlColWidth(2.5, 0, 0) === 2.5, 'calc width is the default')
	})

	test('min-width below the calc does not shrink it', () => {
		assert(resolveHtmlColWidth(6, 0, 2) === 6, 'min-width must not reduce a wider calc')
	})

	test('hidden table (NaN calc) falls back to the explicit width', () => {
		assert(resolveHtmlColWidth(NaN, 5, 0) === 5, 'hidden table must honor data-pptx-width')
	})

	test('hidden table (NaN calc) falls back to the min-width', () => {
		assert(resolveHtmlColWidth(NaN, 0, 4) === 4, 'hidden table must honor data-pptx-min-width')
	})

	test('hidden table with no override yields 0, never NaN', () => {
		const w = resolveHtmlColWidth(NaN, 0, 0)
		assert(w === 0, `hidden table with no override must be 0, got ${w}`)
	})

	test('NaN overrides (invalid attribute values) are ignored', () => {
		assert(resolveHtmlColWidth(3, NaN, NaN) === 3, 'invalid override values fall back to calc')
	})
})
