import { describe, expect, it } from 'vitest'
import { slideList, summarizeNotes } from '../../../www/demos/deck-preview.ts'

/**
 * The demos page's pure helpers.
 *
 * They exist so the component is markup plus assignments — `.vue` files are read by no
 * typechecker and by no test runner here — and this file is the other half of that
 * bargain. Nothing below touches a DOM, a browser or a deck: what is asserted is the
 * grouping and the prose, which are the two things that would go wrong quietly.
 */

/** One fidelity row, with only the fields the helpers read. */
function note(slide, construct, disposition = 'dropped', cause = 'unread', detail = 'because') {
	return { slide, construct, disposition, cause, detail }
}

describe('summarizeNotes', () => {
	it('collapses one construct raised on many slides into a single row', () => {
		// The real shape: a footer field note fires on every slide carrying a footer, and
		// eleven copies of one sentence reads as eleven problems rather than one.
		const rows = summarizeNotes([note(2, 'text.field'), note(4, 'text.field'), note(5, 'text.field')])

		expect(rows).toHaveLength(1)
		expect(rows[0].construct).toBe('text.field')
		expect(rows[0].slides).toEqual([2, 4, 5])
	})

	it('keeps constructs apart when they differ in disposition or cause', () => {
		const rows = summarizeNotes([
			note(1, 'chart.workbook', 'approximated', 'unsupported'),
			note(1, 'chart.workbook', 'dropped', 'unsupported'),
			note(1, 'chart.workbook', 'dropped', 'unread'),
		])

		// Same construct three times, three different declarations about it. Grouping on the
		// construct alone would report one, and would have to pick one of the three details.
		expect(rows).toHaveLength(3)
		expect(rows.map((row) => `${row.disposition}/${row.cause}`)).toEqual([
			'approximated/unsupported',
			'dropped/unsupported',
			'dropped/unread',
		])
	})

	it('does not repeat a slide that raised the same note twice', () => {
		const rows = summarizeNotes([note(3, 'shape.placeholder'), note(3, 'shape.placeholder')])

		expect(rows[0].slides).toEqual([3])
	})

	it('returns nothing for a deck that declared nothing', () => {
		expect(summarizeNotes([])).toEqual([])
	})
})

describe('slideList', () => {
	it('reads as prose rather than as an array', () => {
		expect(slideList([3, 5, 7])).toBe('3, 5 and 7')
		expect(slideList([2, 4])).toBe('2 and 4')
		expect(slideList([9])).toBe('9')
	})

	it('sorts numerically, not lexically', () => {
		// `[2, 10].sort()` is `[10, 2]`, and a deck of eleven slides is exactly where that
		// shows up.
		expect(slideList([10, 2, 11])).toBe('2, 10 and 11')
	})

	it('is empty for no slides, rather than throwing or saying "undefined"', () => {
		expect(slideList([])).toBe('')
	})
})
