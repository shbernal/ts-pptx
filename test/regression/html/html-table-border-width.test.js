import { describe, test } from 'vitest'
import { assert } from '../../helpers.js'
import { htmlBorderToProps } from '../../../src/gen/table/html-dom.ts'

// Acceptance: HTML-table conversion must preserve FRACTIONAL CSS border widths.
// A hairline border such as `0.5px` previously went through `Math.round(...)` and
// collapsed to `0pt`, so the border silently vanished even though the table
// serializer (`valToPts`) emits fractional points fine.
// A computed width of exactly `0` must instead yield `{ type: 'none' }` (no zero-width line).

describe('HTML table fractional border width', () => {
	test('sub-1px width is preserved, not rounded to zero', () => {
		const b = htmlBorderToProps('0.5px', 'rgb(102, 102, 102)')
		assert(b.width === 0.5, `expected width=0.5, got ${JSON.stringify(b)}`)
		assert(b.color === '666666', `expected color 666666, got ${b.color}`)
		assert(b.type === undefined, `solid border must not set type:none; got ${JSON.stringify(b)}`)
	})

	test('0.4px (would round to 0) is preserved', () => {
		const b = htmlBorderToProps('0.4px', 'rgb(0, 0, 0)')
		assert(b.width === 0.4, `expected width=0.4, got ${JSON.stringify(b)}`)
	})

	test('fractional width above 1px is preserved', () => {
		const b = htmlBorderToProps('2.5px', 'rgb(255, 51, 153)')
		assert(b.width === 2.5, `expected width=2.5, got ${JSON.stringify(b)}`)
		assert(b.color === 'FF3399', `expected color FF3399, got ${b.color}`)
	})

	test('zero width yields {type:none} with no width', () => {
		const b = htmlBorderToProps('0px', 'rgb(102, 102, 102)')
		assert(b.type === 'none', `expected type:none, got ${JSON.stringify(b)}`)
		assert(b.width === undefined, `zero-width border must not set width; got ${JSON.stringify(b)}`)
	})

	test('empty / non-finite width yields {type:none}', () => {
		const b = htmlBorderToProps('', 'rgb(0, 0, 0)')
		assert(b.type === 'none', `expected type:none for empty width, got ${JSON.stringify(b)}`)
	})
})
