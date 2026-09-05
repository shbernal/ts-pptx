import { describe, test } from 'vitest'
import { assert } from '../../helpers.js'
import { htmlBorderToProps } from '../../../src/gen/table/html-dom.ts'

// Two properties of the HTML-table border read, and they pull against each other.
//
// The unit: `StrokeProps.width` is POINTS and a computed CSS border width is PX, so the
// magnitude is converted at 96px/in, not copied. A `1px` border is 0.75pt. The copy-across
// reading made every `tableToSlides` border a third too thick, and it survived two earlier
// corrections of the same mistake elsewhere in the file because each was made by hand.
//
// The precision: a hairline such as `0.5px` must not round to `0pt` and silently vanish -- the
// table serializer (`ptsToEmuLenient`) emits fractional points fine. Converting makes the
// smallest widths smaller still, so this is the property the conversion could most easily have
// broken. A computed width of exactly `0` is a different statement and yields `{ type: 'none' }`.

/** px→pt at the CSS reference pixel, spelled out rather than imported, so the test states the ratio. */
const pt = (px) => (px * 72) / 96

describe('HTML table border width is converted to points', () => {
	test('96px is exactly 72pt -- the ratio itself, not a rounded decimal', () => {
		const b = htmlBorderToProps('96px', 'rgb(0, 0, 0)')
		assert(b.width === 72, `expected width=72pt for 96px, got ${JSON.stringify(b)}`)
	})

	test('sub-1px width is converted and preserved, not rounded to zero', () => {
		const b = htmlBorderToProps('0.5px', 'rgb(102, 102, 102)')
		assert(b.width === pt(0.5), `expected width=${pt(0.5)}pt, got ${JSON.stringify(b)}`)
		assert(b.color === '666666', `expected color 666666, got ${b.color}`)
		assert(b.type === undefined, `solid border must not set type:none; got ${JSON.stringify(b)}`)
	})

	test('0.4px (would round to 0) is converted and preserved', () => {
		const b = htmlBorderToProps('0.4px', 'rgb(0, 0, 0)')
		assert(b.width === pt(0.4), `expected width=${pt(0.4)}pt, got ${JSON.stringify(b)}`)
	})

	test('fractional width above 1px is converted and preserved', () => {
		const b = htmlBorderToProps('2.5px', 'rgb(255, 51, 153)')
		assert(b.width === pt(2.5), `expected width=${pt(2.5)}pt, got ${JSON.stringify(b)}`)
		assert(b.color === 'FF3399', `expected color FF3399, got ${b.color}`)
	})

	test('a bare number is read as px, like every other length in this file', () => {
		const b = htmlBorderToProps('3', 'rgb(0, 0, 0)')
		assert(b.width === pt(3), `expected width=${pt(3)}pt for a bare 3, got ${JSON.stringify(b)}`)
	})

	test('zero width yields {type:none} with no width', () => {
		const b = htmlBorderToProps('0px', 'rgb(102, 102, 102)')
		assert(b.type === 'none', `expected type:none, got ${JSON.stringify(b)}`)
		assert(b.width === undefined, `zero-width border must not set width; got ${JSON.stringify(b)}`)
	})

	test('a negative width yields {type:none} rather than a negative line', () => {
		const b = htmlBorderToProps('-1px', 'rgb(0, 0, 0)')
		assert(b.type === 'none', `expected type:none for a negative width, got ${JSON.stringify(b)}`)
	})

	for (const value of ['', 'auto', '3em', '30%'])
		test(`${JSON.stringify(value)} has no absolute length and yields {type:none}`, () => {
			const b = htmlBorderToProps(value, 'rgb(0, 0, 0)')
			assert(b.type === 'none', `expected type:none for ${JSON.stringify(value)}, got ${JSON.stringify(b)}`)
		})
})
