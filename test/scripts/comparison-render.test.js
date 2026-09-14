// The porting page's differences table: which lines of two arms differ, and how they pair up.
//
// The table is where a reader moving code across looks first, and a diff that pairs the wrong
// lines would put a correct call beside an unrelated one and call that the difference. These pin
// the alignment, which is the part that goes wrong quietly: after an inserted line, a line-by-line
// comparison would mark every later line as changed.

import { describe, expect, test } from 'vitest'
import { lineHunks } from '../../scripts/comparison/render.mjs'

describe('lineHunks', () => {
	test('identical arms have no hunks', () => {
		expect(lineHunks('a\nb\nc', 'a\nb\nc')).toEqual([])
	})

	test('one changed line is one hunk holding both versions', () => {
		expect(lineHunks('a\n\tborder: { width: 1 }\nc', 'a\n\tborder: { pt: 1 }\nc')).toEqual([
			{ ours: ['border: { width: 1 }'], upstream: ['border: { pt: 1 }'] },
		])
	})

	test('a line only one arm has does not shift the comparison of the lines after it', () => {
		const ours = "addChart(DATA, {\n\ttype: 'bar',\n\tx: 1,\n})\naddNotes('n')"
		const upstream = "addChart('bar', DATA, {\n\tx: 1,\n})\naddNotes('n')"
		expect(lineHunks(ours, upstream)).toEqual([
			{ ours: ['addChart(DATA, {', "type: 'bar',"], upstream: ["addChart('bar', DATA, {"] },
		])
	})

	test('separate changes stay separate hunks', () => {
		expect(lineHunks('a\nx\nb\ny', 'a\nX\nb\nY')).toEqual([
			{ ours: ['x'], upstream: ['X'] },
			{ ours: ['y'], upstream: ['Y'] },
		])
	})
})
