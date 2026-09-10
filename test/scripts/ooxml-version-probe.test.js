// The OOXML version probe's verdict on a finished run.
//
// The probe is a manual diagnostic with no CI wiring, so the only thing standing between a
// broken run and a green exit is this function. A target with no result row used to print
// "NO RESULT" and exit 0, and nothing checked that the control row, the one fixture invalid at
// every version, was actually non-zero.

import { describe, expect, test } from 'vitest'
import { verdict } from '../../scripts/ooxml-version-probe.mjs'

/** @type {import('ooxml-validate').FileFormat[]} */
const FORMATS = ['Office2007', 'Office2010', 'Microsoft365']
const TARGETS = [
	{ name: 'base', file: '/t/base.pptx' },
	{ name: 'corruption (control)', file: '/t/control.pptx', control: true },
]

/** A probe report over {@link FORMATS}, one row per `[file, counts]` pair. */
const probe = (rows, violated = false) => ({
	formats: FORMATS,
	rows: rows.map(([file, counts]) => ({ file, counts, regresses: false })),
	violated,
})

describe('ooxml version probe verdict', () => {
	test('a clean row and a non-zero control pass', () => {
		expect(
			verdict(
				TARGETS,
				probe([
					['/t/base.pptx', [0, 0, 0]],
					['/t/control.pptx', [1, 1, 1]],
				])
			)
		).toEqual([])
	})

	test('a target with no result row fails', () => {
		expect(verdict(TARGETS, probe([['/t/control.pptx', [1, 1, 1]]]))).toEqual([
			expect.stringContaining('NO RESULT for "base"'),
		])
	})

	test('a control row with a zero anywhere fails, naming the version', () => {
		expect(
			verdict(
				TARGETS,
				probe([
					['/t/base.pptx', [0, 0, 0]],
					['/t/control.pptx', [0, 1, 1]],
				])
			)
		).toEqual([expect.stringMatching(/control .* no errors at Office2007;/)])
	})

	test('an all-zero table fails, since that is what a validator that never ran reports', () => {
		expect(
			verdict(
				TARGETS,
				probe([
					['/t/base.pptx', [0, 0, 0]],
					['/t/control.pptx', [0, 0, 0]],
				])
			)
		).toEqual([expect.stringContaining('Office2007, Office2010, Microsoft365')])
	})

	test('a count that decreases along the axis fails', () => {
		expect(
			verdict(
				TARGETS,
				probe(
					[
						['/t/base.pptx', [2, 1, 1]],
						['/t/control.pptx', [1, 1, 1]],
					],
					true
				)
			)
		).toEqual([expect.stringContaining('MONOTONICITY BROKEN on "base": Office2007=2 -> Office2010=1')])
	})

	test('a row with the wrong number of counts fails rather than being read against the axis', () => {
		expect(
			verdict(
				TARGETS,
				probe([
					['/t/base.pptx', [0, 0]],
					['/t/control.pptx', [1, 1, 1]],
				])
			)
		).toEqual([expect.stringContaining('"base" has 2 count(s) for 3 format(s)')])
	})

	test('an explicit deck with no control needs no control row', () => {
		expect(verdict([{ name: 'deck.pptx', file: '/t/deck.pptx' }], probe([['/t/deck.pptx', [0, 0, 0]]]))).toEqual([])
	})
})
