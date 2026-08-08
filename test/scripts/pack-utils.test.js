// `pnpm pack --json` output parsing.
//
// The three fallback branches exist because pnpm has moved this output between streams
// and formats across versions; the package gates read a tarball name out of whatever
// comes back. A parse that silently picks the wrong thing takes both gates down with a
// confusing error far from the cause, so the shapes are pinned here.

import { describe, expect, test } from 'vitest'
import { parsePackOutput } from '../../scripts/pack-utils.mjs'

describe('parsePackOutput', () => {
	test('reads an object payload printed after progress lines', () => {
		const output = ['Progress: resolved 12, reused 12', '{"filename":"ts-pptx-1.0.0.tgz"}'].join('\n')
		expect(parsePackOutput(output)).toEqual({ filename: 'ts-pptx-1.0.0.tgz' })
	})

	test('reads an array payload', () => {
		const output = ['packing…', '[{"filename":"ts-pptx-1.0.0.tgz"}]'].join('\n')
		expect(parsePackOutput(output)).toEqual([{ filename: 'ts-pptx-1.0.0.tgz' }])
	})

	test('reads a payload that is the entire output', () => {
		expect(parsePackOutput('{"filename":"a.tgz"}')).toEqual({ filename: 'a.tgz' })
	})

	// The reason the scan anchors on a line start: a progress line may legitimately
	// contain a brace, and taking the first one found would parse noise.
	test('ignores a brace inside a progress line', () => {
		const output = ['Resolving {workspace} links', '{"filename":"b.tgz"}'].join('\n')
		expect(parsePackOutput(output)).toEqual({ filename: 'b.tgz' })
	})

	test('takes the last payload when several are printed', () => {
		const output = ['{"filename":"stale.tgz"}', '{"filename":"fresh.tgz"}'].join('\n')
		expect(parsePackOutput(output)).toEqual({ filename: 'fresh.tgz' })
	})

	test('throws rather than returning nothing when no JSON was printed', () => {
		expect(() => parsePackOutput('ERR_PNPM_NO_PKG  No package.json found')).toThrow(/did not print JSON/)
	})
})
