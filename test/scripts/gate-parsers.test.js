// The three gate parsers whose failure mode is a plausible answer rather than an error.
//
// Each of these decides what a check *is*, so a regression in one does not fail loudly — it
// quietly runs less. `run-steps.mjs`'s expansion already produced one vacuous gate (a
// `docs:check` elided by the de-duplication), `path-refs.mjs`'s resolver was rewritten for a
// Windows/Linux case-sensitivity split with nothing pinning either half, and
// `sync-version.mjs`'s whole point is its "exactly one match" requirement.

import { describe, expect, test } from 'vitest'
import path from 'node:path'
import { expand } from '../../scripts/run-steps.mjs'
import { resolves } from '../../scripts/path-refs.mjs'
import { replaceVersion } from '../../scripts/sync-version.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'

describe('run-steps expansion', () => {
	test('a leaf script expands to its own command', () => {
		expect(expand('raw-xml:check')).toEqual([{ step: 'raw-xml:check', command: 'node scripts/raw-xml-ratchet.mjs' }])
	})

	test('an aggregate expands to every leaf beneath it, in order', () => {
		const steps = expand('verify')
		expect(steps.length).toBeGreaterThan(5)
		// `verify` is the gate the repo runs on everything; these are the checks it must contain.
		const names = steps.map((s) => s.step)
		for (const required of ['typecheck', 'raw-xml:check', 'path-refs:check', 'docs:check', 'test']) {
			expect(names, `verify must still run ${required}`).toContain(required)
		}
		// Every entry is a real leaf command, not another `run-steps` invocation: an unexpanded
		// self-reference would run the whole aggregate again inside one of its own steps.
		for (const { command } of steps) expect(command).not.toMatch(/run-steps\.mjs/)
	})

	test('an unknown script name is an error, not an empty plan', () => {
		// The failure this rules out is the quiet one: returning `[]` would make a gate that runs
		// nothing report success.
		expect(() => expand('no-such-script')).toThrow(/no such script/)
	})

	test('a cycle is reported rather than followed', () => {
		// Reached through the real table, so this asserts the guard exists rather than
		// constructing a fake one it would never see.
		expect(() => expand('verify', ['verify'])).toThrow(/script cycle/)
	})
})

describe('path-refs citation resolution', () => {
	const known = new Set(['src/gen/chart/plot-bar.ts', 'src/read/api/shapes/types.ts', 'docs/reference/index.md'])
	const from = path.join(ROOT, 'src', 'gen', 'chart', 'chart-xml.ts')

	test('a repo-root-relative path resolves', () => {
		expect(resolves('src/gen/chart/plot-bar.ts', from, known)).toBe(true)
	})

	test('a path relative to the citing file resolves', () => {
		expect(resolves('./plot-bar.ts', from, known)).toBe(true)
		expect(resolves('../../read/api/shapes/types.ts', from, known)).toBe(true)
	})

	test('a `.js` citation resolves to the `.ts` source it names', () => {
		// A comment citing a sibling module by its emitted `.js` name means the `.ts` it is
		// compiled from, so the resolver tries that extension too.
		expect(resolves('./plot-bar.js', from, known)).toBe(true)
	})

	test('a suffix match needs a path boundary', () => {
		// `types.ts` must not satisfy a citation of `shapes/types.ts`'s tail, or every short
		// filename in the repo would resolve against any longer path ending in it.
		expect(resolves('shapes/types.ts', from, known)).toBe(true)
		expect(resolves('pes/types.ts', from, known)).toBe(false)
	})

	test('resolution is case-exact, on every platform', () => {
		// The whole reason the resolver compares against the walked names rather than calling
		// `stat`: on Windows `stat` says yes to a wrong-case path, so `verify` passed locally and
		// `check:static` failed on ubuntu. A verdict must be a property of the repo, not the machine.
		expect(resolves('src/gen/Chart/plot-bar.ts', from, known)).toBe(false)
		expect(resolves('./Plot-Bar.ts', from, known)).toBe(false)
	})

	test('a path that names nothing does not resolve', () => {
		expect(resolves('src/gen/chart/plot-nothing.ts', from, known)).toBe(false)
	})
})

describe('sync-version replacement', () => {
	const source = (body) => `const x = 1\n${body}\nexport const y = 2\n`

	test('replaces the one version line and reports what was there', () => {
		const { text, previous } = replaceVersion(source("const VERSION = '1.2.3'"), '4.5.6')
		expect(previous).toBe('1.2.3')
		expect(text).toContain("const VERSION = '4.5.6'")
		expect(text).not.toContain('1.2.3')
	})

	test('no match is an error, not a silent no-op', () => {
		// Zero matches means the constant was renamed or moved, and replacing nothing while
		// reporting success is exactly the failure the script exists to remove.
		expect(() => replaceVersion(source("const RELEASE = '1.2.3'"), '4.5.6')).toThrow(/found 0/)
	})

	test('two matches are an error too', () => {
		expect(() => replaceVersion(source("const VERSION = '1.2.3'\nconst VERSION = '1.2.4'"), '4.5.6')).toThrow(/found 2/)
	})
})
