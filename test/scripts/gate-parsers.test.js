// The gate parsers whose failure mode is a plausible answer rather than an error.
//
// Each of these decides what a check *is*, so a regression in one does not fail loudly — it
// quietly runs less. `run-steps.mjs`'s expansion already produced one vacuous gate (a
// `docs:check` elided by the de-duplication), `path-refs.mjs`'s resolver was rewritten for a
// Windows/Linux case-sensitivity split with nothing pinning either half, and
// `sync-version.mjs`'s whole point is its "exactly one match" requirement, `ensure-dist.mjs`
// stands between the whole suite and a stale `dist/`, and `coverage-project.mjs` decides which
// of the browser lane's hits the merged report is allowed to count.

import { describe, expect, test } from 'vitest'
import path from 'node:path'
import fs from 'node:fs/promises'
import os from 'node:os'
import { expand } from '../../scripts/run-steps.mjs'
import { resolves } from '../../scripts/path-refs.mjs'
import { replaceVersion } from '../../scripts/sync-version.mjs'
import { stale } from '../../scripts/ensure-dist.mjs'
import { hitsByLocation, project } from '../../scripts/coverage-project.mjs'
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

describe('ensure-dist staleness', () => {
	// The unacceptable answer is `null` on a stale tree: it runs the whole suite against
	// yesterday's build and reports a pass. Every case here is built in a tmpdir with mtimes
	// set explicitly, because "newer than" is the only thing the guard actually decides.
	const OLD = new Date('2020-01-01T00:00:00Z')
	const NEW = new Date('2020-01-02T00:00:00Z')

	/** A tree holding the inputs and outputs `ensure-dist` looks at, each at a chosen mtime. */
	async function tree({ inputsAt = OLD, outputsAt = NEW, omit = [] } = {}) {
		const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ensure-dist-'))
		await fs.mkdir(path.join(root, 'src', 'gen'), { recursive: true })
		await fs.mkdir(path.join(root, 'dist'), { recursive: true })
		const inputs = [
			'src/index.ts',
			'src/gen/deep.ts',
			'tsdown.config.ts',
			'tsconfig.base.json',
			'tsconfig.json',
			'package.json',
			'pnpm-lock.yaml',
		]
		for (const f of inputs) {
			await fs.writeFile(path.join(root, f), 'x')
			await fs.utimes(path.join(root, f), inputsAt, inputsAt)
		}
		for (const f of ['dist/index.js', 'dist/index.d.ts'].filter((name) => !omit.includes(name))) {
			await fs.writeFile(path.join(root, f), 'x')
			await fs.utimes(path.join(root, f), outputsAt, outputsAt)
		}
		return root
	}

	/** Move one file's mtime past `NEW`, i.e. past every build output. */
	const touch = (root, file) => fs.utimes(path.join(root, file), NEW, new Date(NEW.getTime() + 1000))

	test('a build newer than every input is current', async () => {
		expect(await stale(await tree())).toBe(null)
	})

	test('a source file newer than the build is stale, and the reason names src/', async () => {
		const root = await tree()
		await touch(root, 'src/index.ts')
		expect(await stale(root)).toMatch(/build input is newer than dist\/: src\//)
	})

	test('a source file NESTED under src/ counts too', async () => {
		// The walk is recursive, and a guard that only stated the top level would call a tree
		// with an edited `src/gen/**` fresh - which is most of this repo.
		const root = await tree()
		await touch(root, 'src/gen/deep.ts')
		expect(await stale(root)).toMatch(/build input is newer than dist\//)
	})

	test('a newer build config is stale, and the reason names the file', async () => {
		const root = await tree()
		await touch(root, 'pnpm-lock.yaml')
		expect(await stale(root)).toMatch(/pnpm-lock\.yaml/)
	})

	test('a missing output is stale and says which, rather than reading as fresh', async () => {
		// `mtimeOf` returns 0 for an absent file, so an absent output is the *oldest* rather
		// than a missing one and the tree reads stale either way. What this pins is the
		// reason, which is what tells a reader the build never ran at all.
		const root = await tree({ omit: ['dist/index.d.ts'] })
		expect(await stale(root)).toBe('missing build output: dist/index.d.ts')
	})

	test('the OLDEST output is what inputs are compared against', async () => {
		// Not the newest: a partial rebuild that refreshed `index.js` and left `index.d.ts`
		// behind is stale, and comparing against the newest output would call it current.
		const root = await tree()
		await fs.utimes(path.join(root, 'dist', 'index.d.ts'), OLD, new Date(OLD.getTime() - 1000))
		expect(await stale(root)).toMatch(/build input is newer than dist\//)
	})
})

describe('coverage projection onto the Node report shape', () => {
	/** One file's istanbul coverage over N single-statement lines. */
	const fileData = (lines, hits) => ({
		path: '/dist/x.js',
		statementMap: Object.fromEntries(
			lines.map((line, i) => [String(i), { start: { line, column: 0 }, end: { line, column: 10 } }])
		),
		fnMap: {},
		branchMap: {},
		s: Object.fromEntries(hits.map((h, i) => [String(i), h])),
		f: {},
		b: {},
	})

	test('a hit at a matching location is carried onto the Node report’s own index', () => {
		// The two indices are deliberately reversed: the browser lane numbers its statements in
		// its own order, so a projection that copied by index rather than by location would pass
		// a same-shaped case and fail this one.
		const { coverage, orphans } = project(fileData([10, 20], [0, 0]), fileData([20, 10], [7, 0]))
		expect(coverage.s).toEqual({ 0: 0, 1: 7 })
		expect(orphans).toBe(0)
	})

	test('the Node report’s maps are the ones kept, so the denominator cannot move', () => {
		const node = fileData([10, 20], [0, 0])
		const { coverage } = project(node, fileData([10], [3]))
		expect(coverage.statementMap).toBe(node.statementMap)
		expect(Object.keys(coverage.s)).toEqual(['0', '1'])
	})

	test('a location the Node report has no slot for is counted as an orphan, not merged', () => {
		// Dropping it is the safe direction - it can only under-report - but it has to be
		// *counted*, because the share of them is the gate on whether the two lanes describe the
		// same build at all.
		const { orphans, measured } = project(fileData([10], [0]), fileData([10, 999], [1, 1]))
		expect(orphans).toBe(1)
		expect(measured).toBe(2)
	})

	test('a location the browser never reached projects as zero, not as absent', () => {
		const { coverage } = project(fileData([10, 20], [5, 5]), fileData([10], [0]))
		expect(coverage.s).toEqual({ 0: 0, 1: 0 })
	})

	test('repeated hits at one location are summed across the lane’s scenarios', () => {
		// The lane runs eight scenarios and the same statement appears in several of them.
		const data = fileData([3, 3], [2, 3])
		expect(hitsByLocation(data).get('s:3:0-3:10')).toBe(5)
	})

	test('an open-ended end column compares equal however it is spelled', () => {
		// One side has been through JSON (`null`) and the other has not (`Infinity`). Treating
		// those as two locations is what would double-count a statement.
		const node = fileData([10], [0])
		node.statementMap['0'].end = { line: 10, column: null }
		const browser = fileData([10], [4])
		browser.statementMap['0'].end = { line: 10, column: Infinity }
		expect(project(node, browser).coverage.s).toEqual({ 0: 4 })
	})
})
