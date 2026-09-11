// The shared CLI front end every gate script now routes its arguments through.
//
// The form this replaced was `argv.indexOf('--flag') + 1`, which returns the NEXT FLAG
// when the value is missing — `--dir --verbose` silently set the directory to
// `"--verbose"`. That is the case worth pinning: it must be an error, not a value.

import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { describe, expect, test, vi } from 'vitest'
import { CliExit, ROOT, collect, isMain, parseCli, resolveLocalBin, runCli } from '../../scripts/script-utils.mjs'

const OPTIONS = { dir: { type: 'string' }, verbose: { type: 'boolean', default: false } }
const parse = (argv) => parseCli(argv, { options: OPTIONS, usage: 'usage: thing [--dir <path>]' })

/** Swallow the usage/error output these paths print on the way out. */
const quiet = () => {
	vi.spyOn(console, 'log').mockImplementation(() => {})
	vi.spyOn(console, 'error').mockImplementation(() => {})
	return () => vi.restoreAllMocks()
}

describe('parseCli', () => {
	test('parses values and applies defaults', () => {
		expect(parse(['--dir', 'decks', '--verbose']).values).toMatchObject({ dir: 'decks', verbose: true })
		expect(parse([]).values).toMatchObject({ verbose: false })
	})

	// The regression. Under the old indexOf form this returned dir === '--verbose'.
	test('a flag whose value is missing is an error, not the next flag', () => {
		const restore = quiet()
		expect(() => parse(['--dir', '--verbose'])).toThrow(CliExit)
		expect(() => parse(['--dir'])).toThrow(CliExit)
		restore()
	})

	test('an unknown flag exits 2 and prints usage rather than a stack trace', () => {
		const restore = quiet()
		expect(() => parse(['--nope'])).toThrow(expect.objectContaining({ code: 2 }))
		expect(console.error).toHaveBeenCalledWith(expect.stringContaining("Unknown option '--nope'"))
		expect(console.error).toHaveBeenCalledWith(expect.stringContaining('usage: thing'))
		restore()
	})

	test('--help and -h exit 0 after printing usage', () => {
		const restore = quiet()
		for (const flag of ['--help', '-h']) {
			expect(() => parse([flag])).toThrow(expect.objectContaining({ code: 0 }))
		}
		expect(console.log).toHaveBeenCalledWith(expect.stringContaining('usage: thing'))
		restore()
	})

	test('positionals are rejected unless the caller opts in', () => {
		const restore = quiet()
		expect(() => parseCli(['extra'], { options: OPTIONS, usage: 'u' })).toThrow(CliExit)
		expect(parseCli(['extra'], { options: OPTIONS, usage: 'u', allowPositionals: true }).positionals).toEqual(['extra'])
		restore()
	})
})

// Node realpaths the main module's URL and leaves `argv[1]` as typed. Comparing the two
// directly made every guarded script a silent exit 0 when run through a junction, a
// symlink or a `subst` drive, and `run-steps.mjs` is guarded, so `verify` passed without
// running a step. Only a real link reproduces it.
describe('isMain', () => {
	test('is true for a script run through a directory link', () => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'is-main-'))
		try {
			const target = path.join(dir, 'real')
			const link = path.join(dir, 'link')
			fs.mkdirSync(target)
			// A junction needs no privilege on win32; elsewhere the type is ignored.
			fs.symlinkSync(target, link, 'junction')
			const utils = pathToFileURL(path.join(ROOT, 'scripts', 'script-utils.mjs')).href
			fs.writeFileSync(
				path.join(target, 'guarded.mjs'),
				`import { isMain } from ${JSON.stringify(utils)}\nif (isMain(import.meta.url)) console.log('ran')\n`
			)
			expect(execFileSync(process.execPath, [path.join(target, 'guarded.mjs')], { encoding: 'utf8' })).toBe('ran\n')
			expect(execFileSync(process.execPath, [path.join(link, 'guarded.mjs')], { encoding: 'utf8' })).toBe('ran\n')
		} finally {
			fs.rmSync(dir, { recursive: true, force: true })
		}
	})

	test('is false for a module that was imported', () => {
		expect(isMain(pathToFileURL(path.join(ROOT, 'scripts', 'script-utils.mjs')).href)).toBe(false)
	})
})

describe('runCli', () => {
	const withExitCode = async (fn) => {
		const before = process.exitCode
		try {
			await fn()
			return process.exitCode
		} finally {
			process.exitCode = before
		}
	}

	test('takes the exit code from the return value', async () => {
		expect(await withExitCode(() => runCli(() => 0))).toBe(0)
		expect(await withExitCode(() => runCli(() => 1))).toBe(1)
	})

	test('takes the exit code from a CliExit', async () => {
		expect(await withExitCode(() => runCli(() => Promise.reject(new CliExit(2))))).toBe(2)
	})

	test('an unexpected throw is reported and exits 1', async () => {
		const restore = quiet()
		expect(
			await withExitCode(() =>
				runCli(() => {
					throw new Error('boom')
				})
			)
		).toBe(1)
		expect(console.error).toHaveBeenCalledWith(expect.stringContaining('boom'))
		restore()
	})
})

// Every script that launches a package's command sends its bin to the current node binary, so
// Windows never has to exec a `.bin/*.CMD` shim, which `spawn` refuses without a shell and which
// is only on PATH when a package manager put it there. The lookup used to go through
// `require.resolve(pkg + '/package.json')`, so any package that leaves `./package.json` out of
// its `exports` map resolved to `null` and silently fell back to that shim — `publint` does, and
// `package:lint` died with "'publint.cmd' is not recognized" whenever it ran outside `pnpm run`.
// Assert the entry exists rather than just that it is non-null: a path pointing at nothing fails
// the same way, one spawn later.
describe('resolveLocalBin', () => {
	test.each([
		['publint', 'publint', null],
		['@arethetypeswrong/cli', 'attw', null],
		['tsdown', 'tsdown', null],
		['typescript', 'tsc', null],
		// TypeDoc is installed by the tools/api-docs workspace package for itself, not at the root.
		['typedoc', 'typedoc', path.join(ROOT, 'tools', 'api-docs')],
	])('%s resolves its %s bin to a real JS entry, not a .cmd shim', (pkg, bin, from) => {
		const entry = resolveLocalBin(pkg, bin, from ? { from } : {})
		expect(entry).toBeTruthy()
		expect(fs.existsSync(/** @type {string} */ (entry))).toBe(true)
		expect(entry).not.toMatch(/\.cmd$/i)
	})

	test('a package that is not installed resolves to null', () => {
		expect(resolveLocalBin('no-such-package-installed-here')).toBeNull()
	})

	test('a bin the package does not declare resolves to null', () => {
		expect(resolveLocalBin('typescript', 'no-such-bin')).toBeNull()
	})
})

// The render and COM smokes report a tool's exit rather than throwing on it, so the helper they
// share must resolve on every outcome: a failing exit, a command that is not there, and a hang.
describe('collect', () => {
	test('reports a non-zero exit and what the process printed, without rejecting', async () => {
		const script = 'process.stdout.write("out"); process.stderr.write("err"); process.exit(3)'
		expect(await collect(process.execPath, ['-e', script])).toEqual({
			code: 3,
			out: 'out',
			err: 'err',
			timedOut: false,
		})
	})

	test('a command that cannot be spawned resolves with code -1 and the error', async () => {
		const result = await collect('no-such-command-installed-here', [])
		expect(result.code).toBe(-1)
		expect(result.err).toMatch(/ENOENT/)
		expect(result.timedOut).toBe(false)
	})

	// The timeout is what stops a hung LibreOffice holding a job for the rest of its budget.
	test('kills a process that outlives its timeout, and says so', async () => {
		const started = Date.now()
		const result = await collect(process.execPath, ['-e', 'setTimeout(() => {}, 60000)'], { timeoutMs: 200 })
		expect(result.timedOut).toBe(true)
		expect(Date.now() - started).toBeLessThan(10000)
	})
})
