// The shared CLI front end every gate script now routes its arguments through.
//
// The form this replaced was `argv.indexOf('--flag') + 1`, which returns the NEXT FLAG
// when the value is missing — `--dir --verbose` silently set the directory to
// `"--verbose"`. That is the case worth pinning: it must be an error, not a value.

import fs from 'node:fs'
import { describe, expect, test, vi } from 'vitest'
import { CliExit, parseCli, resolveLocalBin, runCli } from '../../scripts/script-utils.mjs'

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

// `run()` sends these two bins to the current node binary precisely so Windows never has
// to exec a `.bin/*.CMD` shim, which `spawn` refuses without a shell and which is only on
// PATH when a package manager put it there. The lookup used to go through
// `require.resolve(pkg + '/package.json')`, so any package that leaves `./package.json`
// out of its `exports` map resolved to `null` and silently fell back to that shim —
// `publint` does, and `package:lint` died with "'publint.cmd' is not recognized" whenever
// it ran outside `pnpm run`. Assert the entry exists rather than just that it is non-null:
// a path pointing at nothing fails the same way, one spawn later.
describe('resolveLocalBin', () => {
	test.each(['publint', 'attw'])('%s resolves to a real JS entry, not a .cmd shim', (name) => {
		const entry = resolveLocalBin(name)
		expect(entry).toBeTruthy()
		expect(fs.existsSync(/** @type {string} */ (entry))).toBe(true)
	})

	test('a bin this module does not own is left to the shell', () => {
		expect(resolveLocalBin('pnpm')).toBeNull()
	})
})
