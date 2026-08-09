import { spawn } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { parseArgs } from 'node:util'

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

/**
 * Was this module run directly, rather than imported?
 *
 * The gate scripts keep their logic in exported functions and their CLI behind this,
 * so a test can import the parsing and scanning without the script measuring `dist/`,
 * writing a budget file, or calling `process.exit` out from under the test runner.
 * `gen-inspect-snapshot.mjs` open-coded this first; it is shared now that it has
 * more than two callers.
 * @param {string} metaUrl the caller's `import.meta.url`
 * @returns {boolean}
 */
export function isMain(metaUrl) {
	const entry = process.argv[1]
	return entry !== undefined && path.resolve(entry) === fileURLToPath(metaUrl)
}

/**
 * Thrown by {@link parseCli} when the caller should stop and exit with `code`, having
 * already printed everything the user needs. Lets a `main()` keep one exit path
 * instead of threading a "did we print usage?" flag through its early returns.
 */
export class CliExit extends Error {
	/** @param {number} code */
	constructor(code) {
		super('cli exit ' + code)
		this.name = 'CliExit'
		this.code = code
	}
}

/**
 * `node:util`'s `parseArgs`, plus the two things every script here wants around it:
 * a `--help`/`-h` that prints usage and stops, and an unknown flag that reports itself
 * as one line rather than an eight-frame `ERR_PARSE_ARGS_UNKNOWN_OPTION` stack.
 *
 * The hand-rolled alternative this replaces was `argv.indexOf('--flag') + 1`, which
 * silently takes the *next flag* as the value when the argument is omitted —
 * `--dir --verbose` set the directory to `"--verbose"` and reported nothing.
 * `options` is `parseArgs`'s own option map (`{type: 'string'|'boolean', short?, default?}`
 * per flag). It is typed loosely here rather than with `ParseArgsOptionConfig`, which
 * `@types/node` does not export; the shape is `parseArgs`'s to validate at runtime anyway.
 * @param {string[]} argv arguments, already sliced past the script name
 * @param {{options: Record<string, any>, usage: string, allowPositionals?: boolean}} config
 * @returns {{values: Record<string, any>, positionals: string[]}}
 * @throws {CliExit} on `--help` (code 0) or a malformed argument list (code 2)
 */
export function parseCli(argv, { options, usage, allowPositionals = false }) {
	// Cast because spreading widens each `type` to `string`, which no longer matches
	// `ParseArgsOptionsType`. The shape is parseArgs's to validate at runtime.
	const withHelp = /** @type {any} */ ({ ...options, help: { type: 'boolean', short: 'h', default: false } })
	/** @type {{values: Record<string, any>, positionals: string[]}} */
	let parsed
	try {
		parsed = parseArgs({ args: argv, options: withHelp, allowPositionals })
	} catch (error) {
		console.error(error instanceof Error ? error.message : String(error))
		console.error('\n' + usage)
		throw new CliExit(2)
	}
	if (parsed.values.help) {
		console.log(usage)
		throw new CliExit(0)
	}
	return parsed
}

/**
 * {@link parseCli} for a script that runs at module top level and has no `main()` to
 * catch a {@link CliExit}. Exits the process on `--help` or a bad argument list instead
 * of throwing, so those paths end quietly rather than as an unhandled rejection.
 * @param {string[]} argv
 * @param {{options: any, usage: string, allowPositionals?: boolean}} config
 * @returns {{values: Record<string, any>, positionals: string[]}}
 */
export function parseCliOrExit(argv, config) {
	try {
		return parseCli(argv, config)
	} catch (error) {
		if (error instanceof CliExit) process.exit(error.code)
		throw error
	}
}

/**
 * Run a `main()` that may throw {@link CliExit}, setting `process.exitCode` from either
 * the return value or the exit request. Keeps the `isMain` tail of each script to one line.
 * @param {() => number | Promise<number>} main
 */
export async function runCli(main) {
	try {
		process.exitCode = await main()
	} catch (error) {
		if (error instanceof CliExit) {
			process.exitCode = error.code
			return
		}
		console.error(error instanceof Error ? (error.stack ?? error.message) : String(error))
		process.exitCode = 1
	}
}

const packageManagerCache = process.env.TSPPTX_SCRIPT_CACHE_DIR || path.join(ROOT, '.tmp', 'package-manager-cache')

const requireFromRoot = createRequire(path.join(ROOT, 'package.json'))

// Bin names owned by a local devDependency, mapped to the package that declares them.
// These get resolved to their JS entry and run on the current node binary, so Windows
// never has to exec a .bin/*.CMD shim (spawn refuses .cmd without a shell).
const localBinPackages = {
	attw: '@arethetypeswrong/cli',
	publint: 'publint',
}

function resolveLocalBin(name) {
	const pkg = localBinPackages[name]
	if (!pkg) return null
	let manifestPath
	try {
		manifestPath = requireFromRoot.resolve(pkg + '/package.json')
	} catch {
		return null
	}
	const { bin } = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
	const entry = typeof bin === 'string' ? bin : bin?.[name]
	if (!entry) return null
	return path.resolve(path.dirname(manifestPath), entry)
}

function quoteArg(arg) {
	return /[\s"^&|<>()]/.test(arg) ? '"' + arg.replace(/"/g, '\\"') + '"' : arg
}

export function run(command, args, options = {}) {
	return new Promise((resolve, reject) => {
		const env = {
			...process.env,
			npm_config_cache: path.join(packageManagerCache, 'npm'),
			NPM_CONFIG_CACHE: path.join(packageManagerCache, 'npm'),
			...options.env,
		}
		if (command === 'pnpm') {
			env.pnpm_config_store_dir = path.join(packageManagerCache, 'pnpm-store')
			env.PNPM_CONFIG_STORE_DIR = path.join(packageManagerCache, 'pnpm-store')
		}
		/** @type {import('node:child_process').SpawnOptions} */
		const spawnOptions = {
			cwd: options.cwd || ROOT,
			env,
			stdio: options.capture ? ['ignore', 'pipe', 'pipe'] : 'inherit',
		}
		const localBin = resolveLocalBin(command)
		let child
		if (localBin) {
			child = spawn(process.execPath, [localBin, ...args], spawnOptions)
		} else if (process.platform === 'win32' && !path.isAbsolute(command)) {
			// pnpm/npm are .cmd shims on Windows: they need a shell, and Node 24 deprecates
			// passing args alongside shell:true, so hand the shell one pre-quoted line.
			const line = [command + '.cmd', ...args].map(quoteArg).join(' ')
			child = spawn(line, { ...spawnOptions, shell: true })
		} else {
			child = spawn(command, args, spawnOptions)
		}
		let stdout = ''
		let stderr = ''
		if (child.stdout)
			child.stdout.on('data', (chunk) => {
				stdout += chunk
			})
		if (child.stderr)
			child.stderr.on('data', (chunk) => {
				stderr += chunk
			})
		child.on('error', reject)
		child.on('close', (code) => {
			if (code === 0) resolve({ stdout, stderr })
			else reject(new Error(command + ' ' + args.join(' ') + ' exited with code ' + code + '\n' + (stderr || stdout)))
		})
	})
}

// `pnpm pack` helpers used to live here. They moved to `pack-utils.mjs`: only the two
// package-boundary gates call them, while nearly every other importer of this module
// wants `ROOT` and nothing else.
