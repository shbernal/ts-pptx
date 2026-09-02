import { spawn } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { parseArgs } from 'node:util'

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

/**
 * The read-side fixture corpus: the PowerPoint-authored decks every read, round-trip and
 * census gate runs over.
 *
 * `test/read/corpus.js`'s header states the problem this closes -- "a glob is a claim about
 * how many decks are under test, and a claim spelled four times is one that can quietly become
 * false in one of them" -- and then five more copies lived in `scripts/`, two of them with no
 * empty-corpus guard and one with a different filter (`.pptx` **or** `.potx`, so `template.potx`
 * was in the inspect snapshot's corpus and nobody else's). The enumerator is here rather than in
 * `test/read/corpus.js` because the scripts must not import out of `test/`, and not in
 * `pack-utils.mjs` because that module's header deliberately scopes it to the two package gates.
 */
export const FIXTURES_DIR = path.join(ROOT, 'test', 'read', 'fixtures')

/**
 * Where a `--dir` flag points, defaulting to {@link FIXTURES_DIR}.
 *
 * `resolve`, not `join`: an absolute `--dir` must win outright, so a corpus of real decks can
 * live outside the repo rather than under a gitignore rule inside the working tree.
 * @param {string | undefined} dirFlag - the flag's value, or `undefined` for the default
 * @returns {string} an absolute directory
 */
export function resolveCorpusDir(dirFlag) {
	return path.resolve(ROOT, dirFlag ?? FIXTURES_DIR)
}

/**
 * Every deck in a corpus directory, by file name, in a stable order.
 *
 * Throws rather than returning `[]` on an empty corpus, and names the directory when it does.
 * An enumerator that quietly yields nothing turns every invariant built on it into a loop that
 * iterates nothing and passes, which is indistinguishable from success in a reporter.
 * @param {object} [opts]
 * @param {string} [opts.dir] - the corpus directory; defaults to {@link FIXTURES_DIR}
 * @param {string | null} [opts.only] - restrict to one file name, as a `--fixture` flag does
 * @param {readonly string[]} [opts.extensions] - which extensions count as a deck
 * @returns {Promise<string[]>}
 */
export async function corpusDecks({ dir = FIXTURES_DIR, only = null, extensions = ['.pptx'] } = {}) {
	const names = (await fs.promises.readdir(dir))
		.filter((name) => extensions.some((ext) => name.endsWith(ext)))
		.filter((name) => !only || name === only)
		.sort()
	if (names.length === 0)
		throw new Error(`no ${extensions.join('/')} file(s) in ${dir}${only ? ` matching ${only}` : ''}`)
	return names
}

/**
 * Import a built entry point, with a build-first message instead of an unresolved-specifier
 * stack when `dist/` is not there.
 *
 * The two `read-emit-*` scripts each spelled this preflight out, the second a verbatim copy of
 * the first down to the `try`/`catch` around `fs.access`.
 *
 * Prints and exits rather than throwing, unlike {@link corpusDecks}: a missing `dist/` is only
 * ever a CLI's problem, while the corpus enumerator is shared with `test/read/corpus.js`, where
 * exiting the process would take the whole test run with it.
 * @param {string} entry - a file name under `dist/`, e.g. `read.js`
 * @param {string} [scriptName] - the package script that builds first, named in the message
 * @returns {Promise<any>} the module
 */
export async function requireDist(entry, scriptName) {
	const file = path.join(ROOT, 'dist', entry)
	if (!fs.existsSync(file)) {
		const alternative = scriptName ? ` (or use \`pnpm run ${scriptName}\`)` : ''
		console.error(`Missing ${path.relative(ROOT, file)}. Run \`pnpm run build\` first${alternative}.`)
		process.exit(1)
	}
	return import(pathToFileURL(file).href)
}

/**
 * Was this module run directly, rather than imported?
 *
 * The gate scripts keep their logic in exported functions and their CLI behind this,
 * so a test can import the parsing and scanning without the script measuring `dist/`,
 * writing a budget file, or calling `process.exit` out from under the test runner.
 * `gen-inspect-snapshot.mjs` open-coded this first; it is shared now that it has
 * more than two callers -- including, at last, `gen-inspect-snapshot.mjs` itself, which went on
 * open-coding it for long enough that this sentence was the only place the sharing had happened.
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
 * A machine-dependent oracle is missing: SKIP on a workstation, fail under `<envVar>=required`.
 *
 * The three desktop oracles are all in the same bind. Not every workstation has LibreOffice or
 * PowerPoint, so a missing tool has to be a SKIP or nobody can run `verify` — but in CI a SKIP
 * is the worst possible outcome: a lane that installs the tools, fails to find them, and
 * reports green while proving nothing. `required` makes their absence the failure instead,
 * which is the bargain `FONT_ORACLES: required` already strikes for the measurement oracles.
 *
 * Returns the exit code to adopt rather than exiting, so a caller that has already collected
 * failures can print them first.
 * @param {string} envVar - the environment variable that opts into the strict mode
 * @param {string} message - what is missing, as a sentence
 * @returns {number} 0 to carry on skipping, 1 to fail
 */
export function skipOrFail(envVar, message) {
	if (process.env[envVar] === 'required') {
		console.error(`${envVar}=required, but ${message}`)
		return 1
	}
	console.log('SKIP: ' + message)
	return 0
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
/** @type {Record<string, string>} */
const localBinPackages = {
	attw: '@arethetypeswrong/cli',
	publint: 'publint',
}

/**
 * Absolute path to an installed package's own `package.json`, or `null` when it is not
 * installed.
 *
 * `require.resolve(pkg + '/package.json')` is the direct form and works for most
 * packages, but it goes through the `exports` map, so a package that does not list
 * `./package.json` there — `publint` is one — throws `ERR_PACKAGE_PATH_NOT_EXPORTED`.
 * The manifest is still sitting on disk; only the subpath lookup is closed. So fall back
 * to resolving the package's main entry, which `exports` does expose, and walking up to
 * the manifest that names it.
 * @param {string} pkg
 * @returns {string | null}
 */
function resolvePackageManifest(pkg) {
	try {
		return requireFromRoot.resolve(pkg + '/package.json')
	} catch {
		// Fall through to the entry-point walk below.
	}
	let dir
	try {
		dir = path.dirname(requireFromRoot.resolve(pkg))
	} catch {
		return null
	}
	for (let next = dir; ; next = path.dirname(next)) {
		const candidate = path.join(next, 'package.json')
		try {
			if (JSON.parse(fs.readFileSync(candidate, 'utf8')).name === pkg) return candidate
		} catch {
			// No manifest here, or an unreadable one: keep climbing.
		}
		if (next === path.dirname(next)) return null
	}
}

/**
 * Absolute path to a local devDependency's JS entry, or `null` when the bin is not one
 * this module owns or the package is not installed.
 * @param {string} name
 * @returns {string | null}
 */
export function resolveLocalBin(name) {
	const pkg = localBinPackages[name]
	if (!pkg) return null
	const manifestPath = resolvePackageManifest(pkg)
	if (!manifestPath) return null
	const { bin } = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
	const entry = typeof bin === 'string' ? bin : bin?.[name]
	if (!entry) return null
	return path.resolve(path.dirname(manifestPath), entry)
}

/**
 * Quote one argument for the Windows shell line built below.
 * @param {string} arg
 * @returns {string}
 */
function quoteArg(arg) {
	return /[\s"^&|<>()]/.test(arg) ? '"' + arg.replace(/"/g, '\\"') + '"' : arg
}

/**
 * Spawn a command, resolving when it exits 0 and rejecting with its output when it does not.
 *
 * `capture` decides whether the child's output is piped back to the caller or inherited by
 * this process: a gate that reports on what a tool said wants the former, one that just
 * needs the tool's exit status wants the latter, so the resolved `stdout`/`stderr` are
 * empty strings unless `capture` is set.
 * @param {string} command
 * @param {readonly string[]} args
 * @param {{env?: NodeJS.ProcessEnv, cwd?: string, capture?: boolean}} [options]
 * @returns {Promise<{stdout: string, stderr: string}>}
 */
export function run(command, args, options = {}) {
	return new Promise((resolve, reject) => {
		/** @type {NodeJS.ProcessEnv} */
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
