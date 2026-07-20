import { spawn } from 'node:child_process'
import fsSync from 'node:fs'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')

const packageManagerCache = process.env.PPTXGENJS_SCRIPT_CACHE_DIR || path.join(ROOT, '.tmp', 'package-manager-cache')

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
	const { bin } = JSON.parse(fsSync.readFileSync(manifestPath, 'utf8'))
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

export function parsePackOutput(output) {
	const text = output.trim()
	const objectStart = text.lastIndexOf('\n{')
	const arrayStart = text.lastIndexOf('\n[')
	const start = Math.max(objectStart, arrayStart)
	if (start >= 0) return JSON.parse(text.slice(start + 1))

	const firstObject = text.indexOf('{')
	const firstArray = text.indexOf('[')
	const firstJson = [firstObject, firstArray].filter((idx) => idx >= 0).sort((a, b) => a - b)[0]
	if (firstJson === undefined) throw new Error('pack command did not print JSON output')
	return JSON.parse(text.slice(firstJson))
}

export async function packPackage(packDir) {
	await fs.mkdir(packDir, { recursive: true })
	const result = await run('pnpm', ['pack', '--json', '--pack-destination', packDir], { capture: true })
	const output = result.stdout || result.stderr
	let entry
	if (output.trim()) {
		const pack = parsePackOutput(output)
		entry = Array.isArray(pack) ? pack[0] : pack
	}
	if (!entry?.filename) {
		const packedFiles = (await fs.readdir(packDir)).filter((file) => file.endsWith('.tgz')).sort()
		if (packedFiles.length !== 1) throw new Error('pnpm pack did not return exactly one tarball in ' + packDir)
		entry = { filename: packedFiles[0] }
	}

	const tarball = path.isAbsolute(entry.filename) ? entry.filename : path.join(packDir, path.basename(entry.filename))
	await assertFile(tarball)
	return { ...entry, filename: path.basename(entry.filename), tarball }
}

export async function assertFile(file) {
	await fs.access(file)
}

export async function assertNoFile(file) {
	try {
		await fs.access(file)
	} catch {
		return
	}
	throw new Error('unexpected file exists: ' + file)
}
