import { spawn } from 'node:child_process'
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT } from './rollup-options.mjs'

const packageManagerCache = process.env.PPTXGENJS_SCRIPT_CACHE_DIR || path.join(ROOT, '.tmp', 'package-manager-cache')

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
		const child = spawn(command, args, {
			cwd: options.cwd || ROOT,
			env,
			stdio: options.capture ? ['ignore', 'pipe', 'pipe'] : 'inherit',
		})
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
	const pack = parsePackOutput(result.stdout || result.stderr)
	const entry = Array.isArray(pack) ? pack[0] : pack
	if (!entry?.filename) throw new Error('pnpm pack did not return a tarball filename')

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
