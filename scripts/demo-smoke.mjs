#!/usr/bin/env node
import { spawn } from 'node:child_process'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { ROOT } from './rollup-options.mjs'

const ALL_TARGETS = ['node', 'vite']
const requestedTargets = process.argv.slice(2)
const targets = requestedTargets.length > 0 ? requestedTargets : ALL_TARGETS
const keepTmp = process.env.PPTXGENJS_KEEP_DEMO_SMOKE === '1'

for (const target of targets) {
	if (!ALL_TARGETS.includes(target)) {
		throw new Error('unknown demo smoke target "' + target + '"; expected node or vite')
	}
}

function run(command, args, options = {}) {
	return new Promise((resolve, reject) => {
		const packageManagerCache = path.join(os.tmpdir(), 'pptxgenjs-package-manager-cache')
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

function parsePackOutput(stdout) {
	const start = stdout.lastIndexOf('\n[')
	const jsonText = start >= 0 ? stdout.slice(start + 1) : stdout.slice(stdout.indexOf('['))
	const pack = JSON.parse(jsonText)
	if (!Array.isArray(pack) || !pack[0]?.filename) throw new Error('npm pack did not return a tarball filename')
	return pack[0]
}

async function findPackedTarball(packDir) {
	const filename = (await fs.readdir(packDir)).find((file) => file.endsWith('.tgz'))
	if (!filename) throw new Error('npm pack did not create a tarball under ' + packDir)
	return { filename }
}

async function readJson(file) {
	return JSON.parse(await fs.readFile(file, 'utf8'))
}

async function writeJson(file, json) {
	await fs.writeFile(file, JSON.stringify(json, null, '\t') + '\n')
}

async function copyDemoDirs(tmpRoot, dirs) {
	const tmpDemos = path.join(tmpRoot, 'demos')
	await fs.mkdir(tmpDemos, { recursive: true })
	for (const dir of dirs) {
		await fs.cp(path.join(ROOT, 'demos', dir), path.join(tmpDemos, dir), {
			recursive: true,
			filter: (src) => !src.includes(path.sep + 'node_modules' + path.sep) && !src.endsWith(path.sep + 'node_modules'),
		})
	}
	return tmpDemos
}

async function prepareFixturePackage(fixtureDir, tarball) {
	const pkgFile = path.join(fixtureDir, 'package.json')
	const pkg = await readJson(pkgFile)
	pkg.dependencies = { ...pkg.dependencies, pptxgenjs: 'file:' + tarball }
	await writeJson(pkgFile, pkg)
	await fs.rm(path.join(fixtureDir, 'package-lock.json'), { force: true })
	await fs.rm(path.join(fixtureDir, 'pnpm-lock.yaml'), { force: true })
	await fs.rm(path.join(fixtureDir, 'node_modules'), { recursive: true, force: true })
	await run('npm', ['install', '--ignore-scripts', '--no-audit', '--no-fund'], { cwd: fixtureDir })
}

async function smokeNodeDemo(tmpRoot, tarball) {
	const tmpDemos = await copyDemoDirs(tmpRoot, ['node', 'common', 'modules'])
	const fixtureDir = path.join(tmpDemos, 'node')
	await prepareFixturePackage(fixtureDir, tarball)
	await run('npm', ['run', 'demo-text'], { cwd: fixtureDir })
}

async function smokeViteDemo(tmpRoot, tarball) {
	const tmpDemos = await copyDemoDirs(tmpRoot, ['vite-demo', 'common'])
	const fixtureDir = path.join(tmpDemos, 'vite-demo')
	await prepareFixturePackage(fixtureDir, tarball)
	await run('npm', ['run', 'build'], { cwd: fixtureDir })
}

const tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-demo-smoke-'))

try {
	const packDir = path.join(tmp, 'pack')
	await fs.mkdir(packDir, { recursive: true })
	const packResult = await run('npm', ['pack', '--json', '--pack-destination', packDir], { capture: true })
	const packOutput = packResult.stdout || packResult.stderr
	const packInfo = packOutput.trim() ? parsePackOutput(packOutput) : await findPackedTarball(packDir)
	const tarball = path.join(packDir, packInfo.filename)

	for (const target of targets) {
		if (target === 'node') await smokeNodeDemo(path.join(tmp, 'node-fixture'), tarball)
		else if (target === 'vite') await smokeViteDemo(path.join(tmp, 'vite-fixture'), tarball)
	}

	console.log('Demo smoke tests passed: ' + targets.join(', '))
} finally {
	if (keepTmp) console.log('Keeping demo smoke temp directory: ' + tmp)
	else await fs.rm(tmp, { recursive: true, force: true })
}
