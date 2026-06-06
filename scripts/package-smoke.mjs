#!/usr/bin/env node
import { spawn } from 'node:child_process'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import vm from 'node:vm'
import { ROOT } from './rollup-options.mjs'

function run(command, args, options = {}) {
	return new Promise((resolve, reject) => {
		const npmCache = path.join(os.tmpdir(), 'pptxgenjs-npm-cache')
		const child = spawn(command, args, {
			cwd: options.cwd || ROOT,
			env: {
				...process.env,
				npm_config_cache: npmCache,
				NPM_CONFIG_CACHE: npmCache,
				...options.env,
			},
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

async function assertFile(file) {
	await fs.access(file)
}

async function assertNoFile(file) {
	try {
		await fs.access(file)
	} catch {
		return
	}
	throw new Error('unexpected file exists: ' + file)
}

async function smokeBrowserBundle(bundlePath) {
	const code = await fs.readFile(bundlePath, 'utf8')
	const sandbox = {
		ArrayBuffer,
		Blob,
		Buffer,
		clearInterval,
		clearTimeout,
		console,
		DataView,
		Promise,
		setInterval,
		setTimeout,
		TextDecoder,
		TextEncoder,
		Uint8Array,
	}
	sandbox.globalThis = sandbox
	sandbox.self = sandbox
	sandbox.window = sandbox
	sandbox.navigator = { userAgent: 'node' }

	vm.runInNewContext(code, sandbox, { filename: bundlePath })
	const PptxGenJS = sandbox.PptxGenJS || sandbox.window.PptxGenJS || sandbox.globalThis.PptxGenJS
	if (typeof PptxGenJS !== 'function') {
		throw new Error('browser bundle did not expose PptxGenJS')
	}
}

const tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-package-smoke-'))
const keepTmp = process.env.PPTXGENJS_KEEP_PACKAGE_SMOKE === '1'

try {
	const packDir = path.join(tmp, 'pack')
	const fixtureDir = path.join(tmp, 'fixture')
	await fs.mkdir(packDir, { recursive: true })
	await fs.mkdir(fixtureDir, { recursive: true })

	const packResult = await run('npm', ['pack', '--json', '--pack-destination', packDir], { capture: true })
	const packOutput = packResult.stdout || packResult.stderr
	const packInfo = packOutput.trim() ? parsePackOutput(packOutput) : await findPackedTarball(packDir)
	const tarball = path.join(packDir, packInfo.filename)
	await assertFile(tarball)

	await fs.writeFile(
		path.join(fixtureDir, 'package.json'),
		JSON.stringify({ name: 'pptxgenjs-package-smoke', private: true, type: 'module' }, null, 2) + '\n'
	)
	await run('npm', ['install', '--ignore-scripts', '--no-audit', '--no-fund', tarball], { cwd: fixtureDir })

	const installedPkgDir = path.join(fixtureDir, 'node_modules', 'pptxgenjs')
	await Promise.all([
		assertFile(path.join(installedPkgDir, 'dist', 'pptxgen.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'pptxgen.bundle.js')),
		assertFile(path.join(installedPkgDir, 'types', 'pptxgen.d.ts')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.cjs.js')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.es.js')),
	])

	await fs.writeFile(
		path.join(fixtureDir, 'esm-smoke.mjs'),
		[
			"import PptxGenJS from 'pptxgenjs'",
			'const pptx = new PptxGenJS()',
			"if (typeof pptx.version !== 'string') throw new Error('missing version')",
			'',
		].join('\n')
	)
	await fs.writeFile(
		path.join(fixtureDir, 'cjs-unsupported.cjs'),
		[
			"try { require('pptxgenjs') } catch { process.exit(0) }",
			"throw new Error('CommonJS require unexpectedly resolved pptxgenjs')",
			'',
		].join('\n')
	)
	await fs.writeFile(
		path.join(fixtureDir, 'type-smoke.ts'),
		[
			"import PptxGenJS, { type ThemeProps, type WriteFileProps } from 'pptxgenjs'",
			'const pptx = new PptxGenJS()',
			"const theme: ThemeProps = { headFontFace: 'Aptos', bodyFontFace: 'Aptos' }",
			"const options: WriteFileProps = { fileName: 'smoke.pptx' }",
			'pptx.theme = theme',
			'void pptx.writeFile(options)',
			'',
		].join('\n')
	)
	await fs.writeFile(
		path.join(fixtureDir, 'tsconfig.json'),
		JSON.stringify(
			{
				compilerOptions: {
					lib: ['dom', 'es2024'],
					module: 'esnext',
					moduleResolution: 'bundler',
					noEmit: true,
					strict: true,
					target: 'es2024',
				},
				include: ['type-smoke.ts'],
			},
			null,
			2
		) + '\n'
	)

	await run(process.execPath, [path.join(fixtureDir, 'esm-smoke.mjs')], { cwd: fixtureDir })
	await run(process.execPath, [path.join(fixtureDir, 'cjs-unsupported.cjs')], { cwd: fixtureDir })
	await run(process.execPath, [
		path.join(ROOT, 'node_modules', 'typescript', 'bin', 'tsc'),
		'-p',
		path.join(fixtureDir, 'tsconfig.json'),
	])
	await smokeBrowserBundle(path.join(installedPkgDir, 'dist', 'pptxgen.bundle.js'))

	console.log('Packed package smoke test passed: ' + packInfo.filename)
} finally {
	if (keepTmp) console.log('Keeping package smoke temp directory: ' + tmp)
	else await fs.rm(tmp, { recursive: true, force: true })
}
