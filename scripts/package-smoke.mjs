#!/usr/bin/env node
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { ROOT } from './rollup-options.mjs'
import { assertFile, assertNoFile, packPackage, run } from './script-utils.mjs'

async function writeFixtureManifest(fixtureDir, manager) {
	await fs.mkdir(fixtureDir, { recursive: true })
	await fs.writeFile(
		path.join(fixtureDir, 'package.json'),
		JSON.stringify({ name: 'pptxgenjs-package-smoke-' + manager, private: true, type: 'module' }, null, 2) + '\n'
	)
}

async function installPackedPackage(manager, fixtureDir, tarball) {
	if (manager === 'npm') {
		await run('npm', ['install', '--ignore-scripts', '--no-audit', '--no-fund', tarball], { cwd: fixtureDir })
		return
	}
	if (manager === 'pnpm') {
		await run('pnpm', ['add', '--ignore-scripts', tarball], { cwd: fixtureDir })
		return
	}
	throw new Error('unsupported package manager for smoke test: ' + manager)
}

async function smokeInstalledPackage(fixtureDir) {
	const installedPkgDir = path.join(fixtureDir, 'node_modules', 'pptxgenjs')
	await Promise.all([
		assertFile(path.join(installedPkgDir, 'dist', 'pptxgen.js')),
		assertFile(path.join(installedPkgDir, 'types', 'pptxgen.d.ts')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.cjs.js')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.es.js')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.min.js')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.min.js.map')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.bundle.js')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.bundle.js.map')),
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
			'const slide = pptx.addSlide()',
			"const theme: ThemeProps = { headFontFace: 'Aptos', bodyFontFace: 'Aptos' }",
			"const options: WriteFileProps = { fileName: 'smoke.pptx' }",
			'pptx.theme = theme',
			"slide.addImage({ data: 'image/png;base64,AAAA', x: 0, y: 0, w: 1, h: 1 })",
			"slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/example', x: 0, y: 0, w: 1, h: 1 })",
			'slide.addText(42, { x: 0, y: 0, w: 1, h: 1 })',
			'// @ts-expect-error public slides do not expose internal OOXML object storage',
			'slide._slideObjects',
			'// @ts-expect-error presentation slides getter returns the public slide shape',
			'pptx.slides[0]._rels',
			'// @ts-expect-error slide layouts getter does not expose internal relationship storage',
			'pptx.slideLayouts[0]._rels',
			'// @ts-expect-error addImage requires data or path',
			'slide.addImage({ x: 0, y: 0, w: 1, h: 1 })',
			'// @ts-expect-error file media requires data or path',
			"slide.addMedia({ type: 'video', x: 0, y: 0, w: 1, h: 1 })",
			'void pptx.writeFile(options)',
			'',
		].join('\n')
	)
	const typeSmokeConfigs = [
		{
			fileName: 'tsconfig.bundler.json',
			compilerOptions: {
				lib: ['dom', 'es2024'],
				module: 'esnext',
				moduleResolution: 'bundler',
				noEmit: true,
				strict: true,
				target: 'es2024',
			},
		},
		{
			fileName: 'tsconfig.nodenext.json',
			compilerOptions: {
				lib: ['dom', 'es2024'],
				module: 'nodenext',
				moduleResolution: 'nodenext',
				noEmit: true,
				strict: true,
				target: 'es2024',
			},
		},
	]
	await Promise.all(
		typeSmokeConfigs.map((config) =>
			fs.writeFile(
				path.join(fixtureDir, config.fileName),
				JSON.stringify(
					{
						compilerOptions: config.compilerOptions,
						include: ['type-smoke.ts'],
					},
					null,
					2
				) + '\n'
			)
		)
	)

	await run(process.execPath, [path.join(fixtureDir, 'esm-smoke.mjs')], { cwd: fixtureDir })
	await run(process.execPath, [path.join(fixtureDir, 'cjs-unsupported.cjs')], { cwd: fixtureDir })
	for (const config of typeSmokeConfigs) {
		await run(process.execPath, [
			path.join(ROOT, 'node_modules', 'typescript', 'bin', 'tsc'),
			'-p',
			path.join(fixtureDir, config.fileName),
		])
	}
}

const tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-package-smoke-'))
const keepTmp = process.env.PPTXGENJS_KEEP_PACKAGE_SMOKE === '1'

try {
	const packDir = path.join(tmp, 'pack')
	const packInfo = await packPackage(packDir)

	for (const manager of ['npm', 'pnpm']) {
		const fixtureDir = path.join(tmp, manager + '-fixture')
		await writeFixtureManifest(fixtureDir, manager)
		await installPackedPackage(manager, fixtureDir, packInfo.tarball)
		await smokeInstalledPackage(fixtureDir)
	}

	console.log('Packed package smoke test passed with npm and pnpm: ' + packInfo.filename)
} finally {
	if (keepTmp) console.log('Keeping package smoke temp directory: ' + tmp)
	else await fs.rm(tmp, { recursive: true, force: true })
}
