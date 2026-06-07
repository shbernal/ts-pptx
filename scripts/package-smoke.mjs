#!/usr/bin/env node
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { ROOT, assertFile, assertNoFile, packPackage, run } from './script-utils.mjs'

const packageJson = JSON.parse(await fs.readFile(path.join(ROOT, 'package.json'), 'utf8'))
const packageName = packageJson.name
const packagePathParts = packageName.split('/')

function packageImport(subpath = '') {
	return packageName + subpath
}

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
	const installedPkgDir = path.join(fixtureDir, 'node_modules', ...packagePathParts)
	await Promise.all([
		assertFile(path.join(installedPkgDir, 'dist', 'index.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'index.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'core.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'core.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'node.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'node.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'browser.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'browser.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'standalone.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'standalone.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'standalone.js.map')),
		assertFile(path.join(installedPkgDir, 'dist', 'standalone.d.ts.map')),
		assertNoFile(path.join(installedPkgDir, 'types', 'pptxgen.d.ts')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.js')),
	])

	await fs.writeFile(
		path.join(fixtureDir, 'esm-smoke.mjs'),
		[
			`import PptxGenJS, { Presentation, ShapeType } from ${JSON.stringify(packageImport())}`,
			`import NodePptxGenJS from ${JSON.stringify(packageImport('/node'))}`,
			`import BrowserPptxGenJS from ${JSON.stringify(packageImport('/browser'))}`,
			`import StandalonePptxGenJS from ${JSON.stringify(packageImport('/standalone'))}`,
			`import { ChartType } from ${JSON.stringify(packageImport('/core'))}`,
			'const pptx = new PptxGenJS()',
			"if (typeof pptx.version !== 'string') throw new Error('missing version')",
			"if (Presentation !== PptxGenJS) throw new Error('missing Presentation named export')",
			"if (NodePptxGenJS !== PptxGenJS) throw new Error('node entry mismatch')",
			"if (typeof new BrowserPptxGenJS().version !== 'string') throw new Error('browser entry missing version')",
			"if (typeof new StandalonePptxGenJS().version !== 'string') throw new Error('standalone entry missing version')",
			"if (ShapeType.rect !== 'rect') throw new Error('missing ShapeType export')",
			"if (ChartType.bar !== 'bar') throw new Error('missing ChartType export')",
			'',
		].join('\n')
	)
	await fs.writeFile(
		path.join(fixtureDir, 'cjs-contract.cjs'),
		[
			`const pkg = require(${JSON.stringify(packageImport('/package.json'))})`,
			"if (JSON.stringify(pkg.exports).includes('\"require\"')) throw new Error('unexpected require export condition')",
			"if (pkg.main || pkg.module) throw new Error('unexpected legacy main/module field')",
			'',
		].join('\n')
	)
	await fs.writeFile(
		path.join(fixtureDir, 'type-smoke.ts'),
		[
			`import PptxGenJS, { Presentation, type IChartMulti, type ThemeProps, type WriteFileProps } from ${JSON.stringify(packageImport())}`,
			`import NodePptxGenJS from ${JSON.stringify(packageImport('/node'))}`,
			`import BrowserPptxGenJS from ${JSON.stringify(packageImport('/browser'))}`,
			`import StandalonePptxGenJS from ${JSON.stringify(packageImport('/standalone'))}`,
			`import { ShapeType, type PresSlide } from ${JSON.stringify(packageImport('/core'))}`,
			'const pptx = new PptxGenJS()',
			'const nodePptx = new NodePptxGenJS()',
			'const browserPptx = new BrowserPptxGenJS()',
			'const standalonePptx = new StandalonePptxGenJS()',
			'const slide = pptx.addSlide()',
			"const theme: ThemeProps = { headFontFace: 'Aptos', bodyFontFace: 'Aptos' }",
			"const options: WriteFileProps = { fileName: 'smoke.pptx' }",
			"const comboChart: IChartMulti[] = [{ type: 'bar', data: [{ labels: ['A'], values: [1] }], options: {} }]",
			'const presentationCtor: typeof PptxGenJS = Presentation',
			'const typedSlide: PresSlide = slide',
			'pptx.theme = theme',
			"slide.addChart('bar', [{ labels: ['A'], values: [1] }], { x: 0, y: 0, w: 1, h: 1 })",
			'slide.addChart(comboChart, { x: 0, y: 0, w: 1, h: 1 })',
			"slide.addImage({ data: 'image/png;base64,AAAA', x: 0, y: 0, w: 1, h: 1 })",
			"slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/example', x: 0, y: 0, w: 1, h: 1 })",
			'slide.addText(42, { x: 0, y: 0, w: 1, h: 1 })',
			'slide.addShape(ShapeType.rect, { x: 0, y: 0, w: 1, h: 1 })',
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
			'void nodePptx.write()',
			'void browserPptx.write()',
			'void standalonePptx.write()',
			'void presentationCtor',
			'void typedSlide',
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
	await run(process.execPath, [path.join(fixtureDir, 'cjs-contract.cjs')], { cwd: fixtureDir })
	for (const config of typeSmokeConfigs) {
		await run(process.execPath, [
			path.join(ROOT, 'node_modules', 'typescript', 'bin', 'tsc'),
			'-p',
			path.join(fixtureDir, config.fileName),
		])
	}
}

const tmpRoot = process.env.PPTXGENJS_PACKAGE_SMOKE_TMPDIR || os.tmpdir()
await fs.mkdir(tmpRoot, { recursive: true })
const tmp = await fs.mkdtemp(path.join(tmpRoot, '.pptxgenjs-package-smoke-'))
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
