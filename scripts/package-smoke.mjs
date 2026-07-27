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

/**
 * Every export subpath the package publishes, with a sample of named exports per entry and
 * the `typeof` each must have once resolved through the *installed tarball's* `exports`
 * map — not through `dist/`, which is why this file packs and installs first.
 *
 * ADDING AN EXPORT SUBPATH MEANS ADDING A ROW HERE. `files`, the `exports` map, and the
 * chunk tsdown actually emits are three independent places a new subpath can go wrong, and
 * a subpath that is missing from all three still typechecks, still builds, and still passes
 * every suite under `test/` — those import from `dist/` by path and never consult `exports`.
 * This matrix is the only check that resolves a subpath the way a consumer does.
 *
 * `hasDefault` records whether the entry carries a default export. Named exports are a
 * sample, not the full surface: the point is to prove the entry resolves and is populated,
 * so pick a few stable, load-bearing names rather than mirroring every export (which would
 * turn every ordinary API addition into a failing package test).
 */
const EXPORT_MATRIX = [
	{
		subpath: '',
		hasDefault: true,
		exports: {
			TsPptx: 'function',
			ShapeType: 'object',
			inchesToEmu: 'function',
			EMU_PER_INCH: 'number',
			textRun: 'function',
		},
	},
	{
		subpath: '/inspect',
		hasDefault: false,
		exports: {
			inspectPptx: 'function',
			boxAnchor: 'function',
			overlapArea: 'function',
			DEFAULT_INSPECT_SLIDE_SIZE: 'object',
		},
	},
	{
		subpath: '/measure',
		hasDefault: false,
		exports: {
			measureText: 'function',
			measureLayout: 'function',
			FontMetricsRegistry: 'function',
			SINGLE_LINE_PITCH: 'number',
		},
	},
	{
		subpath: '/read',
		hasDefault: false,
		exports: {
			Presentation: 'function',
			OpcPackage: 'function',
			isGroupShape: 'function',
			readCoreProperties: 'function',
		},
	},
	{
		subpath: '/script',
		hasDefault: false,
		exports: {
			printScript: 'function',
			diffDeckIr: 'function',
			canonicalDeckIr: 'function',
			readModelToIr: 'function',
		},
	},
	{ subpath: '/math', hasDefault: false, exports: { latexToOmml: 'function', mathmlToOmml: 'function' } },
	{ subpath: '/zip', hasDefault: false, exports: { ZipWriter: 'function', readZip: 'function' } },
	{ subpath: '/node', hasDefault: true, exports: { TsPptx: 'function', ShapeType: 'object' } },
	{ subpath: '/browser', hasDefault: true, exports: { TsPptx: 'function', ShapeType: 'object' } },
]

/**
 * `.` carries `browser`/`node`/`default` conditions, so which artifact a consumer gets
 * depends on the conditions Node resolves under — something no test that imports `dist/`
 * by path can observe. Each run asserts `.` collapses onto the entry for its condition and
 * stays distinct from the other one; a regression that pointed both conditions at the same
 * chunk (or dropped a condition so `default` won everywhere) shows up here and nowhere else.
 */
const CONDITION_RUNS = [
	{ label: 'default', nodeArgs: [], sameAs: '/node', distinctFrom: '/browser' },
	{ label: 'browser', nodeArgs: ['--conditions=browser'], sameAs: '/browser', distinctFrom: '/node' },
]

async function writeFixtureManifest(fixtureDir, manager) {
	await fs.mkdir(fixtureDir, { recursive: true })
	await fs.writeFile(
		path.join(fixtureDir, 'package.json'),
		JSON.stringify({ name: 'ts-pptx-package-smoke-' + manager, private: true, type: 'module' }, null, 2) + '\n'
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
		assertFile(path.join(installedPkgDir, 'dist', 'inspect.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'inspect.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'node.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'node.d.ts')),
		assertFile(path.join(installedPkgDir, 'dist', 'browser.js')),
		assertFile(path.join(installedPkgDir, 'dist', 'browser.d.ts')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'standalone.js')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'core.js')),
		assertNoFile(path.join(installedPkgDir, 'types', 'pptxgen.d.ts')),
		assertNoFile(path.join(installedPkgDir, 'dist', 'pptxgen.js')),
	])

	const matrixRows = EXPORT_MATRIX.map((row) => ({
		specifier: packageImport(row.subpath),
		label: row.subpath || '.',
		hasDefault: row.hasDefault,
		exports: row.exports,
	}))
	await fs.writeFile(
		path.join(fixtureDir, 'matrix-smoke.mjs'),
		`const MATRIX = ${JSON.stringify(matrixRows, null, 1)}

for (const row of MATRIX) {
	const ns = await import(row.specifier)
	if (row.hasDefault && ns.default === undefined) throw new Error(row.label + ': missing default export')
	if (!row.hasDefault && ns.default !== undefined) throw new Error(row.label + ': unexpected default export')
	for (const [name, kind] of Object.entries(row.exports)) {
		if (!(name in ns)) throw new Error(row.label + ': missing named export ' + name)
		if (typeof ns[name] !== kind)
			throw new Error(row.label + ': ' + name + ' is ' + typeof ns[name] + ', expected ' + kind)
	}
}
console.log('  export matrix: ' + MATRIX.length + ' subpaths resolved')
`
	)

	await fs.writeFile(
		path.join(fixtureDir, 'conditions-smoke.mjs'),
		`const NAME = ${JSON.stringify(packageName)}
const [label, sameAs, distinctFrom] = process.argv.slice(2)

const root = await import(NAME)
const same = await import(NAME + sameAs)
const other = await import(NAME + distinctFrom)

// Identity, not a feature probe: all three artifacts expose the same API surface, so only
// module identity can tell which chunk "." actually resolved to.
if (root.TsPptx !== same.TsPptx) throw new Error(label + ': "." did not resolve to ' + sameAs)
if (root.TsPptx === other.TsPptx) throw new Error(label + ': "." is indistinguishable from ' + distinctFrom)
console.log('  conditions [' + label + ']: "." -> ' + sameAs)
`
	)

	await fs.writeFile(
		path.join(fixtureDir, 'esm-smoke.mjs'),
		[
			`import TsPptx, { ChartType, EMU_PER_INCH, STANDARD_LAYOUTS, ShapeType, inchesToEmu, pixelsToEmu } from ${JSON.stringify(packageImport())}`,
			`import NodeTsPptx from ${JSON.stringify(packageImport('/node'))}`,
			`import BrowserTsPptx from ${JSON.stringify(packageImport('/browser'))}`,
			`import { boxAnchor, inspectPptx, overlapArea } from ${JSON.stringify(packageImport('/inspect'))}`,
			'const pptx = new TsPptx()',
			"if (typeof pptx.version !== 'string') throw new Error('missing version')",
			"if (NodeTsPptx !== TsPptx) throw new Error('node entry mismatch')",
			"if (typeof new BrowserTsPptx().version !== 'string') throw new Error('browser entry missing version')",
			"if (ShapeType.rect !== 'rect') throw new Error('missing ShapeType export')",
			"if (ChartType.bar !== 'bar') throw new Error('missing ChartType export')",
			"if (EMU_PER_INCH !== 914400) throw new Error('missing EMU_PER_INCH export')",
			"if (inchesToEmu(STANDARD_LAYOUTS.LAYOUT_WIDE.widthIn) !== 12192000) throw new Error('missing wide layout helpers')",
			"if (pixelsToEmu(1920, 144) !== 12192000) throw new Error('missing pixel conversion helper')",
			"pptx.layout = 'LAYOUT_WIDE'",
			"pptx.addSlide().addText('Inspect me', { x: 1, y: 1, w: 2, h: 0.5, objectName: 'smoke:text' })",
			'const inspected = await inspectPptx(await pptx.stream())',
			"if (inspected.slides[0]?.elements[0]?.name !== 'smoke:text') throw new Error('inspect subpath failed')",
			"if (boxAnchor({ x: 1, y: 2, w: 3, h: 4 }, 'right', 'x') !== 4) throw new Error('boxAnchor failed')",
			"if (overlapArea({ x: 0, y: 0, w: 2, h: 2 }, { x: 1, y: 1, w: 2, h: 2 }) !== 1) throw new Error('overlapArea failed')",
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
			`import TsPptx, { EMU_PER_INCH, ShapeType, STANDARD_LAYOUTS, inchesToEmu, pixelsToEmu, type ChartMulti, type Slide, type StandardLayoutName, type ThemeProps, type WriteFileProps } from ${JSON.stringify(packageImport())}`,
			`import NodeTsPptx from ${JSON.stringify(packageImport('/node'))}`,
			`import BrowserTsPptx from ${JSON.stringify(packageImport('/browser'))}`,
			`import { inspectPptx, type PptxSlideElement, type PptxSlideSize } from ${JSON.stringify(packageImport('/inspect'))}`,
			'const pptx = new TsPptx()',
			'const nodePptx = new NodeTsPptx()',
			'const browserPptx = new BrowserTsPptx()',
			'const slide = pptx.addSlide()',
			"const theme: ThemeProps = { headFontFace: 'Aptos', bodyFontFace: 'Aptos' }",
			"const options: WriteFileProps = { fileName: 'smoke.pptx' }",
			"const comboChart: ChartMulti[] = [{ type: 'bar', data: [{ labels: ['A'], values: [1] }], options: {} }]",
			'const typedSlide: Slide = slide',
			'const inspectResult = await inspectPptx(await pptx.stream())',
			'const inspectedSlideSize: PptxSlideSize = inspectResult.slideSize',
			'const inspectedElement: PptxSlideElement | undefined = inspectResult.slides[0]?.elements[0]',
			"const layoutName: StandardLayoutName = 'LAYOUT_WIDE'",
			'const wideWidthEmu: number = inchesToEmu(STANDARD_LAYOUTS[layoutName].widthIn)',
			'const pxWidthEmu: number = pixelsToEmu(1920, 144)',
			'pptx.theme = theme',
			"slide.addChart([{ labels: ['A'], values: [1] }], { type: 'bar', x: 0, y: 0, w: 1, h: 1 })",
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
			'void typedSlide',
			'void inspectedSlideSize',
			'void inspectedElement',
			'void EMU_PER_INCH',
			'void wideWidthEmu',
			'void pxWidthEmu',
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

	await run(process.execPath, [path.join(fixtureDir, 'matrix-smoke.mjs')], { cwd: fixtureDir })
	for (const conditionRun of CONDITION_RUNS) {
		await run(
			process.execPath,
			[
				...conditionRun.nodeArgs,
				path.join(fixtureDir, 'conditions-smoke.mjs'),
				conditionRun.label,
				conditionRun.sameAs,
				conditionRun.distinctFrom,
			],
			{ cwd: fixtureDir }
		)
	}
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

const tmpRoot = process.env.TSPPTX_PACKAGE_SMOKE_TMPDIR || os.tmpdir()
await fs.mkdir(tmpRoot, { recursive: true })
const tmp = await fs.mkdtemp(path.join(tmpRoot, '.ts-pptx-package-smoke-'))
const keepTmp = process.env.TSPPTX_KEEP_PACKAGE_SMOKE === '1'

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
