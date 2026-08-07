#!/usr/bin/env node
import fs from 'node:fs/promises'
import { isBuiltin } from 'node:module'
import os from 'node:os'
import path from 'node:path'
import esbuild from 'esbuild'
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
	{ subpath: '/html', hasDefault: false, exports: { tableToSlides: 'function' } },
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

/**
 * Subpaths the Node-target bundler step puts through esbuild. Everything in
 * `EXPORT_MATRIX` except `/browser`, which is the *browser* condition's entry: the browser
 * lane already puts a real bundler (Vite/Rolldown) in front of it and then runs what it
 * emitted, and asking a Node-platform bundler about it answers a question no consumer has.
 *
 * Deriving this from the matrix rather than listing it again is deliberate — a new subpath
 * gets bundled because it was added there, not because someone remembered a second list.
 */
const BUNDLE_MATRIX = EXPORT_MATRIX.filter((row) => row.subpath !== '/browser')

/**
 * Runtime dependencies the bundle must be shown to have *resolved and pulled in*, not
 * quietly left external. `opentype.js` is the one that earns this check: it is reached
 * through a **dynamic** `import()` in the measure/fit chunk, which is precisely the shape
 * that is invisible to `node`'s own resolver (it finds it on disk at call time) and to a
 * grep for `from "…"`. The browser lane found the browser entry's copy of this the hard
 * way — the harness failed with `Failed to resolve module specifier 'opentype.js'` — and
 * nothing had been asking the same question of the Node entry.
 */
const BUNDLED_DEPS = ['@xmldom/xmldom', 'fflate', 'opentype.js']

/**
 * Bundle the *installed* package for Node with esbuild, then run what it emitted.
 *
 * This is the Node-side counterpart to what the browser lane does for the `browser`
 * condition, and it exists because those are different questions with different resolvers:
 * `node` finds a specifier on disk when the call happens, while a bundler must resolve
 * every one of them — including dynamic ones — at build time, walking the `exports` map
 * under the conditions its platform implies. A package can be perfectly importable and
 * still be unbundlable.
 *
 * Three assertions, red for different reasons:
 *
 *   - **it builds, with no warnings.** esbuild warns rather than fails on the interesting
 *     cases (an unresolvable dynamic import, a mis-set condition falling through to a
 *     stub), so a warning is treated as a failure here. There are none to allow today; if
 *     one ever has to be, allow it by name, never by muting the channel.
 *   - **nothing but a Node builtin stayed external.** The failure this catches is a bare
 *     specifier the bundler could not resolve and silently deferred to runtime, which is
 *     the same defect as a build error but arrives as a crash in the consumer's process.
 *     Tested with `isBuiltin`, not a `node:` prefix check: the prefix is a convention, not
 *     the rule. `fflate` imports `createRequire` from bare `module`, so a prefix test reads
 *     a dependency's stylistic choice as an unresolvable specifier — which is exactly what
 *     it did on the first run of this check.
 *   - **the emitted bundle runs and writes a real package.** Resolution proves the graph;
 *     only running it proves the graph was assembled into something that works.
 *
 * Runs against both the npm and pnpm fixtures, which is not redundant: pnpm's symlinked
 * store is a genuinely different shape for a bundler to walk than npm's flat tree.
 */
async function bundleForNode(fixtureDir) {
	const imports = []
	const checks = []
	for (const [index, row] of BUNDLE_MATRIX.entries()) {
		const label = row.subpath || '.'
		// Named imports, not `import * as ns`: a namespace object is a tree-shaking barrier,
		// and the point is to bundle the package the way a consumer's bundler would see it.
		const clauses = Object.keys(row.exports).map((name) => `${name} as e${index}_${name}`)
		if (row.hasDefault) clauses.unshift(`default as d${index}`)
		imports.push(`import { ${clauses.join(', ')} } from ${JSON.stringify(packageImport(row.subpath))}`)
		if (row.hasDefault) checks.push(`if (d${index} === undefined) throw new Error('${label}: missing default export')`)
		for (const [name, kind] of Object.entries(row.exports)) {
			checks.push(
				`if (typeof e${index}_${name} !== '${kind}') throw new Error('${label}: ${name} is not a ${kind} once bundled')`
			)
		}
	}

	const entryFile = path.join(fixtureDir, 'bundle-entry.mjs')
	await fs.writeFile(
		entryFile,
		[
			...imports,
			...checks,
			// Every check above is satisfiable by a binding that resolved to nothing useful.
			// Building a deck is what proves the bundled graph is wired: the zip writer, the
			// XML serializers and the Node runtime adapter all have to be present and reachable.
			`const pptx = new e0_TsPptx()`,
			`pptx.addSlide().addText('bundled', { x: 1, y: 1, w: 2, h: 0.5 })`,
			`const bytes = new Uint8Array(await pptx.stream())`,
			`if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('bundled build did not emit a zip')`,
			`if (bytes.length < 1000) throw new Error('bundled build emitted ' + bytes.length + ' bytes, expected a deck')`,
			'',
		].join('\n')
	)

	const outFile = path.join(fixtureDir, 'bundle-out.mjs')
	const result = await esbuild.build({
		absWorkingDir: fixtureDir,
		entryPoints: [entryFile],
		outfile: outFile,
		bundle: true,
		// `node` platform is what makes esbuild resolve the `exports` map under the `node`
		// condition and treat `node:` builtins as external — the resolution a serverless or
		// `ncc`-style consumer gets. It is also what keeps `dist/zip.js`'s lazy
		// `import('node:fs/promises')` from being a finding: it is a builtin, not a gap.
		platform: 'node',
		format: 'esm',
		metafile: true,
		// Warnings are inspected below rather than printed, so a red run says which one.
		logLevel: 'silent',
	})

	if (result.warnings.length) {
		const formatted = await esbuild.formatMessages(result.warnings, { kind: 'warning', color: false })
		throw new Error('bundling the installed package for Node produced warnings:\n' + formatted.join('\n'))
	}

	const external = []
	for (const output of Object.values(result.metafile.outputs)) {
		for (const imported of output.imports ?? []) {
			if (imported.external && !isBuiltin(imported.path)) external.push(imported.path)
		}
	}
	if (external.length) {
		throw new Error(
			'the Node bundle left non-builtin specifiers external, so a bundler could not resolve them: ' +
				[...new Set(external)].sort().join(', ')
		)
	}

	// esbuild normalizes metafile paths to forward slashes on every platform, so one shape
	// of check works for npm's flat tree and pnpm's `.pnpm/<name>@<version>/node_modules/…`.
	const inputs = Object.keys(result.metafile.inputs)
	const missing = BUNDLED_DEPS.filter((dep) => !inputs.some((input) => input.includes(`node_modules/${dep}/`)))
	if (missing.length) {
		throw new Error(
			'the Node bundle never pulled in ' +
				missing.join(', ') +
				' — a dependency that is neither bundled nor external has been dropped from the graph'
		)
	}

	await run(process.execPath, [outFile], { cwd: fixtureDir })
	console.log(`  node bundle: ${BUNDLE_MATRIX.length} subpaths bundled and run (esbuild ${esbuild.version})`)
}

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
	await bundleForNode(fixtureDir)
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
