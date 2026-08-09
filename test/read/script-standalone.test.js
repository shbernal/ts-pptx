// The standalone printer — `ts-pptx/script`'s Tier A write half, and the round trip over it.
//
// The template-anchored tier gets a deck's whole design back byte for byte because it never
// tries to reproduce it. This one has to, so its failure modes are different in kind and the
// checks below are split along that seam:
//
//  1. **The round trip** (source → IR₁ → script → run → IR₂) covers everything the printer
//     decides: whether a theme is emitted, whether the layouts are, whether a slide binds to
//     the master its source layout became. Mutation testing put 7 of 12 deliberate defects in
//     this bucket, including three that only the *chrome* half of the projection catches.
//  2. **Direct IR expectations** cover what the round trip structurally cannot. Both IRs come
//     from the same mapper, so a mapper that never reads the theme's colour scheme produces an
//     output that also lacks it, and the diff is clean. Those checks read the fixture through
//     `ts-pptx/read` and compare the IR against *that*, never against the converter — the same
//     rule Phase 3's chart-arity bug earned.
//  3. **Two manufactured decks**, because the fixture corpus contains neither a slide-owned
//     background nor an extended chart, so the mutations that delete each survived against every
//     fixture. Both are authored through the write API here rather than waited on as fixtures.

import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'
import { describe, expect, test } from 'vitest'
import TsPptx, { ChartType } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { canonicalDeckIr, diffDeckIr, printScript, printStandaloneScript, readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'

const run = promisify(execFile)
const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = path.join(__dirname, 'fixtures')
const REPO = path.join(__dirname, '..', '..')

// Inside the repo, not the OS temp directory: the emitted script imports this package by its
// published name, which Node resolves by the self-reference rule only from a path underneath
// the package root. `/.tmp/` is gitignored.
const SCRATCH = path.join(REPO, '.tmp')

const fixtureNames = (await readdir(FIXTURES)).filter((name) => name.endsWith('.pptx')).sort()

async function irFor(name) {
	return readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, name))))
}

/**
 * Lay a standalone script out on disk, run it, and read back what it produced.
 *
 * No template is written, which is the whole point of the tier: if the script needed one, this
 * would fail rather than quietly pass with the source deck's chrome in scope.
 */
async function runStandalone(bytes, options = {}) {
	await mkdir(SCRATCH, { recursive: true })
	const dir = await mkdtemp(path.join(SCRATCH, 'standalone-'))
	try {
		const ir = readModelToIr(await Presentation.load(bytes))
		const printed = printStandaloneScript(ir, options)

		await writeFile(path.join(dir, 'script.ts'), printed.code)
		if (printed.assets.size > 0) {
			await mkdir(path.join(dir, 'assets'), { recursive: true })
			for (const [name, data] of printed.assets) await writeFile(path.join(dir, 'assets', name), data)
		}

		// Node's type stripping, so the script under test is the exact text a user is handed.
		await run(process.execPath, ['--no-warnings', path.join(dir, 'script.ts')])
		const outputBytes = await readFile(path.join(dir, 'output.pptx'))
		const output = await Presentation.load(outputBytes)
		const outputIr = readModelToIr(output)
		return {
			ir,
			printed,
			output,
			outputIr,
			report: diffDeckIr(canonicalDeckIr(ir), canonicalDeckIr(outputIr), printed.notes),
		}
	} finally {
		await rm(dir, { recursive: true, force: true })
	}
}

/** Author a deck through the write API and return its bytes, for cases the corpus lacks. */
async function authored(build) {
	const pptx = new TsPptx()
	build(pptx)
	return /** @type {Buffer} */ (await pptx.write({ outputType: 'nodebuffer' }))
}

/** Every shape name on a slide, descending into groups. */
function shapeNames(shapes, out = []) {
	for (const shape of shapes) {
		out.push(shape.name)
		if (shape.shapes) shapeNames(shape.shapes, out)
	}
	return out
}

describe('standalone printer — corpus invariants', () => {
	test('every fixture prints, and prints identically twice', async () => {
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			const first = printStandaloneScript(ir)
			const second = printStandaloneScript(ir)
			assert(first.code.length > 0, `${name}: printed nothing`)
			assertEqual(first.code, second.code, `${name}: printing is not deterministic`)
		}
	})

	test('no printed value is the literal `undefined`', async () => {
		for (const name of fixtureNames) {
			const { code } = printStandaloneScript(await irFor(name))
			const offender = code.split('\n').find((line) => /(?<![A-Za-z0-9_$])undefined(?![A-Za-z0-9_$])/.test(line))
			assertEqual(offender, undefined, `${name}: printed an \`undefined\``)
		}
	})

	test('nothing in the emitted script refers to a template', async () => {
		// The tier's defining claim. A stray `fromTemplate`, `Presentation` import or
		// `template.pptx` path would make the script need the file it is meant to replace.
		for (const name of fixtureNames) {
			const { code } = printStandaloneScript(await irFor(name))
			for (const banned of ['fromTemplate', 'template.pptx', 'appendSlides', 'importSlide']) {
				assert(!code.includes(banned), `${name}: standalone output mentions ${banned}`)
			}
		}
	})

	test('the emitted import specifier is this package’s published name', async () => {
		// The directory is `ts-pptx`, the package is not; that shipped wrong once in the other
		// tier. Pinned to the manifest rather than to a literal anyone can re-guess.
		const manifest = JSON.parse(await readFile(path.join(REPO, 'package.json'), 'utf8'))
		const { code } = printStandaloneScript(await irFor('empty.pptx'))
		assert(code.includes(`import TsPptx from '${manifest.name}'`), `import is not ${manifest.name}`)
		assert(!code.includes(`${manifest.name}/read`), 'a standalone script must not import the read half')
	})

	test('every master title is unique, because it is also the binding key', async () => {
		// `addSlide({ masterTitle })` resolves by title, and `defineSlideMaster` happily accepts a
		// duplicate, so two same-named layouts would silently send both slides to whichever won.
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			const titles = ir.chrome.masters.map((master) => master.props.title)
			assertEqual(new Set(titles).size, titles.length, `${name}: duplicate master titles ${titles.join(', ')}`)
			for (const slide of ir.slides) {
				if (slide.layout === null) continue
				const master = ir.chrome.masters.find((entry) => entry.layoutIndex === slide.layout.index)
				assert(master !== undefined, `${name}: slide ${slide.number} binds to a layout with no master`)
			}
		}
	})
})

describe('standalone printer — the chrome IR, read against the deck rather than the converter', () => {
	// The round trip cannot judge any of this: both IRs come from the same mapper, so a mapper
	// that dropped the whole theme would produce an output that also lacks it and compare clean
	// (measured — three mutations survive there and are covered here instead). Expectations
	// therefore come from `ts-pptx/read`'s own accessors.
	test('the theme IR carries the deck’s colour scheme and font faces', async () => {
		const presentation = await Presentation.load(await readFile(path.join(FIXTURES, 'theme-colors.pptx')))
		const theme = presentation.masters()[0].theme
		assert(theme !== null, 'the fixture has a theme to compare against')

		const { chrome } = readModelToIr(presentation)
		for (const [slot, hex] of Object.entries(theme.colorScheme)) {
			if (hex === null) continue
			assertEqual(chrome.theme.colorScheme?.[slot], hex.replace(/^#/, '').toUpperCase(), `colour slot ${slot}`)
		}
		assertEqual(chrome.theme.headFontFace, theme.fontScheme?.major.latin ?? undefined, 'heading font face')
		assertEqual(chrome.theme.bodyFontFace, theme.fontScheme?.minor.latin ?? undefined, 'body font face')
	})

	test('there is one master per source layout, in gallery order, carrying its background', async () => {
		const presentation = await Presentation.load(await readFile(path.join(FIXTURES, 'mixed.pptx')))
		const gallery = presentation.layouts()
		const { chrome } = readModelToIr(presentation)
		assertEqual(chrome.masters.length, gallery.length, 'one master per source layout')

		const layouts = presentation.masters().flatMap((master) => master.layouts)
		chrome.masters.forEach((master, index) => {
			assertEqual(master.layoutIndex, index, `master ${index} addresses its own gallery position`)
			// The title may be deduplicated or whitespace-collapsed, but it always starts from the
			// source layout's own name.
			const expected = layouts[index].name.replace(/[\t\r\n]+/g, ' ')
			assert(
				String(master.props.title).startsWith(expected),
				`master ${index} title ${JSON.stringify(master.props.title)} does not come from ${JSON.stringify(expected)}`
			)
			const background = layouts[index].background ?? presentation.masters()[0].background
			if (background?.type === 'solid' && background.color) {
				const emitted = /** @type {Record<string, string> | undefined} */ (master.props.background)
				assertEqual(
					emitted?.color,
					background.color.effectiveHex.replace(/^#/, '').toUpperCase(),
					`master ${index} background`
				)
			}
		})
	})

	test('a deck with several masters is flattened, loudly', async () => {
		// `read-stress.pptx` is the corpus's only multi-master deck — `multi-theme.pptx`, despite
		// the name, carries several *themes* under one master.
		const ir = await irFor('read-stress.pptx')
		const presentation = await Presentation.load(await readFile(path.join(FIXTURES, 'read-stress.pptx')))
		assert(presentation.masters().length > 1, 'the fixture has more than one master to flatten')
		assert(
			ir.fidelity.some((note) => note.construct === 'master.multiple'),
			'a multi-master deck must declare the flattening'
		)
		// Every layout still reaches the gallery — it is the masters that collapse, not the layouts.
		assertEqual(ir.chrome.masters.length, presentation.layouts().length, 'no layout is dropped by the flattening')
	})
})

describe('standalone printer — the emitted script runs, with no template in reach', () => {
	test('a plain deck rebuilds its slides, hidden flag, text and layout gallery', async () => {
		const { output, report } = await runStandalone(await readFile(path.join(FIXTURES, 'hidden.pptx')))
		assertEqual(output.slides.length, 2, 'slide count')
		assertEqual(output.slides[1].hidden, true, 'slide 2 stays hidden')
		assert(output.slides[0].shapes.length > 0, 'slide 1 has content')
		assertEqual(report.undeclared.length, 0, 'no undeclared difference')
	})

	test('the theme reaches the output deck, not just the script text', async () => {
		const { ir, output } = await runStandalone(await readFile(path.join(FIXTURES, 'theme-colors.pptx')))
		const theme = output.masters()[0].theme
		assert(theme !== null, 'the output deck has a theme')
		for (const [slot, hex] of Object.entries(ir.chrome.theme.colorScheme ?? {})) {
			assertEqual(theme.colorScheme[slot]?.replace(/^#/, '').toUpperCase(), hex, `colour slot ${slot} survived`)
		}
	})

	test('each slide binds to the master its source layout became', async () => {
		const { ir, outputIr } = await runStandalone(await readFile(path.join(FIXTURES, 'mixed.pptx')))
		ir.slides.forEach((slide, index) => {
			if (slide.layout === null) return
			const master = ir.chrome.masters.find((entry) => entry.layoutIndex === slide.layout.index)
			assertEqual(outputIr.slides[index].layout?.name, master.props.title, `slide ${slide.number} layout binding`)
		})
	})

	test('a deck whose layout names repeat still binds each slide to its own layout', async () => {
		// `read-stress.pptx` carries two layouts called "Title and Text", one per master. A title
		// is the binding key here, so the duplicate has to be suffixed and the slide follow it.
		const { printed, report } = await runStandalone(await readFile(path.join(FIXTURES, 'read-stress.pptx')))
		assert(
			printed.notes.some((note) => note.construct === 'master.nameCollision'),
			'the rename is declared'
		)
		assertEqual(report.undeclared.length, 0, 'no undeclared difference')
	})
})

describe('standalone printer — cases the fixture corpus does not contain', () => {
	// Both of these had a deliberate defect planted in the printer that survived every
	// fixture, because no fixture exercises the construct. Authored here rather than deferred:
	// the write path is the fixture, the same technique `chartex-read.test.js` uses.
	test('a slide-owned background survives, and is reported when it is deleted', async () => {
		const bytes = await authored((pptx) => {
			const slide = pptx.addSlide()
			slide.background = { color: 'C00000' }
			slide.addText('background', { x: 1, y: 1, w: 4, h: 1 })
		})
		const { ir, outputIr, report } = await runStandalone(bytes)
		assertEqual(ir.slides[0].background?.color, 'C00000', 'the source background reaches the IR')
		assertEqual(outputIr.slides[0].background?.color, 'C00000', 'and the output deck')
		assertEqual(report.undeclared.length, 0, 'no undeclared difference')

		// The oracle's half of the claim: drop it from the output IR and the diff must say so.
		// Without this the test above passes just as happily against a printer that emits no
		// background at all, since nothing would then differ between the two sides.
		const perturbed = canonicalDeckIr(outputIr)
		perturbed.slides[0].background = null
		const dirty = diffDeckIr(canonicalDeckIr(ir), perturbed, [])
		assert(
			dirty.undeclared.some((difference) => difference.field === 'background'),
			'a lost slide background must be reported'
		)
	})

	test('an extended chart is transcribed around rather than carried, since there is nothing to carry from', async () => {
		const bytes = await authored((pptx) => {
			const slide = pptx.addSlide()
			slide.addText('kept', { x: 0.5, y: 0.5, w: 4, h: 1, objectName: 'Survivor' })
			slide.addChart([{ name: 'Cash Flow', labels: ['Start', 'Q1', 'End'], values: [100, 40, 190] }], {
				type: ChartType.waterfall,
				x: 1,
				y: 2,
				w: 6,
				h: 4,
			})
		})
		const ir = readModelToIr(await Presentation.load(bytes))
		assertEqual(ir.slides[0].source, 'carried', 'the IR marks the slide for copying')
		assert(ir.slides[0].calls.length > 0, 'and still transcribes it, for a printer that cannot copy')

		const { printed, output, report } = await runStandalone(bytes)
		assertEqual(report.undeclared.length, 0, 'no undeclared difference')
		assert(
			printed.notes.some((note) => note.construct === 'chartEx.all'),
			'the extended chart is declared lost'
		)
		assert(
			!printed.notes.some((note) => note.construct === 'slide.carried'),
			'a standalone script copies nothing, so it must not claim the slide was copied'
		)
		assert(shapeNames(output.slides[0].shapes).includes('Survivor'), 'the rest of the slide survives')

		// The other tier keeps the slide whole, which is the reason the recommendation exists.
		assert(
			printScript(ir).notes.some((note) => note.construct === 'slide.carried'),
			'the template-anchored tier still copies the slide'
		)
	})
})

describe('standalone printer — the fidelity contract', () => {
	test('every shape missing from the output is named by a `dropped` note', async () => {
		const source = await Presentation.load(await readFile(path.join(FIXTURES, 'mixed.pptx')))
		const { printed, output } = await runStandalone(await readFile(path.join(FIXTURES, 'mixed.pptx')))
		const dropped = new Set(
			printed.notes.filter((note) => note.disposition === 'dropped').map((note) => note.shapeName)
		)

		for (const [index, slide] of source.slides.entries()) {
			const survivors = new Set(shapeNames(output.slides[index]?.shapes ?? []))
			for (const name of shapeNames(slide.shapes)) {
				if (survivors.has(name)) continue
				assert(dropped.has(name), `slide ${index + 1}: ${JSON.stringify(name)} vanished with no note`)
			}
		}
	})

	test('the header block lists every applicable note', async () => {
		const printed = printStandaloneScript(await irFor('mixed.pptx'))
		assert(printed.notes.length > 0, 'this fixture has losses to declare')
		const header = printed.code.slice(0, printed.code.indexOf('*/'))
		for (const note of printed.notes) {
			assert(header.includes(note.construct), `header omits ${note.construct}`)
		}
	})

	test('the chrome losses the other tier rescues are all declared here', async () => {
		// The honest result of this tier, stated as an assertion rather than as prose: rebuilding
		// a deck's design from the read model loses the theme's format scheme, the master's text
		// styles, the layouts' decoration and their placeholder definitions, and every one of
		// those has to say so.
		const printed = printStandaloneScript(await irFor('mixed.pptx'))
		const constructs = new Set(printed.notes.map((note) => note.construct))
		for (const construct of [
			'theme.fmtScheme',
			'master.txStyles',
			'master.decoration',
			'master.placeholders',
			'master.default',
			'deck.docProps',
		]) {
			assert(constructs.has(construct), `a standalone script must declare ${construct}`)
		}
	})
})

describe('standalone printer — printed text', () => {
	// Committed so codegen churn is visible in review rather than only as a passing suite.
	// Regenerate with `vitest run script-standalone -u` after an intentional printer change.
	for (const name of ['empty.pptx', 'hidden.pptx']) {
		test(`${name} prints the same source as last time`, async () => {
			await expect(printStandaloneScript(await irFor(name)).code).toMatchFileSnapshot(
				path.join(__dirname, 'snapshots', `${name.replace(/\.pptx$/, '')}.standalone.ts.txt`)
			)
		})
	}
})

describe('standalone round trip — the corpus', () => {
	test('every fixture regenerates with no undeclared difference', { timeout: 300_000 }, async () => {
		const failures = []
		for (const name of fixtureNames) {
			const { report } = await runStandalone(await readFile(path.join(FIXTURES, name)))
			for (const difference of report.undeclared.slice(0, 5)) {
				failures.push(
					`${name} slide ${difference.slideNumber} ${difference.shapeName ?? '—'} ` +
						`${difference.path}: ${difference.expected} → ${difference.actual} [${difference.kind}]`
				)
			}
		}
		assertEqual(failures.join('\n'), '', 'undeclared round-trip differences')
	})
})
