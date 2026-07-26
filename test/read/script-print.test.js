// The template-anchored printer — `ts-pptx/script`'s write half.
//
// The organising fact of this file: **printing a script that typechecks proves almost
// nothing about it**. Every defect found while building this printer — a chart type passed
// positionally instead of as an option, a connector given a bounding box instead of two
// endpoints, an ambiguous layout name, and an import specifier that was simply not this
// package's name — produced text that read perfectly and failed the moment it ran. The IR
// is `IrValue`-typed by design, so `tsc` cannot check an argument against the write API
// signature it is meant to satisfy. Running the output is the type system here.
//
// So the tests that matter below execute the emitted script in a subprocess and read the
// deck it produces. The pure-text assertions are a cheap first filter, not the point.
//
// The strongest of them is the *undeclared-loss* check, which is the fidelity contract
// stated as an oracle: every source shape absent from the output must be named by a note
// with disposition `dropped`. A shape that vanishes quietly fails it. That is the same
// shape of check the full round-trip harness will apply field-by-field.

import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation, isGraphicFrame } from '../../dist/read.js'
import { printScript, readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'

const run = promisify(execFile)
const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = path.join(__dirname, 'fixtures')
const REPO = path.join(__dirname, '..', '..')

// Inside the repo, not the OS temp directory, and deliberately so: the emitted script
// imports this package by its published name, which resolves only from a path underneath
// the package root (Node's self-reference rule). `/.tmp/` is gitignored.
const SCRATCH = path.join(REPO, '.tmp')

const fixtureNames = (await readdir(FIXTURES)).filter((name) => name.endsWith('.pptx')).sort()

async function irFor(name) {
	return readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, name))))
}

/** Every shape name on a slide, descending into groups. */
function shapeNames(shapes, out = []) {
	for (const shape of shapes) {
		out.push(shape.name)
		if (shape.shapes) shapeNames(shape.shapes, out)
	}
	return out
}

/**
 * Lay a printed script out on disk, run it, and load the deck it wrote.
 *
 * The template is the fixture itself, unmodified — the whole premise of this tier is that
 * `fromTemplate` strips a deck's slides for you, so the source deck and the template asset
 * are the same bytes.
 */
async function runPrinted(fixtureName, options = {}) {
	await mkdir(SCRATCH, { recursive: true })
	const dir = await mkdtemp(path.join(SCRATCH, 'script-print-'))
	try {
		const bytes = await readFile(path.join(FIXTURES, fixtureName))
		const printed = printScript(readModelToIr(await Presentation.load(bytes)), options)

		await writeFile(path.join(dir, 'script.ts'), printed.code)
		await writeFile(path.join(dir, 'template.pptx'), bytes)
		if (printed.assets.size > 0) {
			await mkdir(path.join(dir, 'assets'), { recursive: true })
			for (const [name, data] of printed.assets) await writeFile(path.join(dir, 'assets', name), data)
		}

		// Type stripping, so the emitted TypeScript runs as-is rather than through a build.
		await run(process.execPath, ['--no-warnings', path.join(dir, 'script.ts')])
		const output = await Presentation.load(await readFile(path.join(dir, 'output.pptx')))
		return { printed, output }
	} finally {
		await rm(dir, { recursive: true, force: true })
	}
}

describe('script printer — corpus invariants', () => {
	test('every fixture prints, and prints identically twice', async () => {
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			const first = printScript(ir)
			const second = printScript(ir)
			assert(first.code.length > 0, `${name}: printed nothing`)
			assertEqual(first.code, second.code, `${name}: printing is not deterministic`)
		}
	})

	test('no printed value is the literal `undefined`', async () => {
		// The IR forbids `undefined` so that "absent" has one spelling. If one ever leaks in,
		// it prints as a bare identifier and the script throws or silently passes garbage.
		for (const name of fixtureNames) {
			const { code } = printScript(await irFor(name))
			const offender = code.split('\n').find((line) => /(?<![A-Za-z0-9_$])undefined(?![A-Za-z0-9_$])/.test(line))
			assertEqual(offender, undefined, `${name}: printed an \`undefined\``)
		}
	})

	test('every asset reference resolves to a declared binding', async () => {
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			const { code, assets } = printScript(ir)
			assertEqual(assets.size, ir.assets.length, `${name}: asset count`)
			for (const asset of ir.assets) {
				const identifier = asset.name.replace(/\.[^.]*$/, '')
				assert(code.includes(`const ${identifier} = `), `${name}: no binding for ${asset.name}`)
				assert(code.includes(`data: ${identifier}`), `${name}: ${asset.name} is declared but never used`)
			}
		}
	})

	test('the emitted import specifier is this package’s published name', async () => {
		// This shipped wrong once: the directory is `ts-pptx`, the package is not. The script
		// printed, typechecked, and then failed at `import` — so the name is pinned to the
		// manifest rather than to a literal anyone can re-guess.
		const manifest = JSON.parse(await readFile(path.join(REPO, 'package.json'), 'utf8'))
		const { code } = printScript(await irFor('empty.pptx'))
		assert(code.includes(`import TsPptx from '${manifest.name}'`), `write-half import is not ${manifest.name}`)
		assert(code.includes(`from '${manifest.name}/read'`), `read-half import is not ${manifest.name}/read`)
	})
})

describe('script printer — the emitted script runs', () => {
	test('a plain deck round-trips its slides, hidden flag and text', async () => {
		const { output } = await runPrinted('hidden.pptx')
		assertEqual(output.slides.length, 2, 'slide count')
		assertEqual(output.slides[1].hidden, true, 'slide 2 stays hidden')
		assert(output.slides[0].shapes.length > 0, 'slide 1 has content')
	})

	test('a chart survives, which needs `type` inside the options object', async () => {
		// `addChart(data, { type })` — passing the type positionally throws at run time and is
		// invisible to every static check, since the IR types args as `IrValue[]`.
		const { output } = await runPrinted('bar-chart-data-labels.pptx')
		const frames = output.slides[0].shapes.filter(isGraphicFrame).filter((shape) => shape.chart)
		assertEqual(frames.length, 1, 'the chart was regenerated')
	})

	test('connectors keep their endpoints, including the flipped diagonal', async () => {
		// OOXML gives a connector a box plus flip flags; addConnector takes two points. Get the
		// flips wrong and every up-or-leftward connector is silently mirrored.
		const source = await Presentation.load(await readFile(path.join(FIXTURES, 'mixed.pptx')))
		const { output } = await runPrinted('mixed.pptx')
		const connectorsOf = (pres) =>
			pres.slides
				.flatMap((slide) => slide.shapes)
				.filter((shape) => shape.constructor.name === 'Connector')
				.map((shape) => {
					const frame = shape.absoluteFrame
					return `${frame.left},${frame.top},${frame.width},${frame.height},${shape.flipH},${shape.flipV}`
				})
				.sort()
		assertEqual(connectorsOf(output).join(' | '), connectorsOf(source).join(' | '), 'connector geometry')
	})

	test('a deck whose layout names repeat still binds, by gallery position', async () => {
		// `appendSlides` throws on an ambiguous layout name rather than picking one, and a
		// multi-master deck routinely carries two layouts called "Title and Text".
		const { printed, output } = await runPrinted('read-stress.pptx')
		assert(printed.code.includes('deck.layouts()['), 'fell back to positional binding')
		assertEqual(output.slides.length, 2, 'slide count')
	})

	test('hostile text survives escaping, byte for byte', async () => {
		// Deck text is arbitrary user content, and the printer's escaping is the only thing
		// between it and a source file that does not parse. Every character here is one a real
		// deck can contain: quotes of both kinds, a backslash, a backtick, a hard line break
		// inside a run, both Unicode separators, an astral-plane emoji, and an accent.
		const hostile =
			'quote \u0027 backslash \\ backtick \u0060 newline\n sep \u2028 para \u2029 emoji \u{1f600} \u00e9 quote \u0022'
		// A real IR with its slides swapped out, so the slide size still matches the template
		// that `appendSlides` will compare it against.
		const base = await irFor('empty.pptx')
		/** @type {import('../../dist/script.js').SlideIr} */
		const slide = {
			number: 1,
			source: 'authored',
			layout: null,
			hidden: false,
			calls: [{ method: 'addText', args: [[{ text: hostile }], { x: '0emu', y: '0emu', w: '4572000emu' }] }],
		}
		const ir = { ...base, slides: [slide] }

		await mkdir(SCRATCH, { recursive: true })
		const dir = await mkdtemp(path.join(SCRATCH, 'script-print-'))
		try {
			await writeFile(path.join(dir, 'script.ts'), printScript(ir).code)
			await writeFile(path.join(dir, 'template.pptx'), await readFile(path.join(FIXTURES, 'empty.pptx')))
			await run(process.execPath, ['--no-warnings', path.join(dir, 'script.ts')])
			const output = await Presentation.load(await readFile(path.join(dir, 'output.pptx')))

			// U+2028/U+2029 are the exception, and the deviation is downstream of the printer:
			// they reach the write path intact (proving the escaping worked — an unescaped one
			// would have made the file unparseable) and it renders them as real line breaks,
			// which is what a `.pptx` can actually hold. Everything else must survive verbatim.
			const expected = hostile.replaceAll('\u2028', '\n').replaceAll('\u2029', '\n')
			assertEqual(output.slides[0].shapes[0].textFrame.text, expected, 'text survived the round trip')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})
})

describe('script printer — the fidelity contract', () => {
	test('every shape missing from the output is named by a `dropped` note', async () => {
		// The contract in one assertion: a note is a promise that a construct will not come
		// back, so anything that fails to come back without one is a defect. Run over the
		// richest fixture, which drops shapes for four different reasons.
		const source = await Presentation.load(await readFile(path.join(FIXTURES, 'mixed.pptx')))
		const { printed, output } = await runPrinted('mixed.pptx')
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

	test('document properties are not reported lost, because the template carries them', async () => {
		// The read half notes that only five of the twelve docProps have write-API setters.
		// True of a standalone output; false here, where none of them is ever authored. A
		// caveat that does not apply teaches the reader to skim the ones that do.
		const ir = await irFor('mixed.pptx')
		const printed = printScript(ir)
		assert(
			ir.fidelity.some((note) => note.construct === 'deck.docProps'),
			'the IR should still declare the loss — it is real for a standalone output'
		)
		assert(
			!printed.notes.some((note) => note.construct === 'deck.docProps'),
			'a template-anchored script must not report docProps as lost'
		)
		assert(!printed.code.includes('deck.docProps'), 'the header must not carry the note either')
	})

	test('the header block lists every applicable note', async () => {
		const ir = await irFor('custgeom.pptx')
		const printed = printScript(ir)
		assert(printed.notes.length > 0, 'this fixture has losses to declare')
		const header = printed.code.slice(0, printed.code.indexOf('*/'))
		for (const note of printed.notes) {
			assert(header.includes(note.construct), `header omits ${note.construct}`)
		}
	})
})
