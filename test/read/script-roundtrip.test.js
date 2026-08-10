// The round-trip oracle for `ts-pptx/script`: source deck → IR₁ → script → run it →
// output deck → IR₂, then diff IR₁ against IR₂ with the printer's fidelity notes as the
// exclusion list.
//
// **What makes this stronger than the tests before it.** `script-ir.test.js` checks the
// mapping against the write API's own types, and `script-print.test.js` checks that the
// emitted script runs. Neither can tell whether the deck that comes out is the deck that
// went in. This can, and building it found nine defects that both of those passed:
// paragraph bullets read as literal glyphs (`'none'` became an `n`), placeholders emitted
// with no geometry at all (a zero-height box in the corner), a group's rotation applied
// twice, image crops fed to an option that reads inches, every text body re-anchored to
// centre, every uncoloured run repainted black, PowerPoint text boxes demoted to auto
// shapes, an SVG picture reduced to its raster fallback, and a bulleted paragraph split in
// two. Every one produced a script that typechecked and ran.
//
// **What it cannot do, which is the part worth remembering.** The check is *asymmetry*
// detection. Both IRs come from the same reader and the same mapper, so a construct the
// converter never emits is absent from both sides and compares equal. Mutation testing says
// so out loud: deleting the `flipH` mapping, or the text-box detection, leaves this suite
// green, because the output deck then lacks the same thing the IR does. Those belong to
// `script-ir.test.js`, whose expectations come from `src/types/*.ts` rather than from the
// converter. Read a clean run here as "nothing the converter can distinguish was lost",
// never as "nothing was lost".
//
// `pnpm run script:roundtrip` runs the same comparison as a report, with `--verbose` for
// per-difference detail and `--dir` to point it at a corpus of real decks.

import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'
import { describe, expect, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import {
	canonicalDeckIr,
	diffDeckIr,
	isKnownNoteConstruct,
	knownNoteConstructs,
	printScript,
	printStandaloneScript,
	readModelToIr,
} from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'

const run = promisify(execFile)
const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = path.join(__dirname, 'fixtures')
const REPO = path.join(__dirname, '..', '..')

// Inside the repo, not the OS temp directory: the emitted script imports this package by
// its published name, which Node resolves by the self-reference rule only from a path
// underneath the package root. `/.tmp/` is gitignored.
const SCRATCH = path.join(REPO, '.tmp')

const fixtureNames = (await readdir(FIXTURES)).filter((name) => name.endsWith('.pptx')).sort()

/** Print a script for one fixture, execute it, and read the deck it wrote back into an IR. */
async function roundTrip(fixtureName) {
	await mkdir(SCRATCH, { recursive: true })
	const dir = await mkdtemp(path.join(SCRATCH, 'roundtrip-'))
	try {
		const bytes = await readFile(path.join(FIXTURES, fixtureName))
		const ir = readModelToIr(await Presentation.load(bytes))
		const printed = printScript(ir)

		await writeFile(path.join(dir, 'script.ts'), printed.code)
		// The template is the source deck unchanged — `fromTemplate` strips its slides itself.
		await writeFile(path.join(dir, 'template.pptx'), bytes)
		if (printed.assets.size > 0) {
			await mkdir(path.join(dir, 'assets'), { recursive: true })
			for (const [name, data] of printed.assets) await writeFile(path.join(dir, 'assets', name), data)
		}

		// Node's type stripping, so the script under test is the exact text a user is handed.
		await run(process.execPath, ['--no-warnings', path.join(dir, 'script.ts')])
		const output = readModelToIr(await Presentation.load(await readFile(path.join(dir, 'output.pptx'))))
		return { ir, printed, output, report: diffDeckIr(canonicalDeckIr(ir), canonicalDeckIr(output), printed.notes) }
	} finally {
		await rm(dir, { recursive: true, force: true })
	}
}

describe('script round trip — the corpus', () => {
	test('every fixture regenerates with no undeclared difference', { timeout: 300_000 }, async () => {
		const failures = []
		for (const name of fixtureNames) {
			const { report } = await roundTrip(name)
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

describe('script round trip — the oracle has teeth', () => {
	// Every check in this plan that measured nothing looked exactly like a passing one, so
	// the diff is asked to fail on demand before its passing is worth anything.
	test('a perturbed output IR produces an undeclared difference', async () => {
		const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, 'textbox.pptx'))))
		const before = canonicalDeckIr(ir)
		const after = canonicalDeckIr(ir)

		const clean = diffDeckIr(before, after, [])
		assertEqual(clean.differences.length, 0, 'a deck compared against itself must be identical')

		// Drop one run's bold flag, the smallest change a real regression could be.
		const runs = /** @type {any[]} */ (after.slides[0].calls.find((call) => call.method === 'addText')?.args[0])
		const bold = runs.find((item) => item.options?.bold === true)
		assert(bold, 'the fixture should contain a bold run to perturb')
		delete bold.options.bold

		const dirty = diffDeckIr(before, after, [])
		assertEqual(dirty.undeclared.length, 1, 'the dropped flag must be reported, and exactly once')
		assertEqual(dirty.undeclared[0].field, 'bold', 'reported against the option that changed')
		assertEqual(dirty.undeclared[0].kind, 'lost', 'and in the direction it changed')
	})

	test('a perturbed transition is caught, which is what putting it in the projection buys', async () => {
		// A slide's transition is set by property assignment rather than by a method call, so it
		// is invisible to the call comparison that catches everything else — a printer that
		// stopped emitting transitions entirely would produce identical `calls` and report a
		// clean round trip. `CanonicalSlide.transition` closes that, and this is the check that
		// it is closed rather than merely declared.
		const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, 'slide-transition.pptx'))))
		const before = canonicalDeckIr(ir)
		const after = canonicalDeckIr(ir)
		assertEqual(diffDeckIr(before, after, []).differences.length, 0, 'a deck compared against itself is identical')

		// Slide 2 is `push` with `dir="d"`: flip the direction only, the smallest change that
		// still renders differently, and one a whole-value compare would report far too coarsely.
		const pushed = /** @type {any} */ (after.slides[1].transition)
		pushed.variant.dir = 'u'
		const report = diffDeckIr(before, after, [])
		assertEqual(report.undeclared.length, 1, 'the flipped direction must be reported, and exactly once')
		assertEqual(report.undeclared[0].field, 'dir', 'reported against the variant attribute that changed')
		assertEqual(report.undeclared[0].slideNumber, 2, 'and against the slide it changed on')
	})

	test('a lost transition sound is reported narrowly enough that only its own note excuses it', async () => {
		// The reason `transition` is compared structurally rather than as one opaque value. The
		// template-anchored tier keeps the transition and drops its embedded sound, so if the
		// difference landed on `transition` the note declaring the sound would have to name that
		// field — and would then also excuse a transition whose type or duration came back wrong.
		const source = await Presentation.load(await readFile(path.join(FIXTURES, 'slide-transition-sound.pptx')))
		const ir = readModelToIr(source)
		const before = canonicalDeckIr(ir)
		const after = canonicalDeckIr(ir)
		const silenced = /** @type {any} */ (after.slides[0].transition)
		delete silenced.sound

		const undeclared = diffDeckIr(before, after, []).undeclared
		assertEqual(undeclared.length, 1, 'the missing sound is one difference')
		assertEqual(undeclared[0].field, 'sound', 'reported one level inside the transition, not on the transition')

		/** @type {(construct: string) => import('../../dist/script.js').FidelityNote[]} */
		const note = (construct) => [{ slideNumber: 1, shapeName: null, construct, disposition: 'dropped', cause: 'unsupported', detail: '' }] // prettier-ignore
		assertEqual(diffDeckIr(before, after, note('slide.transitionSound')).undeclared.length, 0, 'its own note declares it') // prettier-ignore
		assertEqual(
			diffDeckIr(before, after, note('slide.transition')).undeclared.length,
			1,
			'the note for a wholly dropped transition must not also excuse a dropped sound'
		)
	})

	test('a note declaring that field silences it, and only that field', async () => {
		const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, 'textbox.pptx'))))
		const before = canonicalDeckIr(ir)
		const after = canonicalDeckIr(ir)
		const call = after.slides[0].calls.find((c) => c.method === 'addText')
		const bold = /** @type {any[]} */ (call.args[0]).find((item) => item.options?.bold === true)
		delete bold.options.bold

		// `line.width` names a different construct, so it must not cover a lost `bold`.
		const wrong = /** @type {import('../../dist/script.js').FidelityNote[]} */ ([
			{ slideNumber: 1, shapeName: call.shapeName, construct: 'line.width', disposition: 'dropped', cause: 'unread', detail: '' }, // prettier-ignore
		])
		assertEqual(diffDeckIr(before, after, wrong).undeclared.length, 1, 'an unrelated note must not excuse it')
	})
})

describe('script round trip — the note coverage table', () => {
	test('every construct the corpus emits has a field mapping, in both tiers', async () => {
		// A note whose construct is absent from the table excludes nothing, so a typo or a new
		// note silently turns the round trip back into a snapshot. Checked against the notes the
		// corpus actually produces rather than a hand-kept list — and against *both* printers,
		// since each suppresses notes the other keeps and adds ones of its own.
		//
		// `isKnownNoteConstruct` rather than the raw key set, because a `layout.`-prefixed
		// construct resolves through its slide twin: a layout's decoration is transcribed by the
		// slide shape mapper, so `layout.line.width` is `line.width`'s entry seen from the chrome.
		const missing = new Set()
		for (const name of fixtureNames) {
			const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, name))))
			for (const note of [...printScript(ir).notes, ...printStandaloneScript(ir).notes]) {
				if (!isKnownNoteConstruct(note.construct)) missing.add(note.construct)
			}
		}
		assertEqual([...missing].sort().join(', '), '', 'note constructs with no entry in the coverage table')
	})

	test('the table itself is the only thing `isKnownNoteConstruct` accepts', async () => {
		// The fallback above widens what counts as known, so it needs a floor: without this a
		// prefix rule that accepted anything would make the check above vacuous.
		const known = new Set(knownNoteConstructs())
		assert(known.has('line.width'), 'the table lists the slide constructs')
		assert(isKnownNoteConstruct('layout.line.width'), 'and resolves the layout spelling of one')
		assert(!isKnownNoteConstruct('layout.not.a.construct'), 'but invents no entry for a construct it lacks')
		assert(!isKnownNoteConstruct('line.widht'), 'and a typo stays unknown')
	})

	test('the two tiers suppress different notes, and neither suppresses one it causes', async () => {
		// The suppression lists are the place a caveat quietly disappears. Pinned on the fixture
		// that exercises the most chrome: the template-anchored tier must not report the chrome
		// losses (it never rebuilds the chrome), and the standalone tier must report every one.
		const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, 'mixed.pptx'))))
		const anchored = new Set(printScript(ir).notes.map((note) => note.construct))
		const standalone = new Set(printStandaloneScript(ir).notes.map((note) => note.construct))

		for (const construct of ['theme.fmtScheme', 'master.txStyles', 'master.placeholders', 'deck.docProps']) {
			assert(
				ir.fidelity.some((note) => note.construct === construct),
				`the IR should declare ${construct} — it is a real loss for a standalone output`
			)
			assert(!anchored.has(construct), `a template-anchored script must not report ${construct} as lost`)
			assert(standalone.has(construct), `a standalone script must report ${construct} as lost`)
		}
		// And the mirror: the standalone tier copies no slide, so it must not claim it will.
		assert(!standalone.has('slide.carried'), 'a standalone script cannot carry a slide, so it must not say it does')
	})
})

describe('script round trip — canonicalisation is an equivalence', () => {
	// The canonicaliser is the one place a wrong rule hides a defect permanently, so its rules
	// are tested in both directions: the default is dropped, and a real value never is.
	const deck = (options) => ({
		slideSize: { widthEmu: 1, heightEmu: 1 },
		props: {},
		chrome: { theme: {}, masters: [] },
		assets: [],
		fidelity: [],
		slides: [{ number: 1, source: 'authored', layout: null, hidden: false, calls: [{ method: 'addText', args: [[], options] }] }], // prettier-ignore
	})
	const optionsOf = (ir) => canonicalDeckIr(ir).slides[0].calls[0].args[1]

	test('OOXML defaults spelled out compare equal to omitted', () => {
		assertEqual(JSON.stringify(optionsOf(deck({ bold: false, rotate: 0, flipH: false }))), '{}', 'defaults dropped')
	})

	test('a value that is not the default always survives', () => {
		const kept = optionsOf(deck({ bold: true, rotate: 90, flipH: true, width: 1 }))
		assertEqual(JSON.stringify(kept), '{"bold":true,"rotate":90,"flipH":true,"width":1}', 'real values kept')
	})

	test('line width is not treated as a default, because 1pt is not OOXML’s', () => {
		assertEqual(JSON.stringify(optionsOf(deck({ width: 1 }))), '{"width":1}', 'a 1pt outline is a real outline')
	})

	test('assets compare by content, so a renamed image is still the same image', async () => {
		const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, 'image.pptx'))))
		const renamed = structuredClone(ir)
		for (const asset of renamed.assets) asset.name = `renamed-${asset.name}`
		// A `seen` set, because one `AssetRef` object is shared by every call using that image
		// and `structuredClone` preserves that sharing — renaming per visit renames it twice.
		const seen = new WeakSet()
		const rename = (value) => {
			if (Array.isArray(value)) return void value.forEach(rename)
			if (value === null || typeof value !== 'object' || seen.has(value)) return
			seen.add(value)
			if (typeof value.$asset === 'string') value.$asset = `renamed-${value.$asset}`
			else Object.values(value).forEach(rename)
		}
		rename(renamed.slides)
		const report = diffDeckIr(canonicalDeckIr(ir), canonicalDeckIr(renamed), [])
		assertEqual(report.differences.length, 0, 'renaming an asset must not read as a different image')
	})
})

describe('script round trip — printed text', () => {
	// Committed so codegen churn is visible in review rather than only as a passing suite.
	// Regenerate with `vitest run script-roundtrip -u` after an intentional printer change.
	for (const name of ['empty.pptx', 'hidden.pptx']) {
		test(`${name} prints the same source as last time`, async () => {
			const ir = readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, name))))
			await expect(printScript(ir).code).toMatchFileSnapshot(
				path.join(__dirname, 'snapshots', `${name.replace(/\.pptx$/, '')}.script.ts.txt`)
			)
		})
	}
})
