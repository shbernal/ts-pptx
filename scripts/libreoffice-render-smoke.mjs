#!/usr/bin/env node
/**
 * LibreOffice render smoke — a second render oracle, on any machine that has LibreOffice.
 *
 * Every other check in this repo is either a byte assertion or PowerPoint. That leaves one
 * class of regression invisible: markup that is byte-correct and that PowerPoint renders
 * correctly *because PowerPoint recomputes it*. The canonical case is SmartArt. A deck
 * stores every drawn string twice — once in the `dgm:dataModel` that PowerPoint reads, and
 * once in the `dsp:drawing` cache that every renderer without a SmartArt layout engine
 * paints — and PowerPoint regenerates the cache from the data model on open. So PowerPoint
 * cannot observe whether the cache was written: it looks right either way. Reading the
 * bytes of `ppt/diagrams/drawing1.xml` is real evidence, but it is evidence about a part,
 * not about a pixel.
 *
 * LibreOffice is the oracle that disagrees with PowerPoint by design. It has no SmartArt
 * layout engine, so it paints the cache and nothing else, which makes it the only renderer
 * available here that can tell a written cache from an unwritten one. It is already used in
 * this repo as an independent *measurement* oracle — `test/read/fixtures/authoring/measure-lo.py`
 * reads LibreOffice's recomputed `spAutoFit` sizes back over UNO — and this is the same
 * second opinion applied to what actually gets drawn.
 *
 * **How it reads text back.** `--convert-to pdf`, then `pdftotext`. PDF export runs through
 * LibreOffice's drawing layer, the same code path as the screen render, so a string in the
 * PDF is a string that was painted. Going via PDF also beats `--convert-to png`, and not
 * only because it avoids OCR: LibreOffice's PNG export writes the *first slide only* and
 * ignores a `PageRange` filter option, so a multi-slide fixture is unreachable that way.
 *
 * **Every run carries its own sensitivity check.** `mirrored` proves the new string reaches
 * the renderer; `stale` edits the data model alone and proves the renderer keeps painting
 * the *old* string. If the pair ever both pass trivially — the deck stopped being rendered,
 * the text stopped being extracted — `stale` is what goes red. Each case also asserts that
 * the ten sibling nodes it did not touch are still painted intact, which is what separates
 * "the mirror wrote the right string" from "the mirror wrote *a* string".
 *
 * **The second family: zero-baseline constructs.** SmartArt is not the only thing this
 * oracle can see. `byte-identity` only covers what the showcase decks emit, so a construct
 * the corpus never reaches has no evidence behind it beyond "the right bytes are in the
 * part" — true of `numCol`/`spcCol`, `a:prstTxWarp`, `a:buBlip`, `a:buClr`, `rtl="1"` and
 * `altLang`. Rendering answers a question the bytes cannot: whether an independent
 * implementation *acts* on the markup we emit. `PAIRS` below carries the three of those six
 * that a text-extraction oracle can actually observe; the header on `PAIRS` records what was
 * measured for the other three and why they are not here, so the probe is not run twice.
 *
 * Each pair renders the construct and an otherwise identical control, and asserts that the
 * two extractions DIFFER while both still paint their canary strings. Differential, never
 * absolute: the canary is the sensitivity check (a blank render fails instead of passing),
 * and comparing two decks from the same run keeps the check independent of which pdftotext
 * build resolved — xpdf and poppler disagree on details like the codepoint a bullet glyph
 * extracts as, but they agree on whether two renders came out the same.
 *
 *   node scripts/libreoffice-render-smoke.mjs           # run every case
 *   node scripts/libreoffice-render-smoke.mjs --keep    # ...and keep the decks, PDFs and text
 *
 * Requirements: LibreOffice and `pdftotext` (xpdf-tools or poppler-utils). Both are looked
 * up on PATH, then at the usual per-platform locations, and both can be pointed at
 * explicitly with `TSPPTX_SOFFICE` / `TSPPTX_PDFTOTEXT`. Missing either is a clean SKIP on a
 * workstation, where not every machine has LibreOffice — but a failure under
 * `TSPPTX_RENDER_ORACLE=required`, which is how the CI lane runs it, because a lane that
 * installs the tools and then skips would be green without checking anything.
 *
 * It does run in CI, on `ubuntu-latest`, unlike the COM smoke it sits beside. Measured on a
 * hosted runner: `libreoffice-impress` plus `poppler-utils` install in ~41s from apt with no
 * caching, and the whole check then runs in ~2s. Neither the missing Microsoft fonts (Arial
 * resolves to Liberation Sans there) nor poppler's disagreements with xpdf disturb it — see
 * the note on `PAIRS` for why the assertions are differential rather than absolute.
 */
import os from 'node:os'
import fs from 'node:fs/promises'
import path from 'node:path'
import { spawn, spawnSync } from 'node:child_process'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { parseCliOrExit } from './script-utils.mjs'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const ROOT = path.resolve(__dirname, '..')

// --- args -------------------------------------------------------------------
const USAGE = `LibreOffice render smoke — paint decks in a renderer with no SmartArt engine.

  pnpm run test:lo
  pnpm run test:lo -- --keep

Options:
  --keep          leave the generated decks, PDFs and extracted text on disk
  -h, --help      show this message

Environment:
  TSPPTX_SOFFICE         path to soffice (skips the search)
  TSPPTX_PDFTOTEXT       path to pdftotext (skips the search)
  TSPPTX_RENDER_ORACLE   set to "required" to fail, not SKIP, when a tool is missing`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: { keep: { type: 'boolean', default: false } },
})
const KEEP = values.keep

// LibreOffice can take a while on a cold profile; a hung conversion must still end the run.
const CONVERT_TIMEOUT_MS = 180_000

// --- the fixture and what it says -------------------------------------------
/** `mixed.pptx` slide 2 holds a PowerPoint-authored `hList1` diagram. */
const FIXTURE = path.join(ROOT, 'test', 'read', 'fixtures', 'mixed.pptx')
const DIAGRAM_PAGE = 2

/** The node every case edits. */
const ANCHOR = 'Understand Data Modelling Principles'

/**
 * The ten nodes no case edits. Asserted painted in *every* case, so a mirror that writes
 * the right string into the wrong point fails even though its own sentinel shows up.
 */
const NEIGHBOURS = [
	'Uncontrollable Inputs (environmental factors)',
	'Controllable Inputs (decision variables)',
	'Mathematical model',
	'Be able to create complex models using spreadsheets',
	'Create models, simulate changes of uncontrollable inputs, see the impact on results',
	'Graph the results',
	'Be able to use Pivot Tables',
	'Use probabilities in models',
	'Simulate uncontrollable inputs using probability distribution',
	'Conduct simulations',
]

const MIRRORED = 'ZZMIRROREDZZ'
const STALE = 'ZZSTALEZZ'

/**
 * One deck to build, render and read back.
 * @typedef {object} Case
 * @property {string} label
 * @property {string} claim what a pass proves, printed with the result
 * @property {number} page the PDF page the diagram lands on
 * @property {((point: any) => void) | null} edit applied to the anchor point, or null for the control
 * @property {string[]} expect strings that must appear in the painted text
 * @property {string[]} reject strings that must not
 */
/** @type {Case[]} */
const CASES = [
	{
		label: 'baseline',
		claim: 'LibreOffice paints the diagram at all',
		page: DIAGRAM_PAGE,
		edit: null,
		expect: [ANCHOR],
		reject: [MIRRORED, STALE],
	},
	{
		label: 'mirrored',
		claim: 'DiagramPoint.text reaches a renderer with no SmartArt engine',
		page: DIAGRAM_PAGE,
		edit: (point) => {
			point.text = MIRRORED
		},
		expect: [MIRRORED],
		reject: [ANCHOR],
	},
	{
		label: 'stale',
		claim: 'DiagramPoint.textFrame edits the data model only, and is therefore invisible',
		page: DIAGRAM_PAGE,
		edit: (point) => {
			point.textFrame.text = STALE
		},
		expect: [ANCHOR],
		reject: [STALE],
	},
]

// --- the zero-baseline construct pairs --------------------------------------
/** Painted by every pair's control box. A pair whose canary is missing rendered nothing. */
const CANARY = 'ZZCANARYZZ'

/**
 * Body geometry for the pair decks, so `on` and `off` differ only in the construct.
 *
 * A factory and not a shared literal, which is load-bearing rather than style: `addText`
 * writes its own state (`_bodyProp`, `objectName`) onto the options object it is handed, so
 * two calls given shallow copies of one literal end up aliasing a single `_bodyProp`. That
 * is read at emit time, so `columns` set on one shape reaches every shape sharing it — which
 * is exactly how the first draft of the `numCol` control came out with `numCol="2"` in it.
 * @param {Record<string, unknown>} [extra]
 * @returns {Record<string, unknown>}
 */
const pairBody = (extra = {}) => ({ x: 0.6, y: 1.6, w: 8.5, h: 4.0, fontSize: 18, fontFace: 'Arial', ...extra })

/** A long enough single paragraph that a two-column body wraps visibly narrower. */
const FLOW =
	'Alpha node one. Beta node two. Gamma node three. Delta node four. ' +
	'Epsilon node five. Zeta node six. Eta node seven. Theta node eight.'

/** The four bulleted lines the bullet pairs paint. */
const BULLET_LINES = ['Alpha node one', 'Beta node two', 'Gamma node three', 'Delta node four']

/**
 * An 8x8 PNG, inline so the check needs no fixture on disk. Content is irrelevant — what
 * matters is that a picture bullet resolves to a rel and paints instead of a glyph.
 */
const BULLET_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAKUlEQVR4nGP8z8DwnwEPYMInOTiAsRs2MDIyMjIyMDAwMDAwMDAwMAAAlgQFAaBqvNsAAAAASUVORK5CYII='

/**
 * One construct, rendered against its own control.
 * @typedef {object} Pair
 * @property {string} label
 * @property {string} claim what a pass proves, printed with the result
 * @property {boolean} layout extract with `-layout` (physical layout) instead of reading order
 * @property {(slide: any) => void} on builds the body WITH the construct
 * @property {(slide: any) => void} off builds the same body WITHOUT it
 * @property {string[]} canaries strings both sides must paint
 * @property {(on: string, off: string) => string[]} verify extra assertions, one message per failure
 */

/**
 * The three zero-baseline constructs a text-extraction oracle can observe.
 *
 * The other three were probed the same way — construct deck against control deck, with the
 * emitted parts checked to confirm each construct was actually in the `on` deck and absent
 * from the `off` one — and are deliberately NOT here:
 *
 *   - `rtl="1"`  LibreOffice paints a Latin paragraph byte-for-byte identically with the
 *                flag set and unset (same rasterised page hash), so there is nothing to
 *                assert. With Hebrew text the page hash does change, but the extraction is
 *                identical either way because neither pdftotext build could map the Hebrew
 *                glyphs back to codepoints at all — evidence a *pixel* channel could use and
 *                this one cannot.
 *   - `altLang`  a language tag with no visual semantics. Correctly invisible to any
 *                renderer; not a gap in this oracle.
 *   - `a:buClr`  the one genuine near miss. Extraction is identical (a bullet's colour does
 *                not change its glyph) but the rasterised page differs, so it is reachable —
 *                only by adding a rasteriser channel. Not taken: it buys the weakest claim of
 *                the six ("something changed"), and it would put a second tool in the path
 *                whose name differs between the two builds (`pdftopng` on xpdf,
 *                `pdftoppm -png` on poppler) for the sake of one construct.
 *
 * @type {Pair[]}
 */
const PAIRS = [
	{
		label: 'buBlip',
		claim: '<a:buBlip> reaches the renderer — a picture bullet paints something a character bullet does not',
		layout: false,
		on: (slide) =>
			slide.addText(
				BULLET_LINES.map((text) => ({ text, options: { bullet: { image: { data: BULLET_PNG } } } })),
				pairBody()
			),
		off: (slide) =>
			slide.addText(
				BULLET_LINES.map((text) => ({ text, options: { bullet: true } })),
				pairBody()
			),
		// Both sides must still paint every line: that is what stops "the deck rendered blank"
		// from reading as "the bullets differ".
		canaries: [CANARY, ...BULLET_LINES],
		verify: (on, off) =>
			on === off ? ['a picture bullet extracted identically to a character bullet — buBlip was dropped'] : [],
	},
	{
		label: 'prstTxWarp',
		claim: '<a:prstTxWarp> reaches the renderer — warped text is painted as rotated glyphs, not as a line of text',
		layout: false,
		on: (slide) => slide.addText(FLOW, pairBody({ h: 2.5, textWarp: 'textArchUp' })),
		off: (slide) => slide.addText(FLOW, pairBody({ h: 2.5 })),
		// Only the canary, on purpose. The warped box is expected NOT to extract, so asserting
		// its text on both sides would fail the case it exists to prove.
		canaries: [CANARY],
		verify: (on, off) =>
			on === off ? ['warped text extracted identically to unwarped text — prstTxWarp was dropped'] : [],
	},
	{
		label: 'numCol',
		claim: 'numCol/spcCol reach the renderer — a two-column body wraps at a visibly narrower measure',
		// The whole signal is horizontal: in reading order both bodies extract the same words in
		// the same order, and only `-layout` preserves where the lines actually broke.
		layout: true,
		on: (slide) => slide.addText(FLOW, pairBody({ columns: 2, columnSpacing: 0.4 })),
		off: (slide) => slide.addText(FLOW, pairBody()),
		canaries: [CANARY],
		verify: (on, off) => {
			/** @param {string} text */
			const widest = (text) =>
				Math.max(0, ...text.split('\n').map((/** @type {string} */ line) => line.trimEnd().length))
			const onWidest = widest(on)
			const offWidest = widest(off)
			// A ratio, not a pinned column count: the two pdftotext builds lay out `-layout`
			// slightly differently, but a two-column measure is far narrower in both. Measured
			// here: 34 against 66 characters, a ratio of 0.52.
			if (!(onWidest < offWidest * 0.75)) {
				return [
					`two-column body wrapped at ${onWidest} chars against ${offWidest} for one column; ` +
						'expected well under 75% of it, so numCol looks dropped',
				]
			}
			return []
		},
	},
]

// --- finding the two tools --------------------------------------------------
/**
 * First existing path in `candidates`, or null. Blank entries are skipped, so a caller can
 * splice in a platform-specific guess without guarding it.
 * @param {(string | undefined)[]} candidates
 * @returns {Promise<string | null>}
 */
async function firstExisting(candidates) {
	for (const candidate of candidates) {
		if (!candidate) continue
		try {
			await fs.access(candidate)
			return candidate
		} catch {
			// Not here; try the next.
		}
	}
	return null
}

/**
 * An explicit override, validated.
 *
 * Deliberately *not* one more entry in the candidate list: someone who sets `TSPPTX_SOFFICE`
 * has said which binary to use, so a typo in it must be an error naming the path they gave.
 * Letting it fall through to the search would run a different binary than the one they asked
 * for, or SKIP as though they had never set anything.
 * @param {string} envVar
 * @returns {Promise<string | null>}
 * @throws when the variable is set and names nothing
 */
async function overrideFrom(envVar) {
	const value = process.env[envVar]
	if (!value) return null
	if (!(await firstExisting([value]))) throw new Error(`${envVar} is set to "${value}", which does not exist`)
	return value
}

/**
 * Resolve `name` through PATH, or null.
 * @param {string} name
 * @returns {string | undefined}
 */
function onPath(name) {
	const probe = spawnSync(process.platform === 'win32' ? 'where' : 'which', [name], {
		stdio: ['ignore', 'pipe', 'ignore'],
		encoding: 'utf8',
	})
	if (probe.status !== 0) return undefined
	const first = String(probe.stdout).split(/\r?\n/).find(Boolean)
	return first ? first.trim() : undefined
}

/**
 * The LibreOffice binary. On Windows this must be `soffice.com` and not `soffice.exe`: the
 * `.com` is the console front end and blocks until the conversion finishes, while the `.exe`
 * returns immediately and the script would read PDFs that are not written yet.
 * @returns {Promise<string | null>}
 */
async function findSoffice() {
	const local = process.env.LOCALAPPDATA
	const win = process.platform === 'win32'
	return (
		(await overrideFrom('TSPPTX_SOFFICE')) ??
		(await firstExisting([
			win ? onPath('soffice.com') : onPath('soffice'),
			win ? undefined : onPath('libreoffice'),
			// A no-admin administrative extract lands here; see the powerpoint-fixture-authoring skill.
			win && local ? path.join(local, 'Programs', 'LibreOffice', 'program', 'soffice.com') : undefined,
			win ? 'C:\\Program Files\\LibreOffice\\program\\soffice.com' : undefined,
			process.platform === 'darwin' ? '/Applications/LibreOffice.app/Contents/MacOS/soffice' : undefined,
			'/usr/bin/soffice',
			'/usr/bin/libreoffice',
			'/snap/bin/libreoffice',
		]))
	)
}

/** @returns {Promise<string | null>} */
async function findPdfToText() {
	const local = process.env.LOCALAPPDATA
	const win = process.platform === 'win32'
	return (
		(await overrideFrom('TSPPTX_PDFTOTEXT')) ??
		(await firstExisting([
			onPath(win ? 'pdftotext.exe' : 'pdftotext'),
			win && local ? path.join(local, 'Programs', 'xpdf-tools', 'bin', 'pdftotext.exe') : undefined,
			'/usr/bin/pdftotext',
			'/opt/homebrew/bin/pdftotext',
		]))
	)
}

// --- running them -----------------------------------------------------------
/**
 * Spawn and collect. Never rejects: the caller decides what a non-zero code means, because
 * LibreOffice writes noise to stderr on a perfectly good run (a bundled-Python warning,
 * `Could not find platform independent libraries`), and a script that read stderr as failure
 * would fail every time.
 * @param {string} command
 * @param {string[]} args
 * @param {number} timeoutMs
 * @returns {Promise<{code: number, out: string, err: string, timedOut: boolean}>}
 */
function runTool(command, args, timeoutMs) {
	return new Promise((resolve) => {
		const child = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] })
		let out = ''
		let err = ''
		let timedOut = false
		const timer = setTimeout(() => {
			timedOut = true
			child.kill()
		}, timeoutMs)
		child.stdout?.on('data', (d) => (out += d))
		child.stderr?.on('data', (d) => (err += d))
		child.on('close', (code) => {
			clearTimeout(timer)
			resolve({ code: code ?? -1, out, err, timedOut })
		})
		child.on('error', (e) => {
			clearTimeout(timer)
			resolve({ code: -1, out, err: String(e), timedOut })
		})
	})
}

/**
 * Convert every deck to PDF in one LibreOffice invocation.
 *
 * `-env:UserInstallation` is not optional. It gives this run a throwaway profile, which
 * keeps the check off the user's real LibreOffice settings and, more importantly, stops it
 * deadlocking against an interactive LibreOffice that already holds the default one.
 * @param {string} soffice
 * @param {string[]} decks
 * @param {string} outDir
 * @param {string} profileDir
 * @returns {Promise<string[]>} one message per failure
 */
async function convertToPdf(soffice, decks, outDir, profileDir) {
	const result = await runTool(
		soffice,
		[
			'--headless',
			'--norestore',
			'--nolockcheck',
			'-env:UserInstallation=' + pathToFileURL(profileDir).href,
			'--convert-to',
			'pdf',
			'--outdir',
			outDir,
			...decks,
		],
		CONVERT_TIMEOUT_MS
	)
	if (result.timedOut) return [`LibreOffice did not finish within ${CONVERT_TIMEOUT_MS / 1000}s and was killed`]
	if (result.code !== 0) return [`LibreOffice exited ${result.code}: ${(result.err || result.out).trim()}`]
	return []
}

/**
 * The text LibreOffice painted on one page, whitespace-normalized.
 *
 * Normalizing matters: a node whose text wraps in its box comes back split across lines, so
 * `Be able to use Pivot Tables` is only contiguous once runs of whitespace collapse.
 * Pass `layout` to keep the physical layout instead, for a case whose signal is *where* the
 * lines broke. That output is left unnormalized apart from trailing blank lines — collapsing
 * whitespace there would destroy the only thing it was extracted for.
 * @param {string} pdftotext
 * @param {string} pdf
 * @param {number} page
 * @param {boolean} [layout]
 * @returns {Promise<{text: string, failures: string[]}>}
 */
async function paintedText(pdftotext, pdf, page, layout = false) {
	const txt = pdf.replace(/\.pdf$/, layout ? '.layout.txt' : '.txt')
	const args = [...(layout ? ['-layout'] : []), '-f', String(page), '-l', String(page), pdf, txt]
	const result = await runTool(pdftotext, args, CONVERT_TIMEOUT_MS)
	if (result.code !== 0) {
		return { text: '', failures: [`pdftotext exited ${result.code}: ${(result.err || result.out).trim()}`] }
	}
	try {
		const raw = await fs.readFile(txt, 'utf8')
		if (layout) return { text: raw.replace(/\f/g, '').replace(/\s+$/, ''), failures: [] }
		return { text: raw.replace(/\s+/g, ' ').trim(), failures: [] }
	} catch (e) {
		return { text: '', failures: [`pdftotext wrote no text for ${path.basename(pdf)}: ${e}`] }
	}
}

// --- building the decks -----------------------------------------------------
/**
 * Write one case's deck and return its path.
 * @param {Case} testCase
 * @param {Buffer} fixture
 * @param {string} outDir
 * @returns {Promise<string>}
 */
async function buildDeck(testCase, fixture, outDir) {
	const { Presentation, isGraphicFrame } = await import(pathToFileURL(path.join(ROOT, 'dist', 'read.js')).href)
	const presentation = await Presentation.load(fixture)
	if (testCase.edit) {
		const frame = presentation.slides[DIAGRAM_PAGE - 1].shapes.find(isGraphicFrame)
		if (!frame?.diagram) throw new Error(`${FIXTURE} slide ${DIAGRAM_PAGE} no longer holds a diagram`)
		const point = frame.diagram.points.find((/** @type {any} */ candidate) => candidate.text === ANCHOR)
		if (!point) throw new Error(`the anchor node "${ANCHOR}" is no longer in the fixture`)
		testCase.edit(point)
	}
	const file = path.join(outDir, `${testCase.label}.pptx`)
	await fs.writeFile(file, Buffer.from(await presentation.save()))
	return file
}

/**
 * Write one side of a construct pair and return its path.
 *
 * Built from scratch rather than edited out of a fixture: these constructs are about what the
 * *emitter* writes, so the deck has to come through the public authoring API the same way a
 * caller's would. The canary box is laid down first and identically on both sides, which is
 * what makes "these two renders differ" a statement about the construct alone.
 * @param {Pair} pair
 * @param {'on' | 'off'} side
 * @param {string} outDir
 * @returns {Promise<string>}
 */
async function buildPairDeck(pair, side, outDir) {
	const { TsPptx } = await import(pathToFileURL(path.join(ROOT, 'dist', 'node.js')).href)
	const pptx = new TsPptx()
	const slide = pptx.addSlide()
	slide.addText(CANARY, { x: 0.6, y: 0.4, w: 8.5, h: 0.6, fontSize: 18, fontFace: 'Arial' })
	pair[side](slide)
	const file = path.join(outDir, `pair-${pair.label}-${side}.pptx`)
	await pptx.writeFile({ fileName: file })
	return file
}

/**
 * @param {Pair} pair
 * @param {string} onText
 * @param {string} offText
 * @returns {string[]} one message per failure
 */
function verifyPair(pair, onText, offText) {
	/** @type {string[]} */
	const failures = []
	/** @type {[string, string][]} */
	const sides = [
		['on', onText],
		['off', offText],
	]
	for (const [side, text] of sides) {
		for (const needle of pair.canaries) {
			if (!text.includes(needle)) failures.push(`[${pair.label}] the ${side} deck did not paint "${needle}"`)
		}
	}
	// Only meaningful once both sides are known to have rendered; a blank page trivially
	// "differs" from a populated one, and reporting that as a pass is the failure mode this
	// whole check exists to avoid.
	if (!failures.length) failures.push(...pair.verify(onText, offText).map((m) => `[${pair.label}] ${m}`))
	return failures
}

// --- verifying --------------------------------------------------------------
/**
 * @param {Case} testCase
 * @param {string} text the painted text, whitespace-normalized
 * @returns {string[]} one message per failure
 */
function verify(testCase, text) {
	/** @type {string[]} */
	const failures = []
	for (const needle of [...testCase.expect, ...NEIGHBOURS]) {
		if (!text.includes(needle)) failures.push(`[${testCase.label}] LibreOffice did not paint "${needle}"`)
	}
	for (const needle of testCase.reject) {
		if (text.includes(needle)) failures.push(`[${testCase.label}] LibreOffice painted "${needle}", which it must not`)
	}
	return failures
}

// --- orchestrate ------------------------------------------------------------
async function main() {
	// On a workstation a missing tool is a SKIP: this is a machine-dependent check and not
	// everyone has LibreOffice. In CI that would be the worst possible outcome — a lane that
	// installs the tools, fails to find them, and reports green while proving nothing. The
	// `required` mode makes their absence the failure instead, the same bargain
	// `FONT_ORACLES: required` strikes for the measurement oracles.
	const required = process.env.TSPPTX_RENDER_ORACLE === 'required'
	/** @param {string} message */
	const missing = (message) => {
		if (!required) {
			console.log('SKIP: ' + message)
			return
		}
		console.error('TSPPTX_RENDER_ORACLE=required, but ' + message)
		process.exitCode = 1
	}

	const soffice = await findSoffice()
	if (!soffice) {
		missing('LibreOffice not found. Set TSPPTX_SOFFICE to its soffice binary to run this check.')
		return
	}
	const pdftotext = await findPdfToText()
	if (!pdftotext) {
		missing('pdftotext not found (xpdf-tools or poppler-utils). Set TSPPTX_PDFTOTEXT to run this check.')
		return
	}
	console.log('LibreOffice: ' + soffice)
	console.log('pdftotext:   ' + pdftotext)

	const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'ts-pptx-lo-smoke-'))
	const profileDir = path.join(workDir, 'profile')
	const fixture = await fs.readFile(FIXTURE)

	/** @type {string[]} */
	const failures = []
	try {
		/** @type {{testCase: Case, deck: string}[]} */
		const built = []
		for (const testCase of CASES) built.push({ testCase, deck: await buildDeck(testCase, fixture, workDir) })

		/** @type {{pair: Pair, on: string, off: string}[]} */
		const builtPairs = []
		for (const pair of PAIRS) {
			builtPairs.push({
				pair,
				on: await buildPairDeck(pair, 'on', workDir),
				off: await buildPairDeck(pair, 'off', workDir),
			})
		}

		console.log(
			`Built ${built.length} decks from ${path.relative(ROOT, FIXTURE)} and ` +
				`${builtPairs.length * 2} construct decks; rendering...`
		)

		// One invocation for everything. Starting LibreOffice is most of the wall clock here,
		// so a second cold start per family would roughly double the run for no extra coverage.
		const decks = [...built.map((entry) => entry.deck), ...builtPairs.flatMap((entry) => [entry.on, entry.off])]
		failures.push(...(await convertToPdf(soffice, decks, workDir, profileDir)))

		if (!failures.length) {
			for (const { testCase, deck } of built) {
				const pdf = deck.replace(/\.pptx$/, '.pdf')
				const painted = await paintedText(pdftotext, pdf, testCase.page)
				failures.push(...painted.failures)
				if (painted.failures.length) continue
				const caseFailures = verify(testCase, painted.text)
				failures.push(...caseFailures)
				console.log(`[${testCase.label}] ${caseFailures.length ? 'FAIL' : 'ok'} — ${testCase.claim}`)
			}

			for (const { pair, on, off } of builtPairs) {
				const onPainted = await paintedText(pdftotext, on.replace(/\.pptx$/, '.pdf'), 1, pair.layout)
				const offPainted = await paintedText(pdftotext, off.replace(/\.pptx$/, '.pdf'), 1, pair.layout)
				const extractionFailures = [...onPainted.failures, ...offPainted.failures]
				failures.push(...extractionFailures)
				if (extractionFailures.length) continue
				const pairFailures = verifyPair(pair, onPainted.text, offPainted.text)
				failures.push(...pairFailures)
				console.log(`[${pair.label}] ${pairFailures.length ? 'FAIL' : 'ok'} — ${pair.claim}`)
			}
		}
	} finally {
		if (KEEP) console.log('Kept: ' + workDir)
		else await fs.rm(workDir, { recursive: true, force: true })
	}

	if (failures.length) {
		console.error('\nLibreOffice render smoke FAILED:')
		for (const f of failures) console.error('  - ' + f)
		process.exitCode = 1
		return
	}
	console.log('\nLibreOffice render smoke PASSED.')
}

main().catch((e) => {
	console.error('LibreOffice render smoke errored: ' + (e?.stack || e))
	process.exitCode = 1
})
