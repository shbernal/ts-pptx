// Generates the *.cases.json manifests for the autofit calibration decks.
// Output is committed next to each .pptx in test/read/fixtures/.
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'

const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..', 'test', 'read', 'fixtures')
const SLIDE_W = 960,
	SLIDE_H = 540
const DEFAULT_INSETS = { l: 7.2, r: 7.2, t: 3.6, b: 3.6 } // PowerPoint defaults (0.1"/0.05")
const FONTS = [
	{ face: 'Aptos', slug: 'aptos' },
	{ face: 'Aptos SemiBold', slug: 'aptossemibold' },
	{ face: 'Calibri', slug: 'calibri' },
	{ face: 'Tahoma', slug: 'tahoma' },
	{ face: 'Arial', slug: 'arial' },
]
const SIZES = [12, 18, 32]

// one shared metric word; line height is metric-driven (ascent+descent+gap), not glyph-driven
const METRIC_WORD = 'Hamburgefontsiv'
const LC = 'abcdefghijklmnopqrstuvwxyz'
const UC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
const DG = '0123456789'

const run = (text, font, sizePt, o = {}) => ({ text, font, sizePt, bold: !!o.bold, italic: !!o.italic, ...o })
const para = (runs, o = {}) => ({ runs: Array.isArray(runs) ? runs : [runs], ...o })

function deck1() {
	const cases = []
	let slide = 0
	for (const f of FONTS) {
		for (const sz of SIZES) {
			// --- line-height slide: 1/2/3 non-wrapping lines, single spacing ---
			slide++
			const ys = [30, 180, 340]
			for (let n = 1; n <= 3; n++) {
				const text = Array.from({ length: n }, () => METRIC_WORD).join('\n')
				cases.push({
					id: `linemetrics__${f.slug}__sz${sz}__lines${n}`,
					kind: 'resize',
					slide,
					xPt: 30,
					yPt: ys[n - 1],
					wPt: 400,
					hPt: 60,
					wrap: false,
					anchor: 't',
					insetsPt: DEFAULT_INSETS,
					paragraphs: [para(run(text, f.face, sz), { lineSpacingPct: 100, spaceBeforePts: 0, spaceAfterPts: 0 })],
				})
			}
			// --- advance-width slide: lowercase / uppercase / digits, AutoSizeNone fixed boxes ---
			slide++
			const strs = [
				['lc', LC],
				['uc', UC],
				['dg', DG],
			]
			const ays = [30, 150, 290]
			strs.forEach(([tag, str], i) => {
				cases.push({
					id: `advance__${f.slug}__sz${sz}__${tag}`,
					kind: 'fixed',
					slide,
					xPt: 30,
					yPt: ays[i],
					wPt: 900,
					hPt: 80,
					wrap: false,
					anchor: 't',
					insetsPt: DEFAULT_INSETS,
					paragraphs: [para(run(str, f.face, sz))],
				})
			})
		}
	}
	return {
		deck: 'autofit-line-metrics',
		slideWidthPt: SLIDE_W,
		slideHeightPt: SLIDE_H,
		fontsRequired: FONTS.map((f) => f.face),
		notes:
			'Deck 1: per-font line height (resize cy at 1/2/3 lines) + advance-width strings (AutoSizeNone). LibreOffice cross-measure applies here.',
		cases,
	}
}

// ---- Deck 2: shrink (normAutofit) calibration ----
const SENT = 'The quick brown fox jumps over the lazy dog. '
const overflow = (n) => SENT.repeat(n).trim()
const IN = 72 // pt per inch
const BOLDABLE = [
	{ face: 'Aptos', slug: 'aptos' },
	{ face: 'Calibri', slug: 'calibri' },
	{ face: 'Tahoma', slug: 'tahoma' },
	{ face: 'Arial', slug: 'arial' },
]

function deck2() {
	const cases = []
	let slide = 0
	const add = (id, o) => {
		slide++
		cases.push({
			id,
			kind: 'shrink',
			slide,
			xPt: 60,
			yPt: 160,
			wPt: 3.0 * IN,
			hPt: 1.0 * IN,
			wrap: true,
			anchor: 't',
			insetsPt: DEFAULT_INSETS,
			paragraphs: [para(run(overflow(3), 'Aptos', 18), { lineSpacingPct: 100, spaceBeforePts: 0, spaceAfterPts: 0 })],
			...o,
		})
	}
	const oneRun = (text, face, size, o = {}) => [
		para(run(text, face, size, o.run || {}), o.para || { lineSpacingPct: 100 }),
	]

	// --- per-font core: 5 regular + 4 bold, fixed 3x overflow scenario ---
	for (const f of FONTS) {
		add(`shrink__${f.slug}__sz18__b0i0__w3.00h1.00__core`, { paragraphs: oneRun(overflow(3), f.face, 18) })
	}
	for (const f of BOLDABLE) {
		add(`shrink__${f.slug}__sz18__b1i0__w3.00h1.00__core`, {
			paragraphs: oneRun(overflow(3), f.face, 18, { run: { bold: true } }),
		})
	}

	// --- policy: overflow-magnitude ladder (Aptos regular) ---
	for (const n of [2, 3, 4, 6, 10]) {
		add(`shrink__aptos__sz18__b0i0__ovr${n}x`, { paragraphs: oneRun(overflow(n), 'Aptos', 18) })
	}
	// --- policy: overflow ladder (Aptos bold) ---
	for (const n of [2, 4, 8]) {
		add(`shrink__aptos__sz18__b1i0__ovr${n}x`, {
			paragraphs: oneRun(overflow(n), 'Aptos', 18, { run: { bold: true } }),
		})
	}

	// --- policy: line-spacing variants (Aptos, 3x overflow) ---
	add('shrink__aptos__sz18__lnspc_single', {
		paragraphs: oneRun(overflow(3), 'Aptos', 18, { para: { lineSpacingPct: 100 } }),
	})
	add('shrink__aptos__sz18__lnspc_150', {
		paragraphs: oneRun(overflow(3), 'Aptos', 18, { para: { lineSpacingPct: 150 } }),
	})
	add('shrink__aptos__sz18__lnspc_exact18', {
		paragraphs: oneRun(overflow(3), 'Aptos', 18, { para: { lineSpacingPts: 18 } }),
	})

	// --- policy: space before/after, multi-paragraph ---
	add('shrink__aptos__sz18__spba10_2para', {
		paragraphs: [
			para(run(overflow(2), 'Aptos', 18), { lineSpacingPct: 100, spaceBeforePts: 10, spaceAfterPts: 10 }),
			para(run(overflow(2), 'Aptos', 18), { lineSpacingPct: 100, spaceBeforePts: 10, spaceAfterPts: 10 }),
		],
	})

	// --- policy: multi-run paragraph (mixed bold + mixed size on one wrapped line) ---
	add('shrink__aptos__sz18__multirun', {
		paragraphs: [
			para(
				[
					run('Big bold start ', 'Aptos', 28, { bold: true }),
					run('then normal eighteen point text that continues ', 'Aptos', 18),
					run('and small twelve tail that keeps going to overflow the box for sure. ', 'Aptos', 12),
					run(overflow(2), 'Aptos', 18),
				],
				{ lineSpacingPct: 100 }
			),
		],
	})

	// --- policy: non-default insets ---
	add('shrink__aptos__sz18__insets', {
		insetsPt: { l: 20, r: 5, t: 15, b: 2 },
		paragraphs: oneRun(overflow(3), 'Aptos', 18),
	})

	// --- policy: non-default charSpacing (tracking) ---
	add('shrink__aptos__sz18__charspc2', { paragraphs: oneRun(overflow(3), 'Aptos', 18, { run: { charSpacingPts: 2 } }) })

	// --- policy: vertical anchor (should not change fontScale) ---
	for (const a of ['t', 'ctr', 'b']) {
		add(`shrink__aptos__sz18__anchor_${a}`, { anchor: a, paragraphs: oneRun(overflow(3), 'Aptos', 18) })
	}

	// --- policy: width-driven (short box, one long line) vs height-driven (tall-ish, many lines) ---
	add('shrink__aptos__sz18__widthdriven', { hPt: 0.5 * IN, paragraphs: oneRun(overflow(1), 'Aptos', 18) })
	add('shrink__aptos__sz18__heightdriven', { hPt: 1.0 * IN, paragraphs: oneRun(overflow(4), 'Aptos', 18) })

	return {
		deck: 'autofit-shrink',
		slideWidthPt: SLIDE_W,
		slideHeightPt: SLIDE_H,
		fontsRequired: FONTS.map((f) => f.face),
		notes:
			'Deck 2: normAutofit (shrink) calibration. Per-font core (5 regular + 4 bold, fixed 3x overflow in a 3.0x1.0in box at 18pt) + Aptos policy sweep (overflow ladder, line-spacing, space before/after, multi-run, insets, charSpacing, vertical anchor, width- vs height-driven).',
		cases,
	}
}

// ---- Deck 3: resize (spAutoFit / baked cy) calibration ----
function deck3() {
	const cases = []
	let slide = 0
	const three = [METRIC_WORD, METRIC_WORD, METRIC_WORD].join('\n')
	const mk = (id, o) => {
		slide++
		cases.push({
			id,
			kind: 'resize',
			slide,
			xPt: 60,
			yPt: 200,
			wPt: 3.0 * IN,
			hPt: 0.5 * IN,
			wrap: true,
			anchor: 't',
			insetsPt: DEFAULT_INSETS,
			paragraphs: [para(run(three, 'Aptos', 18), { lineSpacingPct: 100 })],
			...o,
		})
	}
	const oneRun = (text, face, size, o = {}) => [
		para(run(text, face, size, o.run || {}), o.para || { lineSpacingPct: 100 }),
	]

	// --- per-font core: single line + 3 hard lines (corroborates Deck 1 line height via cy) ---
	for (const f of FONTS) {
		mk(`resize__${f.slug}__lines1`, { paragraphs: oneRun(METRIC_WORD, f.face, 18) })
		mk(`resize__${f.slug}__lines3`, { paragraphs: oneRun(three, f.face, 18) })
	}

	// --- policy: vertical anchor → how off.y moves as the box grows ---
	for (const a of ['t', 'ctr', 'b']) {
		mk(`resize__aptos__anchor_${a}`, { anchor: a, paragraphs: oneRun(three, 'Aptos', 18) })
	}
	// --- policy: under-filled box (tall authored height, one short line) → does spAutoFit shrink cy? ---
	mk('resize__aptos__underfilled', { hPt: 2.0 * IN, paragraphs: oneRun(METRIC_WORD, 'Aptos', 18) })
	// --- policy: line-spacing contribution to cy ---
	mk('resize__aptos__lnspc150', { paragraphs: oneRun(three, 'Aptos', 18, { para: { lineSpacingPct: 150 } }) })
	mk('resize__aptos__lnspc_exact24', { paragraphs: oneRun(three, 'Aptos', 18, { para: { lineSpacingPts: 24 } }) })
	// --- policy: space before/after contribution (multi-paragraph) ---
	mk('resize__aptos__spba12_2para', {
		paragraphs: [
			para(run(METRIC_WORD, 'Aptos', 18), { lineSpacingPct: 100, spaceBeforePts: 12, spaceAfterPts: 12 }),
			para(run(METRIC_WORD, 'Aptos', 18), { lineSpacingPct: 100, spaceBeforePts: 12, spaceAfterPts: 12 }),
		],
	})
	// --- policy: insets contribution to cy ---
	mk('resize__aptos__insets', { insetsPt: { l: 7.2, r: 7.2, t: 20, b: 20 }, paragraphs: oneRun(three, 'Aptos', 18) })
	// --- policy: genuine soft-wrap (honours the wrap simulator path) ---
	mk('resize__aptos__softwrap', { paragraphs: oneRun(overflow(1), 'Aptos', 18) })

	return {
		deck: 'autofit-resize',
		slideWidthPt: SLIDE_W,
		slideHeightPt: SLIDE_H,
		fontsRequired: FONTS.map((f) => f.face),
		notes:
			'Deck 3: spAutoFit (resize) cy calibration. Per-font core (single + 3 hard lines, corroborates Deck 1) + Aptos policy sweep (vertical anchor off.y on growth, under-filled box, line-spacing, space before/after, insets, soft-wrap). LibreOffice cross-measure applies (resize cy vs LO height).',
		cases,
	}
}

// ---- Deck 4: edge cases (Aptos) most likely to break the simulator ----
function deck4() {
	const cases = []
	let slide = 0
	const LONGTOKEN = 'Supercalifragilisticexpialidociousantidisestablishmentarianism'
	const push = (id, o) => {
		slide++
		cases.push({
			id,
			kind: 'resize',
			slide,
			xPt: 60,
			yPt: 200,
			wPt: 2.0 * IN,
			hPt: 0.5 * IN,
			wrap: true,
			anchor: 't',
			insetsPt: DEFAULT_INSETS,
			paragraphs: [para(run('', 'Aptos', 18), { lineSpacingPct: 100 })],
			...o,
		})
	}
	const oneRun = (text, o = {}) => [para(run(text, 'Aptos', 18, o.run || {}), o.para || { lineSpacingPct: 100 })]

	// over-long single token, no break opportunity -> character wrap vs width overflow
	push('edge__longtoken_charwrap', {
		wrap: true,
		note: 'Over-long unbreakable token in a fixed-width wrap box; pins PowerPoint character-wrap behaviour.',
		paragraphs: oneRun(LONGTOKEN),
	})
	push('edge__longtoken_nowrap', {
		wrap: false,
		note: 'Same token, wrap off; box grows in width instead of char-wrapping (comparison).',
		paragraphs: oneRun(LONGTOKEN),
	})

	// trailing spaces: do they count toward line width? (wrap off so width grows to content)
	push('edge__trailing_spaces', {
		wrap: false,
		note: 'Trailing spaces before line end; compare extCx to the no-trailing-space control.',
		paragraphs: oneRun('Hamburgefontsiv          '),
	})
	push('edge__no_trailing_spaces', {
		wrap: false,
		note: 'Control for trailing-spaces width.',
		paragraphs: oneRun('Hamburgefontsiv'),
	})

	// leading whitespace
	push('edge__leading_spaces', {
		wrap: false,
		note: 'Leading spaces; do they widen the box?',
		paragraphs: oneRun('          Hamburgefontsiv'),
	})

	// empty paragraph / blank line height contribution
	push('edge__blank_line', {
		note: 'Blank middle paragraph; pins empty-line height contribution to cy.',
		paragraphs: [
			para(run('Line A', 'Aptos', 18), { lineSpacingPct: 100 }),
			para(run('', 'Aptos', 18), { lineSpacingPct: 100 }),
			para(run('Line B', 'Aptos', 18), { lineSpacingPct: 100 }),
		],
	})

	// run that is only whitespace
	push('edge__only_whitespace', {
		note: 'A run that is only whitespace; does the line collapse?',
		paragraphs: oneRun('          '),
	})

	// tab character handling
	push('edge__tab', {
		wrap: false,
		note: 'Tab character between tokens; pins tab advance handling.',
		paragraphs: oneRun('A\tB'),
	})

	// mixed font sizes in one paragraph -> line height = max run height
	push('edge__mixed_sizes', {
		note: 'Mixed run sizes on one line; line height should follow the max run.',
		paragraphs: [
			para([run('small ', 'Aptos', 12), run('BIG', 'Aptos', 36), run(' small', 'Aptos', 12)], { lineSpacingPct: 100 }),
		],
	})

	// documentation-only unsupported scripts (kept as fixed boxes; recorded as known gaps)
	push('edge__cjk_unsupported', {
		kind: 'fixed',
		wrap: false,
		note: 'UNSUPPORTED (documentation only): CJK text. Aptos lacks CJK glyphs; recorded as a known measurement gap, not a calibration target.',
		paragraphs: oneRun('日本語テキスト'),
	})
	push('edge__rtl_unsupported', {
		kind: 'fixed',
		wrap: false,
		note: 'UNSUPPORTED (documentation only): RTL (Arabic) text. Bidi/shaping out of scope; recorded as a known gap.',
		paragraphs: oneRun('مرحبا بالعالم'),
	})

	return {
		deck: 'autofit-edge',
		slideWidthPt: SLIDE_W,
		slideHeightPt: SLIDE_H,
		fontsRequired: ['Aptos'],
		notes:
			'Deck 4: Aptos edge cases most likely to break a wrap/line simulator - over-long tokens, trailing/leading whitespace, blank line, whitespace-only run, tab, mixed sizes, plus documentation-only CJK + RTL boxes marked unsupported.',
		cases,
	}
}

const decks = { 'autofit-line-metrics': deck1, 'autofit-shrink': deck2, 'autofit-resize': deck3, 'autofit-edge': deck4 }
const which = process.argv[2]
const names = which ? [which] : Object.keys(decks)
for (const name of names) {
	const spec = decks[name]()
	const out = resolve(FIX, `${name}.cases.json`)
	writeFileSync(out, JSON.stringify(spec, null, 2) + '\n')
	console.log(`wrote ${out} (${spec.cases.length} cases)`)
}
