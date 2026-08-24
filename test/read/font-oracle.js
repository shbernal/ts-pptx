// Font resolution for the two PowerPoint-authored measurement oracles
// (`autofit-calibration-oracle.test.js`, `cjk-line-breaking-oracle.test.js`).
//
// Those oracles compare this repo's measured-fit model against what desktop PowerPoint
// baked for the same box, which only means anything if the model measures with the same
// advances PowerPoint did. That needs the genuine faces: Aptos, Aptos SemiBold, Arial,
// Calibri, Tahoma (autofit) and Malgun Gothic (CJK). None of them can be committed, so
// there are two sources here and a policy for choosing between them.
//
// **Genuine font files**, when the machine has them. On Windows that is the font registry
// (`resolveGenuineFontFile` below) rather than a directory scan, because that is the map
// GDI itself resolves a family name through and it is the only one that spans both font
// directories: Office installs Aptos per-user, under `%LOCALAPPDATA%`, with no elevation.
// Elsewhere it is `fc-match`, with a substitution guard — a metric-compatible clone
// (Carlito for Calibri) would answer the query and quietly recalibrate the oracle against
// a font PowerPoint never used.
//
// **The committed metrics sidecar** (`fixtures/autofit-font-metrics.json`) otherwise. It
// records the raw `hmtx` advance of every code point the committed cases actually use,
// per face, straight out of the genuine font, so the model charges the same width it
// would have charged with the file present. That is what lets the whole oracle run on a
// hosted runner: no runner will ever have Aptos, which ships with Microsoft 365 and not
// with Windows, and 35 of the 47 asserted autofit cases are Aptos.
//
// The sidecar is derived data and is treated as such: `font-metrics-sidecar.test.js`
// re-derives it from the genuine fonts wherever they resolve and fails on any drift, so
// a hand-edited or stale sidecar stops matching its source. `authoring/build-font-metrics.mjs`
// writes it.
//
// Two environment knobs, both fail-closed, because the failure mode this whole file
// exists to prevent is an oracle that resolves nothing and reports green:
//
//   FONT_ORACLES=required            every face a case needs must resolve from SOME
//                                    source. A skip becomes a failure. Set on every CI
//                                    leg that runs these suites.
//   FONT_ORACLES_GENUINE=A,B,C       these families must resolve to a real installed
//                                    font file. Declares what a runner image is expected
//                                    to provide, so an image that silently drops a font
//                                    fails the leg instead of falling back to the sidecar.
//                                    Enforced by `scripts/font-oracle-probe.mjs` and by
//                                    the sidecar test.
//
// Unset, both default to today's behaviour: use what is there, skip what is not. That is
// the right default for a workstation and for a fork with no fonts installed.

import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { parseFontMetrics } from '../../dist/measure.js'
import { FIXTURES } from './corpus.js'

/** The committed metrics sidecar. */
export const SIDECAR_PATH = path.join(FIXTURES, 'autofit-font-metrics.json')

/** Every face a case needs must resolve from some source, or the suite fails. */
export const REQUIRED = process.env.FONT_ORACLES === 'required'

/**
 * Measure from the committed sidecar even where the genuine font is installed.
 *
 * The point is reproducibility in both directions: it is how a workstation with all six
 * faces exercises the path every Linux runner takes, and how a CI failure that only
 * happens on the sidecar path can be reproduced on a machine that has the fonts. It does
 * NOT affect `resolveGenuineFontFile`, so `font-metrics-sidecar.test.js` still verifies
 * the sidecar against the real files while this is set.
 */
export const SIDECAR_ONLY = process.env.FONT_ORACLES_SIDECAR_ONLY === '1'

/** Families this runner declares it has installed for real. */
export const GENUINE_REQUIRED = (process.env.FONT_ORACLES_GENUINE ?? '')
	.split(',')
	.map((s) => s.trim())
	.filter(Boolean)

/**
 * A face, as both oracles and the sidecar name one.
 *
 * @typedef {{ family: string, bold?: boolean, italic?: boolean }} Face
 */

/** @param {Face} face */
export function faceLabel(face) {
	const style = [face.bold ? 'Bold' : '', face.italic ? 'Italic' : ''].filter(Boolean).join(' ')
	return style ? `${face.family} ${style}` : face.family
}

// ---------------------------------------------------------------------------
// Genuine font files
// ---------------------------------------------------------------------------

/** @type {Map<string, string> | null} */
let winFontIndex = null

/**
 * Display name (`arial bold`) to file path, read from the Windows font registry.
 *
 * Both hives: `HKLM` holds the machine fonts as bare filenames under `%SystemRoot%\Fonts`,
 * `HKCU` holds per-user installs as absolute paths, which is where Office puts Aptos.
 * A value can name several files separated by `|`; the first is the one the display name
 * refers to.
 */
function windowsFontIndex() {
	if (winFontIndex) return winFontIndex
	const index = new Map()
	const systemFonts = path.join(process.env.SystemRoot ?? 'C:\\Windows', 'Fonts')
	for (const hive of ['HKLM', 'HKCU']) {
		let out
		try {
			out = execFileSync('reg', ['query', `${hive}\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts`], {
				encoding: 'utf8',
				stdio: ['ignore', 'pipe', 'ignore'],
			})
		} catch {
			continue // no per-user fonts key, or no registry access; the other hive still counts
		}
		for (const line of out.split(/\r?\n/)) {
			const m = /^\s{4}(\S.*?)\s{4}REG_SZ\s{4}(.+?)\s*$/.exec(line)
			if (!m?.[1] || !m[2]) continue
			// Strip the format suffix the registry appends: `Arial Bold (TrueType)`.
			const name = m[1].replace(/\s*\((TrueType|OpenType)\)\s*$/i, '').toLowerCase()
			// A value can name several files separated by `|`; the first is this entry's.
			const file = (m[2].split('|')[0] ?? '').trim()
			if (!name || !file) continue
			// A per-user entry must win over a machine one of the same name: it is what the
			// user's own PowerPoint resolved when it authored the fixture.
			index.set(name, path.isAbsolute(file) ? file : path.join(systemFonts, file))
		}
	}
	winFontIndex = index
	return index
}

/**
 * Resolve `face` to a genuine installed font file, or null when it is absent or would be
 * substituted.
 *
 * @param {Face} face
 * @returns {string | null}
 */
export function resolveGenuineFontFile(face) {
	if (process.platform === 'win32') return windowsFontIndex().get(faceLabel(face).toLowerCase()) ?? null
	try {
		const styleBits = [face.bold ? 'bold' : '', face.italic ? 'italic' : ''].filter(Boolean).join(' ')
		const pattern = styleBits ? `${face.family}:style=${styleBits}` : face.family
		const out = execFileSync('fc-match', ['-f', '%{family}\t%{file}', pattern], {
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'ignore'],
		})
		const [fam, file] = out.split('\t')
		if (!fam || !file) return null
		// Reject substitution: the resolved family must contain the requested name.
		if (!fam.toLowerCase().includes(face.family.toLowerCase())) return null
		return file.trim()
	} catch {
		return null
	}
}

// ---------------------------------------------------------------------------
// The metrics sidecar
// ---------------------------------------------------------------------------

/** @type {{ schema: string, faces: Array<{ family: string, bold: boolean, italic: boolean, unitsPerEm: number, advances: Record<string, number>, uncovered: number[] }> } | null} */
let sidecarDoc = null

/** The committed sidecar, parsed once. */
export function readSidecar() {
	sidecarDoc ??= JSON.parse(readFileSync(SIDECAR_PATH, 'utf8'))
	return /** @type {NonNullable<typeof sidecarDoc>} */ (sidecarDoc)
}

/** @param {Face} face */
function sidecarEntry(face) {
	return (
		readSidecar().faces.find(
			(f) =>
				f.family.toLowerCase() === face.family.toLowerCase() && !!f.bold === !!face.bold && !!f.italic === !!face.italic
		) ?? null
	)
}

/**
 * A `FontMetrics` backed by recorded advances rather than by a font file.
 *
 * Deliberately throws on a code point the sidecar does not carry instead of charging a
 * default: an unrecorded code point means a case was added or edited without the sidecar
 * being regenerated, and the whole point of this file is that such a gap cannot pass.
 *
 * @param {{ family: string, bold: boolean, italic: boolean, unitsPerEm: number, advances: Record<string, number>, uncovered: number[] }} entry
 */
function sidecarFontMetrics(entry) {
	const advances = new Map(Object.entries(entry.advances).map(([cp, adv]) => [Number(cp), adv]))
	const uncovered = new Set(entry.uncovered ?? [])
	const label = faceLabel(entry)
	/** @param {number} cp */
	const advanceOf = (cp) => {
		const adv = advances.get(cp)
		if (adv === undefined) {
			const hex = cp.toString(16).toUpperCase().padStart(4, '0')
			throw new Error(
				`font metrics sidecar has no advance for ${label} U+${hex}. ` +
					`Regenerate it on a machine with the genuine fonts: pnpm run font-metrics:build`
			)
		}
		return adv
	}
	return {
		unitsPerEm: entry.unitsPerEm,
		/**
		 * @param {string} text
		 * @param {number} sizePt
		 * @param {number} [charSpacingPt]
		 */
		advanceWidthPt(text, sizePt, charSpacingPt = 0) {
			if (!text) return 0
			// Same arithmetic as OpentypeFontMetrics: raw advances, no shaping, char spacing
			// added per code point. See src/measure/font-metrics.ts.
			const scale = sizePt / entry.unitsPerEm
			let width = 0
			let count = 0
			for (const ch of text) {
				width += advanceOf(/** @type {number} */ (ch.codePointAt(0))) * scale
				count++
			}
			if (charSpacingPt) width += charSpacingPt * count
			return width
		},
		/** @param {number} cp */
		hasCodepoint(cp) {
			advanceOf(cp) // same staleness guard: an unknown code point is not an uncovered one
			return !uncovered.has(cp)
		},
	}
}

// ---------------------------------------------------------------------------
// The policy
// ---------------------------------------------------------------------------

/** What each requested face resolved to, for the end-of-suite accounting. */
/** @type {Map<string, 'genuine' | 'sidecar' | 'missing'>} */
const sources = new Map()

/** @type {Map<string, unknown>} */
const cache = new Map()

/**
 * Metrics for `face`: the genuine font where this machine has it, the committed sidecar
 * otherwise, and null when neither can answer (which throws under `FONT_ORACLES=required`).
 *
 * @param {Face} face
 * @returns {Promise<import('../../dist/measure.js').FontMetrics | null>}
 */
export async function oracleMetrics(face) {
	const key = faceLabel(face).toLowerCase()
	if (!cache.has(key)) {
		const file = SIDECAR_ONLY ? null : resolveGenuineFontFile(face)
		if (file) {
			cache.set(key, await parseFontMetrics(new Uint8Array(readFileSync(file))))
			sources.set(key, 'genuine')
		} else {
			const entry = sidecarEntry(face)
			cache.set(key, entry ? sidecarFontMetrics(entry) : null)
			sources.set(key, entry ? 'sidecar' : 'missing')
		}
	}
	const metrics = cache.get(key)
	if (!metrics && REQUIRED) {
		throw new Error(
			`FONT_ORACLES=required, but ${faceLabel(face)} resolved neither an installed font ` +
				`nor a sidecar entry. Install the face, or regenerate the sidecar on a machine ` +
				`that has it: pnpm run font-metrics:build`
		)
	}
	return /** @type {import('../../dist/measure.js').FontMetrics | null} */ (metrics ?? null)
}

/** How every face requested so far resolved. Read by the suites' accounting tests. */
export function resolutionTally() {
	let genuine = 0
	let sidecar = 0
	const missing = []
	for (const [key, source] of sources) {
		if (source === 'genuine') genuine++
		else if (source === 'sidecar') sidecar++
		else missing.push(key)
	}
	return { genuine, sidecar, missing, total: sources.size }
}

// ---------------------------------------------------------------------------
// What the committed cases need, and how a face is derived from a font
// ---------------------------------------------------------------------------

/**
 * The fixture sidecars whose cases the metrics sidecar has to cover. Enumerated from the
 * committed cases rather than hand-listed so a new case cannot quietly go unmeasured: it
 * either shows up here and gets recorded, or the sidecar-backed metrics throw on its first
 * unknown code point.
 */
const CASE_FILES = ['autofit-shrink.cases.json', 'autofit-resize.cases.json']
const CJK_ORACLE = 'autofit-cjk-wrap.oracle.json'

/**
 * Every (face, code point) pair the committed cases measure, keyed by `faceLabel`.
 *
 * @returns {Array<{ family: string, bold: boolean, italic: boolean, codepoints: number[] }>}
 */
export function neededFaces() {
	/** @type {Map<string, { family: string, bold: boolean, italic: boolean, codepoints: Set<number> }>} */
	const faces = new Map()
	/**
	 * @param {string} family
	 * @param {boolean} bold
	 * @param {boolean} italic
	 * @param {string} text
	 */
	const add = (family, bold, italic, text) => {
		const key = faceLabel({ family, bold, italic }).toLowerCase()
		let face = faces.get(key)
		if (!face) faces.set(key, (face = { family, bold, italic, codepoints: new Set() }))
		for (const ch of text) face.codepoints.add(/** @type {number} */ (ch.codePointAt(0)))
	}

	for (const file of CASE_FILES) {
		const spec = JSON.parse(readFileSync(path.join(FIXTURES, file), 'utf8'))
		for (const c of spec.cases) {
			for (const para of c.paragraphs) {
				for (const run of para.runs) add(run.font, !!run.bold, !!run.italic, run.text)
			}
		}
	}

	const cjk = JSON.parse(readFileSync(path.join(FIXTURES, CJK_ORACLE), 'utf8'))
	for (const c of cjk.cases) add(c.fontFace ?? cjk.fontFace, !!c.bold, !!c.italic, c.text)

	return [...faces.values()]
		.map((f) => ({ ...f, codepoints: [...f.codepoints].sort((a, b) => a - b) }))
		.sort((a, b) => faceLabel(a).localeCompare(faceLabel(b)))
}

/**
 * Record `codepoints` of `metrics` as a sidecar face entry.
 *
 * Advances are stored in font design units, which is what `hmtx` holds and what makes the
 * entry independent of any point size: asking the metrics for the width of one character
 * at `sizePt = unitsPerEm` scales by exactly 1, so this reads the same integer the font
 * carries, through the same code path the model charges widths with.
 *
 * `uncovered` is the cmap answer, kept separately because it is not derivable from the
 * advance: a code point the face lacks still has the `.notdef` advance charged against it
 * (see `OpentypeFontMetrics.advanceWidthPt`), and the CJK oracle needs to tell the two
 * apart to skip the cases PowerPoint resolved by falling back to another face.
 *
 * @param {{ family: string, bold: boolean, italic: boolean, codepoints: number[] }} face
 * @param {import('../../dist/measure.js').FontMetrics} metrics
 */
export function deriveFace(face, metrics) {
	/** @type {Record<string, number>} */
	const advances = {}
	/** @type {number[]} */
	const uncovered = []
	for (const cp of face.codepoints) {
		const units = metrics.advanceWidthPt(String.fromCodePoint(cp), metrics.unitsPerEm)
		const rounded = Math.round(units)
		if (Math.abs(units - rounded) > 1e-6) {
			throw new Error(`${faceLabel(face)} U+${cp.toString(16).toUpperCase()} has a non-integer advance (${units})`)
		}
		advances[String(cp)] = rounded
		if (!metrics.hasCodepoint(cp)) uncovered.push(cp)
	}
	return {
		family: face.family,
		bold: face.bold,
		italic: face.italic,
		unitsPerEm: metrics.unitsPerEm,
		advances,
		uncovered,
	}
}

/**
 * Differences between a recorded entry and one freshly derived from the genuine font, as
 * human-readable lines. Empty means the sidecar still describes the font it came from.
 *
 * @param {ReturnType<typeof deriveFace>} recorded
 * @param {ReturnType<typeof deriveFace>} derived
 */
export function diffFace(recorded, derived) {
	/** @type {string[]} */
	const diffs = []
	if (recorded.unitsPerEm !== derived.unitsPerEm) {
		diffs.push(`unitsPerEm: recorded ${recorded.unitsPerEm}, font says ${derived.unitsPerEm}`)
	}
	for (const [cp, advance] of Object.entries(derived.advances)) {
		const was = recorded.advances[cp]
		const hex = Number(cp).toString(16).toUpperCase().padStart(4, '0')
		if (was === undefined) diffs.push(`U+${hex}: missing from the sidecar (font says ${advance})`)
		else if (was !== advance) diffs.push(`U+${hex}: recorded ${was}, font says ${advance}`)
	}
	for (const cp of Object.keys(recorded.advances)) {
		// An entry no case reaches any more. Harmless to measure with, but it means the
		// sidecar was not regenerated when the cases changed, and the next reader cannot
		// tell that from a recording that is still current.
		if (derived.advances[cp] === undefined) {
			diffs.push(`U+${Number(cp).toString(16).toUpperCase().padStart(4, '0')}: recorded, but no case measures it`)
		}
	}
	const recordedUncovered = new Set(recorded.uncovered ?? [])
	const derivedUncovered = new Set(derived.uncovered)
	for (const cp of derivedUncovered) {
		if (!recordedUncovered.has(cp))
			diffs.push(`U+${cp.toString(16).toUpperCase()}: the font lacks it, the sidecar says it has it`)
	}
	for (const cp of recordedUncovered) {
		if (!derivedUncovered.has(cp))
			diffs.push(`U+${cp.toString(16).toUpperCase()}: the font has it, the sidecar says it lacks it`)
	}
	return diffs
}

/**
 * Parse the genuine font for `face`, or null when this machine does not have it.
 *
 * @param {Face} face
 */
export async function genuineMetrics(face) {
	const file = resolveGenuineFontFile(face)
	return file ? await parseFontMetrics(new Uint8Array(readFileSync(file))) : null
}
