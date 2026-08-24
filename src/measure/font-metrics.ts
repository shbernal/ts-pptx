/**
 * Write-side font metrics.
 *
 * Provides per-glyph advance widths for measured text fit (see `text-fit.ts`).
 * The library has no write-side font registration otherwise; consumers register
 * a face's font file via `pptx.registerFontMetrics(...)` and the measured-fit
 * pass uses these advances to compute a baked `fontScale` for `fit: 'shrink'`.
 *
 * Width is summed from **raw `hmtx` advances with no kerning/shaping** (GPOS/GSUB
 * are deliberately ignored): kerning almost always narrows a line, so raw advances
 * over-estimate width — the conservative direction (shrink a touch too much, never
 * overflow). See `docs/measured-text-fit.md` ("Font metrics provider").
 */

import { UnsupportedFeatureError } from '../errors.js'
import type { FitParagraph, MetricsResolver } from './text-fit.js'

/** Per-(face,bold,italic) advance-width source backed by a parsed font file. */
export interface FontMetrics {
	/** Font design units per em (advances are in these units). */
	readonly unitsPerEm: number
	/**
	 * Advance width of `text` at `sizePt`, in points. Sums raw glyph advances
	 * (no kerning) and adds `charSpacingPt` after every character, matching how
	 * PowerPoint widens runs with character spacing.
	 */
	advanceWidthPt: (text: string, sizePt: number, charSpacingPt?: number) => number
	/**
	 * True if this face's cmap maps code point `cp` to a real glyph (not the
	 * `.notdef` fallback). Use it to flag source code points the face cannot
	 * render. Authoritative only for file-backed metrics; the unregistered-font
	 * heuristic has no cmap and reports every code point as covered.
	 */
	hasCodepoint: (cp: number) => boolean
}

/** Minimal slice of an opentype.js `Font` that `OpentypeFontMetrics` relies on. */
interface OpentypeFontLike {
	unitsPerEm: number
	charToGlyph: (ch: string) => { advanceWidth?: number } | undefined
	hasChar: (ch: string) => boolean
}

class OpentypeFontMetrics implements FontMetrics {
	private readonly font: OpentypeFontLike
	constructor(font: OpentypeFontLike) {
		this.font = font
	}

	get unitsPerEm(): number {
		return this.font.unitsPerEm
	}

	advanceWidthPt(text: string, sizePt: number, charSpacingPt = 0): number {
		if (!text) return 0
		// Sum raw per-glyph `hmtx` advances via cmap (charToGlyph) — deliberately
		// NOT getAdvanceWidth(), which runs GPOS/GSUB shaping (kerning narrows lines,
		// and some lookups even throw). Raw advances over-estimate width, the
		// conservative direction. `for..of` iterates code points (astral-safe).
		const scale = sizePt / this.font.unitsPerEm
		let width = 0
		let count = 0
		for (const ch of text) {
			const glyph = this.font.charToGlyph(ch)
			width += (glyph?.advanceWidth ?? 0) * scale
			count++
		}
		if (charSpacingPt) width += charSpacingPt * count
		return width
	}

	hasCodepoint(cp: number): boolean {
		// opentype's charToGlyph() falls back to .notdef for uncovered code points,
		// so it can't answer coverage; hasChar() keys on the cmap (glyph index > 0).
		return this.font.hasChar(String.fromCodePoint(cp))
	}
}

const HEURISTIC_UNITS_PER_EM = 1000

/**
 * Advance width of an ASCII/Unicode code point as a fraction of em, for the
 * unregistered-font fallback. Values are biased to **over-estimate** a typical
 * proportional Latin face: over-estimating width is the conservative direction for
 * both solvers (shrink a touch too much / grow a touch too tall, never overflow).
 * Non-Latin code points are treated as full-em (safe for CJK, which is otherwise
 * out of scope; slightly over-wide for accented Latin, which is harmless).
 */
function heuristicCharRatio(ch: string): number {
	const cp = ch.codePointAt(0) ?? 0
	if (cp > 0x2e7f) return 1.0 // CJK / full-width and beyond
	if (cp > 0xff) return 0.62 // Latin-Extended / accented / misc symbols
	if (ch === ' ' || ch === '\t') return 0.3
	if ("ijl.,'`!|:;".includes(ch)) return 0.3
	if ('ftr()[]{}/\\I'.includes(ch)) return 0.4
	if (ch === 'm' || ch === 'w') return 0.9
	if (ch === 'M' || ch === 'W') return 0.98
	if (ch === '@' || ch === '%') return 1.0
	if (ch >= 'A' && ch <= 'Z') return 0.72
	return 0.58 // lowercase, digits, default punctuation
}

class HeuristicFontMetrics implements FontMetrics {
	readonly unitsPerEm = HEURISTIC_UNITS_PER_EM
	advanceWidthPt(text: string, sizePt: number, charSpacingPt = 0): number {
		if (!text) return 0
		let width = 0
		let count = 0
		for (const ch of text) {
			width += heuristicCharRatio(ch) * sizePt
			count++
		}
		if (charSpacingPt) width += charSpacingPt * count
		return width
	}

	hasCodepoint(): boolean {
		return true // no cmap to consult; never fabricate a 'missing' glyph
	}
}

let heuristicSingleton: FontMetrics | undefined

/**
 * Shared, font-independent `FontMetrics` used when a named face has no registered
 * metrics but the deck has opted into measured fit (some other face is registered).
 * Approximate but conservative — see {@link heuristicCharRatio}. A deck with no
 * registered metrics at all does not engage measured fit, so this never fires there.
 */
export function getHeuristicFontMetrics(): FontMetrics {
	if (!heuristicSingleton) heuristicSingleton = new HeuristicFontMetrics()
	return heuristicSingleton
}

/** Parse a font file (TTF/OTF) into a `FontMetrics`. Lazily imports opentype.js. */
export async function parseFontMetrics(data: Uint8Array): Promise<FontMetrics> {
	// Dynamic import keeps opentype.js off the critical/browser path until a
	// consumer actually registers metrics.
	const mod = (await import('opentype.js')) as unknown as {
		parse?: (buf: ArrayBuffer) => OpentypeFontLike
		default?: { parse: (buf: ArrayBuffer) => OpentypeFontLike }
	}
	const parse = mod.parse ?? mod.default?.parse
	if (typeof parse !== 'function')
		throw new UnsupportedFeatureError('font/opentype-unavailable', 'opentype.js: parse() not found in module exports')
	// opentype.parse needs a standalone ArrayBuffer view of exactly the font bytes.
	const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
	return new OpentypeFontMetrics(parse(buf))
}

/**
 * Registry of `FontMetrics` keyed by `(face, bold, italic)`. Variant advances
 * differ, so bold/italic are stored separately; lookup falls back to the regular
 * variant (then any registered variant of the face) so a deck that only registers
 * the regular weight still measures bold runs approximately rather than not at all.
 */
export class FontMetricsRegistry {
	private readonly map = new Map<string, FontMetrics>()

	private key(face: string, bold: boolean, italic: boolean): string {
		return `${face.toLowerCase()}\0${bold ? 1 : 0}\0${italic ? 1 : 0}`
	}

	set(face: string, metrics: FontMetrics, opts?: { bold?: boolean; italic?: boolean }): void {
		this.map.set(this.key(face, !!opts?.bold, !!opts?.italic), metrics)
	}

	/** True if any variant of `face` has registered metrics. */
	hasFace(face: string): boolean {
		const prefix = `${face.toLowerCase()}\0`
		for (const k of this.map.keys()) if (k.startsWith(prefix)) return true
		return false
	}

	/**
	 * Coverage of code point `cp` for `face`, read from the registered font's cmap:
	 * `true`/`false` when a variant of `face` is registered (see {@link get} for the
	 * exact→regular→any-variant fallback), or `undefined` when the face has **no**
	 * registered metrics — "unknown", deliberately distinct from a `false`
	 * "registered but not covered". A coverage audit must not treat an unregistered
	 * face as fully covered, so callers get `undefined` here rather than the per-face
	 * {@link FontMetrics.hasCodepoint} boolean (which, for the heuristic fallback,
	 * always reports covered). Register the face via {@link parseFontMetrics} +
	 * {@link set} first.
	 */
	hasCodepoint(
		face: string | undefined | null,
		cp: number,
		opts?: { bold?: boolean; italic?: boolean }
	): boolean | undefined {
		return this.get(face, opts?.bold, opts?.italic)?.hasCodepoint(cp)
	}

	/** Resolve metrics for a run, falling back from exact variant → regular → any variant. */
	get(face: string | undefined | null, bold = false, italic = false): FontMetrics | undefined {
		// A run with no explicit face inherits the theme font, which we cannot resolve
		// reliably here — treat as unregistered so measured fit degrades gracefully.
		if (typeof face !== 'string' || face.length === 0) return undefined
		const exact = this.map.get(this.key(face, bold, italic))
		if (exact) return exact
		const regular = this.map.get(this.key(face, false, false))
		if (regular) return regular
		const prefix = `${face.toLowerCase()}\0`
		for (const [k, v] of this.map) if (k.startsWith(prefix)) return v
		return undefined
	}

	get size(): number {
		return this.map.size
	}
}

/**
 * A {@link MetricsResolver} backed by a registry: exact metrics for a registered face, otherwise
 * the shared heuristic fallback (with `onHeuristic` told which face it stood in for). A run with
 * no `fontFace` at all resolves to `undefined`, which the callers treat as unmeasurable.
 */
export function makeRegistryResolver(
	registry: FontMetricsRegistry,
	onHeuristic?: (face: string) => void
): MetricsResolver {
	return (run) => {
		const exact = registry.get(run.fontFace, run.bold, run.italic)
		if (exact) return exact
		if (typeof run.fontFace === 'string' && run.fontFace.length > 0) {
			onHeuristic?.(run.fontFace)
			return getHeuristicFontMetrics()
		}
		return undefined
	}
}

/**
 * Accumulate, per face, the code points `paragraphs` uses that the run's **registered**
 * metrics have no glyph for.
 *
 * This is the one thing the width arithmetic cannot see. PowerPoint substitutes another
 * face per code point and lays the run out in *that* font's advances; the model has no
 * fallback and charges the registered font's `.notdef` advance instead, which is a single
 * flat number unrelated to the glyph that actually paints. Unlike the other approximations
 * here it is **not** conservative in a fixed direction: a `.notdef` wider than the real
 * glyph over-reports (a phantom line — safe), a narrower one under-reports and can drop a
 * line, which is the overflow direction the resize bake has no safety net for. So it is
 * surfaced rather than absorbed — see `docs/measured-text-fit.md` ("No font fallback").
 *
 * Runs whose face has no registered metrics at all are skipped: they measure through the
 * cmap-less heuristic, which has no coverage to report, and the caller already flags them
 * as approximated faces.
 * @param paragraphs - the runs to audit, as handed to the layout
 * @param registry - the metrics the layout resolved against
 * @param into - face → uncovered code points, added to in place
 */
export function collectUncoveredCodepoints(
	paragraphs: FitParagraph[],
	registry: FontMetricsRegistry,
	into: Map<string, Set<number>>
): void {
	for (const para of paragraphs) {
		for (const run of para.runs) {
			const face = run.fontFace
			if (typeof face !== 'string' || face.length === 0) continue
			const metrics = registry.get(face, run.bold, run.italic)
			if (!metrics) continue
			let missing: Set<number> | undefined
			// `for..of` iterates code points, so an astral character is audited whole
			// rather than as two uncovered surrogates.
			for (const ch of run.text) {
				const cp = ch.codePointAt(0)
				if (cp === undefined || metrics.hasCodepoint(cp)) continue
				missing ??= into.get(face) ?? new Set<number>()
				missing.add(cp)
			}
			if (missing) into.set(face, missing)
		}
	}
}
