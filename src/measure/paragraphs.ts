/**
 * ts-pptx: turning a slide text object into the simulator's inputs
 *
 * The lower half of the measured-fit pass: normalize a text object's runs into `FitParagraph`s,
 * resolve its body insets and box, and work out how a height change is shared between `off.y` and
 * `ext.cy` for a given anchor. Everything here is about *reading* an authored object; the solving
 * and the rewriting live in the siblings (`table-fit.ts`, `fit.ts`).
 */

import { DEF_FONT_SIZE } from '../constants-internal.js'
import { EMU_PER_POINT } from '../units.js'
import { getSmartParseNumber, resolveInsetsEmu as resolveMarginInsetsEmu } from '../units-internal.js'
import { BODY_INSET_DEFAULTS_EMU } from '../ooxml/body-insets.js'
import type { FitBox, FitParagraph, FitRun } from './text-fit.js'
import { TextAnchor } from '../enums.js'
import type { TextProps, TextPropsOptions } from '../types/index.js'
import type { ObjectOptionsInternal, SlideObject, PresSlideInternal } from '../types/internal.js'
import { pickDefined } from '../options-internal.js'

const CRLF_RE = /\r*\n/g

export type RunOpts = TextPropsOptions & ObjectOptionsInternal

/**
 * The two fields a run list is read from, and the whole contract {@link normalizeRuns} and
 * {@link extractParagraphs} have.
 *
 * Both used to take a full `SlideObject`, which is why both of their table call sites had to
 * write `{ text: cell.text, options: eff } as unknown as SlideObject` — a cell's `text` is
 * `string | TableCellInternal[]` and a slide object's is `TextProps[]`, so nothing weaker than a
 * double cast would fit. Naming the two fields these functions actually read makes a table
 * cell an ordinary argument, and a `SlideObject` still one too.
 *
 * `text` is deliberately loose: `normalizeRuns` accepts a bare string, a number, one
 * `TextProps`, or an array of them, and branches on the shape at runtime.
 */
interface RunSource {
	text?: unknown
	options?: RunOpts
}

/** Normalize `slideObj.text` (string | TextProps | TextProps[]) to a run list. */
export function normalizeRuns(obj: RunSource): TextProps[] {
	const opts = obj.options ?? {}
	const text = obj.text
	if (text == null) return []
	if (typeof text === 'string' || typeof text === 'number') return [{ text: String(text), options: opts }]
	// `pickDefined` for `text`: a run that carries none keeps the key off, which is what the
	// `TextProps` the generator builds do — this list is handed straight to `extractParagraphs`
	// and, from the table path, compared against runs that came from there.
	if (!Array.isArray(text) && typeof text === 'object' && 'text' in text) {
		const t = text as TextProps
		return [{ ...pickDefined(t, ['text']), options: t.options ?? opts }]
	}
	if (Array.isArray(text))
		return (text as TextProps[]).map((t) => ({ ...pickDefined(t, ['text']), options: t.options ?? opts }))
	return []
}

/** Build a measurable `FitParagraph[]` from a text object, or null if not measurable. */
export function extractParagraphs(obj: RunSource): FitParagraph[] | null {
	const opts = obj.options ?? {}
	const runs = normalizeRuns(obj)
	if (runs.length === 0) return null
	return buildFitParagraphs(runs, opts)
}

/**
 * Convert a run list (+ box-level default options) into a measurable
 * `FitParagraph[]`, or null if empty. The single converter shared by the
 * export-time pass ({@link extractParagraphs}) and the public layout-time
 * `measureText` API, so a layout-time prediction and the baked export never drift.
 */
export function buildFitParagraphs(runs: TextProps[], opts: RunOpts): FitParagraph[] | null {
	if (runs.length === 0) return null

	// Expand "\n" inside a run into separate pieces, flagging the paragraph break
	// after each (mirrors `gen/slide/objects/table.ts` STEP 4). `breakLine` ends a paragraph too.
	interface Piece {
		text: string
		options: RunOpts
		breakAfter: boolean
	}
	const pieces: Piece[] = []
	for (const run of runs) {
		const ro = (run.options ?? opts) as RunOpts
		const raw = String(run.text ?? '').replace(CRLF_RE, '\n')
		if (raw.includes('\n')) {
			const lines = raw.split('\n')
			lines.forEach((line, i) => {
				const isLast = i === lines.length - 1
				pieces.push({ text: line, options: ro, breakAfter: isLast ? !!ro.breakLine : true })
			})
		} else {
			pieces.push({ text: raw, options: ro, breakAfter: !!ro.breakLine })
		}
	}

	const toRun = (p: Piece): FitRun => {
		const ro = p.options
		const sizePt = Number(ro.fontSize ?? opts.fontSize ?? DEF_FONT_SIZE)
		return {
			text: p.text,
			sizePt: Number.isFinite(sizePt) && sizePt > 0 ? sizePt : DEF_FONT_SIZE,
			bold: !!(ro.bold ?? opts.bold),
			italic: !!(ro.italic ?? opts.italic),
			fontFace: ro.fontFace ?? opts.fontFace,
			charSpacingPt: (ro.charSpacing ?? opts.charSpacing) || undefined,
		}
	}
	const toPara = (runsForPara: FitRun[], paraOpts: RunOpts): FitParagraph => {
		const lineSpacing = paraOpts.lineSpacing ?? opts.lineSpacing
		const lineSpacingMultiple = paraOpts.lineSpacingMultiple ?? opts.lineSpacingMultiple
		return {
			runs: runsForPara,
			lineSpacingPts: typeof lineSpacing === 'number' && lineSpacing > 0 ? lineSpacing : undefined,
			lineSpacingPct:
				typeof lineSpacingMultiple === 'number' && lineSpacingMultiple > 0 ? lineSpacingMultiple * 100 : 100,
			spaceBeforePts: Number(paraOpts.paraSpaceBefore ?? opts.paraSpaceBefore ?? 0) || 0,
			spaceAfterPts: Number(paraOpts.paraSpaceAfter ?? opts.paraSpaceAfter ?? 0) || 0,
		}
	}

	const paras: FitParagraph[] = []
	let cur: FitRun[] = []
	let curParaOpts: RunOpts = opts
	for (const piece of pieces) {
		cur.push(toRun(piece))
		curParaOpts = piece.options
		if (piece.breakAfter) {
			paras.push(toPara(cur, curParaOpts))
			cur = []
		}
	}
	if (cur.length > 0) paras.push(toPara(cur, curParaOpts))
	return paras.length > 0 ? paras : null
}

interface InsetsEmu {
	lIns: number
	rIns: number
	tIns: number
	bIns: number
}

/**
 * Resolve a text frame's insets (EMU): explicit `_bodyProp` (from `inset`) → `margin` →
 * PowerPoint's defaults.
 *
 * Named for the question it answers, which is not the one `units-internal.resolveInsetsEmu`
 * answers: that one takes a `margin` alone and returns `null` when the caller stated none, so
 * an emitter can leave the attributes off. This one takes the whole option bag, consults
 * `_bodyProp` first, and always produces four numbers, because a solver has to measure against
 * something. Two exported functions with one name, imported from two places, was a coin-flip
 * for anyone reading a call site.
 *
 * The `[T, R, B, L]` shuffle and the defaults both come from the modules that own them rather
 * than being spelled again here.
 */
export function resolveBodyInsetsEmu(opts: RunOpts): InsetsEmu {
	const bp = opts._bodyProp ?? {}
	let lIns = bp.lIns
	let rIns = bp.rIns
	let tIns = bp.tIns
	let bIns = bp.bIns
	if (lIns == null && rIns == null && tIns == null && bIns == null) {
		const insets = resolveMarginInsetsEmu(opts.margin)
		if (insets) ({ l: lIns, r: rIns, t: tIns, b: bIns } = insets)
	}
	return {
		lIns: lIns ?? BODY_INSET_DEFAULTS_EMU.left,
		rIns: rIns ?? BODY_INSET_DEFAULTS_EMU.right,
		tIns: tIns ?? BODY_INSET_DEFAULTS_EMU.top,
		bIns: bIns ?? BODY_INSET_DEFAULTS_EMU.bottom,
	}
}

/**
 * Resolve the inner box (shape minus insets) in points; null if degenerate.
 *
 * Reads the object's own `opts.w/h` with no ancestor walk — correct only because a group keeps an
 * identity child coordinate space (`chOff/chExt == off/ext`, see the `gen/slide/object.ts` group renderer and
 * `docs/groups.md`), so a group never scales its children and a grouped text box's authored w/h is its
 * true rendered size. `applyMeasuredFit` reaches here for grouped text via its explicit group descent
 * (`measureObject`). If `addGroup` ever authored a scaled group (non-identity `chExt`), this would need
 * the same ancestor-scale composition `Shape.absoluteFrame` performs on the read path.
 */
export function computeBox(obj: SlideObject, presLayout: PresSlideInternal['_presLayout']): FitBox | null {
	const opts = (obj.options ?? {}) as RunOpts
	const wEmu = getSmartParseNumber(opts.w, 'X', presLayout)
	const hEmu = getSmartParseNumber(opts.h, 'Y', presLayout)
	if (!(wEmu > 0) || !(hEmu > 0)) return null

	const { lIns, rIns, tIns, bIns } = resolveBodyInsetsEmu(opts)

	// wrap=none lays text out one line per paragraph (no width-wrapping, handled by
	// the solver via the wrap flag), but the box width is still a real constraint:
	// solveShrink enforces it against the widest line so an over-wide non-wrapping
	// line shrinks to fit horizontally instead of spilling out of the box.
	const wrap = (opts._bodyProp ?? {}).wrap !== false
	const innerWidthPt = (wEmu - lIns - rIns) / EMU_PER_POINT
	const innerHeightPt = (hEmu - tIns - bIns) / EMU_PER_POINT
	if (!(innerHeightPt > 0)) return null
	if (!(innerWidthPt > 0)) return null
	return { innerWidthPt, innerHeightPt, wrap }
}

/** Vertical-anchor share of a height change that moves the box top up (`off.y` shift). */
export function anchorTopShareOfDelta(opts: RunOpts): number {
	// `_bodyProp.anchor` is the resolved valign ('t' | 'ctr' | 'b'); default 'ctr'.
	const anchor = (opts._bodyProp ?? {}).anchor
	if (anchor === TextAnchor.t) return 0 // grow downward — top fixed
	if (anchor === TextAnchor.b) return 1 // grow upward — bottom fixed
	return 0.5 // centered growth (default)
}
