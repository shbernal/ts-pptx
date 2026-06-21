/**
 * PptxGenJS: Measured-fit serialization pass
 *
 * Bridges slide text objects to the line-break simulator / shrink solver
 * (`text-fit.ts`). Runs during async export, BEFORE `gen-xml` builds the body:
 * for each text box with `fit: 'shrink'` and a registered font, it computes the
 * `fontScale` PowerPoint would have baked and rewrites `options.fit` to the
 * explicit object form (`{ type:'shrink', fontScale }`) so `genXmlNormAutofit`
 * emits `<a:normAutofit fontScale=…/>`. Without metrics it leaves the bare flag
 * untouched (current behavior) and warns once. See `PLAN-measured-text-fit.md`.
 */

import { SLIDE_OBJECT_TYPES } from './core-enums.js'
import { DEF_FONT_SIZE } from './core-enums.js'
import { EMU_PER_POINT } from './units.js'
import { getSmartParseNumber, valToPts } from './gen-utils.js'
import type { FontMetricsRegistry } from './font-metrics.js'
import { solveShrink, type FitParagraph, type FitRun, type MetricsResolver } from './text-fit.js'
import type { ISlideObject, ObjectOptions, PresSlideInternal, TextProps, TextPropsOptions } from './core-interfaces.js'

// PowerPoint's default text-frame insets (EMU): l/r = 0.1in, t/b = 0.05in.
const DEF_INS_LR_EMU = 91440
const DEF_INS_TB_EMU = 45720

const CRLF_RE = /\r*\n/g

type RunOpts = TextPropsOptions & ObjectOptions

/** Normalize `slideObj.text` (string | TextProps | TextProps[]) to a run list. */
function normalizeRuns (obj: ISlideObject): TextProps[] {
	const opts = (obj.options ?? {}) as ObjectOptions
	const text = obj.text as unknown
	if (text == null) return []
	if (typeof text === 'string' || typeof text === 'number') return [{ text: String(text), options: opts }]
	if (!Array.isArray(text) && typeof text === 'object' && 'text' in (text as object)) {
		const t = text as TextProps
		return [{ text: t.text, options: t.options ?? opts }]
	}
	if (Array.isArray(text)) return (text as TextProps[]).map(t => ({ text: t.text, options: t.options ?? opts }))
	return []
}

/** Build a measurable `FitParagraph[]` from a text object, or null if not measurable. */
function extractParagraphs (obj: ISlideObject): FitParagraph[] | null {
	const opts = (obj.options ?? {}) as RunOpts
	const runs = normalizeRuns(obj)
	if (runs.length === 0) return null

	// Expand "\n" inside a run into separate pieces, flagging the paragraph break
	// after each (mirrors gen-xml STEP 4). `breakLine` ends a paragraph too.
	interface Piece { text: string, options: RunOpts, breakAfter: boolean }
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
			lineSpacingPct: typeof lineSpacingMultiple === 'number' && lineSpacingMultiple > 0 ? lineSpacingMultiple * 100 : 100,
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

/** Resolve the inner box (shape minus insets) in points; null if degenerate. */
function computeBox (obj: ISlideObject, presLayout: PresSlideInternal['_presLayout']): { innerWidthPt: number, innerHeightPt: number } | null {
	const opts = (obj.options ?? {}) as RunOpts
	const wEmu = getSmartParseNumber(opts.w, 'X', presLayout)
	const hEmu = getSmartParseNumber(opts.h, 'Y', presLayout)
	if (!(wEmu > 0) || !(hEmu > 0)) return null

	// Insets: explicit `_bodyProp` (set from `inset`) → `margin` → PowerPoint defaults.
	const bp = opts._bodyProp ?? {}
	const margin = opts.margin
	let lIns = bp.lIns
	let rIns = bp.rIns
	let tIns = bp.tIns
	let bIns = bp.bIns
	if (lIns == null && rIns == null && tIns == null && bIns == null && margin != null) {
		if (Array.isArray(margin)) {
			tIns = valToPts(margin[0] ?? 0)
			rIns = valToPts(margin[1] ?? 0)
			bIns = valToPts(margin[2] ?? 0)
			lIns = valToPts(margin[3] ?? 0)
		} else if (typeof margin === 'number') {
			lIns = rIns = tIns = bIns = valToPts(margin)
		}
	}
	lIns = lIns ?? DEF_INS_LR_EMU
	rIns = rIns ?? DEF_INS_LR_EMU
	tIns = tIns ?? DEF_INS_TB_EMU
	bIns = bIns ?? DEF_INS_TB_EMU

	// wrap=none means no width wrapping (only hard breaks): unbounded line width.
	const wrap = bp.wrap !== false
	const innerWidthPt = wrap ? (wEmu - lIns - rIns) / EMU_PER_POINT : Infinity
	const innerHeightPt = (hEmu - tIns - bIns) / EMU_PER_POINT
	if (!(innerHeightPt > 0)) return null
	if (wrap && !(innerWidthPt > 0)) return null
	return { innerWidthPt, innerHeightPt }
}

/**
 * Apply measured shrink-to-fit across every slide. Mutates `options.fit` in place
 * for text boxes that opt in via `fit:'shrink'` and whose font has registered
 * metrics. Safe to call with an empty registry (no-op). Warns once if any opted-in
 * box could not be measured (missing metrics) so overflow is not silently ignored.
 */
export function applyMeasuredFit (slides: PresSlideInternal[], registry: FontMetricsRegistry): void {
	if (registry.size === 0) return

	const resolve: MetricsResolver = run => registry.get(run.fontFace, run.bold, run.italic)
	const unmeasuredFaces = new Set<string>()

	for (const slide of slides) {
		for (const obj of slide._slideObjects ?? []) {
			if (obj._type !== SLIDE_OBJECT_TYPES.text) continue
			// Only the bare string form opts into measurement; an explicit object form is
			// already baked by the caller and 'none'/'resize' are out of scope for P1.
			if (obj.options?.fit !== 'shrink') continue

			const paragraphs = extractParagraphs(obj)
			if (!paragraphs) continue
			const box = computeBox(obj, slide._presLayout)
			if (!box) continue

			const outcome = solveShrink(paragraphs, box, resolve)
			if (outcome.kind === 'shrink') {
				const { fontScalePct, lnSpcReductionPct } = outcome.result
				obj.options.fit = {
					type: 'shrink',
					fontScale: fontScalePct,
					lnSpcReduction: lnSpcReductionPct || undefined,
				}
			} else if (outcome.kind === 'unmeasurable') {
				for (const para of paragraphs) for (const run of para.runs) if (!resolve(run)) unmeasuredFaces.add(run.fontFace ?? '(theme default)')
			}
			// 'fits' → leave the bare flag; the text already fits, so no scale is needed.
		}
	}

	if (unmeasuredFaces.size > 0) {
		console.warn(
			`Warning: fit:'shrink' could not be measured for font(s) [${[...unmeasuredFaces].join(', ')}] — ` +
				'no registered metrics. Emitting bare <a:normAutofit/> (text will not pre-shrink in headless renders). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) to enable measured fit.'
		)
	}
}
