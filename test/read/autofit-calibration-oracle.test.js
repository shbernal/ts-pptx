// Conservative-against-PowerPoint regression for the measured-fit shrink solver.
//
// The shrink solver (src/measure/text-fit.ts) must never under-shrink: its computed
// `fontScale` has to be ≤ the value PowerPoint itself baked for the same box, so
// the text never overflows in PowerPoint or LibreOffice. This holds the solver to
// the PowerPoint-authored oracle (test/read/fixtures/autofit-*.cases.json +
// autofit-calibration.json), measured with the *real* fonts or with the advances
// recorded from them.
//
// Proprietary fonts (Aptos/Calibri/Tahoma/Arial) cannot be committed, so the faces come
// from `font-oracle.js`: the genuine installed file where the machine has one, and the
// committed metrics sidecar otherwise. Under `FONT_ORACLES=required` a face that resolves
// through neither is a failure rather than a skip, which is what CI sets.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import { describe, test, expect } from 'vitest'
import { solveShrink, solveResize } from '../../src/measure/text-fit.ts'
import { FontMetricsRegistry } from '../../src/measure/font-metrics.ts'
import { oracleMetrics, resolutionTally } from './font-oracle.js'

const EMU_PER_PT = 12700

const FIX = resolve(dirname(fileURLToPath(import.meta.url)), 'fixtures')

/**
 * Register every face a case's runs name, and report whether all of them resolved.
 *
 * @param {import('../../src/measure/font-metrics.ts').FontMetricsRegistry} registry
 * @param {{ paragraphs: Array<{ runs: Array<{ font: string, bold?: boolean, italic?: boolean }> }> }} c
 */
async function registerCaseFaces(registry, c) {
	for (const para of c.paragraphs) {
		for (const run of para.runs) {
			const face = { family: run.font, bold: !!run.bold, italic: !!run.italic }
			const metrics = await oracleMetrics(face)
			if (!metrics) return false
			registry.set(face.family, metrics, { bold: face.bold, italic: face.italic })
		}
	}
	return true
}

const calibration = JSON.parse(readFileSync(resolve(FIX, 'autofit-calibration.json'), 'utf8'))
const shrinkSpec = JSON.parse(readFileSync(resolve(FIX, 'autofit-shrink.cases.json'), 'utf8'))
const resizeSpec = JSON.parse(readFileSync(resolve(FIX, 'autofit-resize.cases.json'), 'utf8'))
const ppById = new Map()
const loById = new Map()
for (const deck of calibration.decks) {
	for (const c of deck.cases) {
		ppById.set(c.id, c.powerpoint)
		if (c.libreoffice) loById.set(c.id, c.libreoffice)
	}
}

/** Build the FitParagraph[] a case describes (shared by the shrink/resize oracles). */
function paragraphsOf(c) {
	return c.paragraphs.map((p) => ({
		runs: p.runs.map((r) => ({
			text: r.text,
			sizePt: r.sizePt,
			bold: !!r.bold,
			italic: !!r.italic,
			fontFace: r.font,
			charSpacingPt: r.charSpacingPts ?? undefined,
		})),
		lineSpacingPct: p.lineSpacingPct,
		lineSpacingPts: p.lineSpacingPts,
		spaceBeforePts: p.spaceBeforePts,
		spaceAfterPts: p.spaceAfterPts,
	}))
}

// Set by every case that got past face resolution, in BOTH blocks. File scope because the
// accounting test now sits below them and asserts on both.
let ranAny = false

describe('autofit calibration oracle: shrink solver is conservative vs PowerPoint', () => {
	for (const c of shrinkSpec.cases) {
		const pp = ppById.get(c.id)
		// Only cases where PowerPoint actually baked a fontScale are conservativeness targets.
		if (!pp || pp.fontScale == null) continue

		test(c.id, async (ctx) => {
			const registry = new FontMetricsRegistry()
			const resolve = (run) => registry.get(run.fontFace, !!run.bold, !!run.italic)

			if (!(await registerCaseFaces(registry, c))) {
				ctx.skip('a face this case uses resolved neither an installed font nor a sidecar entry')
				return
			}
			ranAny = true

			const paragraphs = paragraphsOf(c)
			const box = {
				innerWidthPt: c.wPt - (c.insetsPt?.l ?? 0) - (c.insetsPt?.r ?? 0),
				innerHeightPt: c.hPt - (c.insetsPt?.t ?? 0) - (c.insetsPt?.b ?? 0),
			}

			const out = solveShrink(paragraphs, box, resolve)
			const computedPct = out.kind === 'shrink' ? out.result.fontScalePct : 100
			const ppPct = pp.fontScale / 1000

			// CONSERVATIVE: computed scale must be ≤ PowerPoint's (shrink at least as much).
			expect(computedPct).toBeLessThanOrEqual(ppPct)
		})
	}
})

describe('autofit calibration oracle: resize solver is conservative vs PowerPoint + LibreOffice', () => {
	for (const c of resizeSpec.cases) {
		const pp = ppById.get(c.id)
		const lo = loById.get(c.id)
		if (!pp || pp.extCyEmu == null) continue

		test(c.id, async (ctx) => {
			const registry = new FontMetricsRegistry()
			const resolve = (run) => registry.get(run.fontFace, !!run.bold, !!run.italic)

			if (!(await registerCaseFaces(registry, c))) {
				ctx.skip('a face this case uses resolved neither an installed font nor a sidecar entry')
				return
			}
			ranAny = true

			const paragraphs = paragraphsOf(c)
			const box = {
				innerWidthPt: c.wrap === false ? Infinity : c.wPt - (c.insetsPt?.l ?? 0) - (c.insetsPt?.r ?? 0),
				innerHeightPt: c.hPt - (c.insetsPt?.t ?? 0) - (c.insetsPt?.b ?? 0),
			}

			const out = solveResize(paragraphs, box, resolve)
			if (out.kind !== 'resize') throw new Error(`expected a resize outcome, got ${out.kind}`)
			// Bake the same way measure-fit does: needed inner height + top/bottom insets.
			const insetsPt = (c.insetsPt?.t ?? 0) + (c.insetsPt?.b ?? 0)
			const computedCyEmu = Math.round(out.neededInnerHeightPt * EMU_PER_PT) + Math.round(insetsPt * EMU_PER_PT)

			// CONSERVATIVE (resize has no safety net): computed cy must be ≥ PowerPoint's
			// baked height AND ≥ the LibreOffice-rendered height, so text never overflows.
			expect(computedCyEmu).toBeGreaterThanOrEqual(pp.extCyEmu)
			if (lo?.hEmu != null) expect(computedCyEmu).toBeGreaterThanOrEqual(lo.hEmu)
		})
	}
})

// Accounting, not decoration. This test used to pass unconditionally and only warn, so a
// run that resolved nothing and asserted nothing was indistinguishable from a run that
// proved every case. It now fails when NEITHER solver measured anything, and reports where
// the advances came from so a leg that quietly fell back to the sidecar says so.
//
// Last in the file, and outside both blocks, deliberately. `missing` tallies what has been
// ASKED for so far, so while this sat at the foot of the shrink block a face only the resize
// cases name was not yet in it — the one place a dropped font would have shown up was the
// one place it could not be seen. Describes run in source order, so by here both have run.
test('the conservativeness assertions actually ran', () => {
	const { genuine, sidecar, missing } = resolutionTally()
	console.info(`autofit oracle: ${genuine} face(s) from installed fonts, ${sidecar} from the metrics sidecar.`)
	expect(missing).toEqual([])
	expect(ranAny).toBe(true)
})
