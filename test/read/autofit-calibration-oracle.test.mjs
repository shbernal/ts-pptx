// Conservative-against-PowerPoint regression for the measured-fit shrink solver.
//
// The shrink solver (src/text-fit.ts) must never under-shrink: its computed
// `fontScale` has to be ≤ the value PowerPoint itself baked for the same box, so
// the text never overflows in PowerPoint or LibreOffice. This holds the solver to
// the PowerPoint-authored oracle (test/read/fixtures/autofit-*.cases.json +
// autofit-calibration.json) using the *real* fonts.
//
// Proprietary fonts (Aptos/Calibri/Tahoma/Arial) cannot be committed and are not
// present on CI, so each case is skipped unless `fc-match` resolves the genuine
// family (no substitution). On a workstation with the fonts installed this runs
// for real; on CI it degrades to a no-op rather than calibrating to a substitute.
import { readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import { describe, test, expect } from 'vitest'
import { solveShrink } from '../../src/text-fit.ts'
import { parseFontMetrics, FontMetricsRegistry } from '../../src/font-metrics.ts'

const FIX = resolve(dirname(fileURLToPath(import.meta.url)), 'fixtures')

/** Resolve a genuine font file for `family` (+style), or null if substituted/missing. */
function resolveFontFile(family, bold, italic) {
	try {
		const styleBits = [bold ? 'bold' : '', italic ? 'italic' : ''].filter(Boolean).join(' ')
		const pattern = styleBits ? `${family}:style=${styleBits}` : family
		const out = execFileSync('fc-match', ['-f', '%{family}\t%{file}', pattern], {
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'ignore'],
		})
		const [fam, file] = out.split('\t')
		if (!fam || !file) return null
		// Reject substitution: the resolved family must contain the requested name.
		if (!fam.toLowerCase().includes(family.toLowerCase())) return null
		return file.trim()
	} catch {
		return null
	}
}

// Cache the parsed metrics (or `false` if the genuine font is unavailable) across
// tests, but ALWAYS register into the current test's fresh registry on a hit.
const cache = new Map()
async function metricsFor(registry, family, bold, italic) {
	const key = `${family} ${bold} ${italic}`
	if (!cache.has(key)) {
		const file = resolveFontFile(family, bold, italic)
		cache.set(key, file ? await parseFontMetrics(new Uint8Array(readFileSync(file))) : false)
	}
	const m = cache.get(key)
	if (!m) return false
	registry.set(family, m, { bold, italic })
	return true
}

const calibration = JSON.parse(readFileSync(resolve(FIX, 'autofit-calibration.json'), 'utf8'))
const shrinkSpec = JSON.parse(readFileSync(resolve(FIX, 'autofit-shrink.cases.json'), 'utf8'))
const ppById = new Map()
for (const deck of calibration.decks) for (const c of deck.cases) ppById.set(c.id, c.powerpoint)

describe('autofit calibration oracle: shrink solver is conservative vs PowerPoint', () => {
	let ranAny = false

	for (const c of shrinkSpec.cases) {
		const pp = ppById.get(c.id)
		// Only cases where PowerPoint actually baked a fontScale are conservativeness targets.
		if (!pp || pp.fontScale == null) continue

		test(c.id, async () => {
			const registry = new FontMetricsRegistry()
			const resolve = (run) => registry.get(run.fontFace, !!run.bold, !!run.italic)

			// Register every distinct face this case uses; skip the case if any is unavailable.
			for (const para of c.paragraphs) {
				for (const run of para.runs) {
					const ok = await metricsFor(registry, run.font, !!run.bold, !!run.italic)
					if (!ok) {
						expect(true).toBe(true) // recorded as a (skipped) pass — font not installed
						return
					}
				}
			}
			ranAny = true

			const paragraphs = c.paragraphs.map((p) => ({
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

	test('at least one real-font case ran (informational)', () => {
		if (!ranAny)
			console.warn('autofit oracle: no genuine fonts resolved — conservativeness assertions skipped (expected on CI).')
		expect(true).toBe(true)
	})
})
