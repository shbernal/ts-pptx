// PowerPoint-authored oracle for CJK line breaking in the measured-fit wrap model.
//
// The wrap simulator (src/measure/text-fit.ts) has to break Chinese and Japanese
// text between characters rather than only at whitespace, because that is what
// PowerPoint does — and has to *not* do it for Hangul, because PowerPoint does not.
// Neither claim can be settled against prose or against self-generated XML, so both
// are pinned to `autofit-cjk-wrap.pptx`: a deck desktop PowerPoint authored, one
// fixed-width `spAutoFit` box per case, with PowerPoint's own layout recorded in
// `autofit-cjk-wrap.oracle.json` by `fixtures/authoring/author-cjk-wrap.ps1`.
//
// Two suites, because the two halves of that sidecar have different reachability:
//
//   1. `bakedHeightPt` is `a:ext/@cy` in the committed package, so it is re-derived
//      from the deck here and compared. That runs everywhere and is what keeps the
//      sidecar honest — a hand-edited oracle stops matching its own deck.
//   2. `lines` / `lineCount` came from `TextRange.Lines()` over COM at authoring
//      time (nothing in the package records where a line broke). Checking the model
//      against them needs the real Malgun Gothic advances, so those cases skip when
//      the genuine font does not resolve — the same degradation contract as
//      `autofit-calibration-oracle.test.js`, and expected on CI.
import { readFileSync, existsSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { unzipSync } from 'fflate'
import { describe, test, expect } from 'vitest'
import { measureLayout, WIDTH_SAFETY_FACTOR, HEIGHT_SAFETY_FACTOR } from '../../src/measure/text-fit.ts'
import { collectUncoveredCodepoints, parseFontMetrics, FontMetricsRegistry } from '../../src/measure/font-metrics.ts'
import { fixturePath, readOracle } from './corpus.js'

const EMU_PER_PT = 12700
const DECK = 'autofit-cjk-wrap'

const oracle = await readOracle(DECK)

/**
 * Resolve a genuine font file for `family`, or null when it is missing or would be
 * substituted. Malgun Gothic ships with Windows, so the Windows font directory is
 * checked directly; elsewhere `fc-match` decides, and a substituted family counts
 * as missing.
 *
 * @param {string} family
 * @returns {string | null}
 */
function resolveFontFile(family) {
	if (process.platform === 'win32') {
		const win = process.env.SystemRoot ?? 'C:\\Windows'
		const file = `${win}\\Fonts\\malgun.ttf`
		return family === 'Malgun Gothic' && existsSync(file) ? file : null
	}
	try {
		const out = execFileSync('fc-match', ['-f', '%{family}\t%{file}', family], {
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'ignore'],
		})
		const [fam, file] = out.split('\t')
		if (!fam || !file) return null
		if (!fam.toLowerCase().includes(family.toLowerCase())) return null
		return file.trim()
	} catch {
		return null
	}
}

const fontFile = resolveFontFile(oracle.fontFace)
const metrics = fontFile ? await parseFontMetrics(new Uint8Array(readFileSync(fontFile))) : null
const registry = new FontMetricsRegistry()
if (metrics) registry.set(oracle.fontFace, metrics)
const resolve = (run) => registry.get(run.fontFace, !!run.bold, !!run.italic)

describe('CJK oracle: the sidecar still describes the committed deck', () => {
	// One slide, one box per case, each named after its case id.
	const zip = unzipSync(new Uint8Array(readFileSync(fixturePath(`${DECK}.pptx`))))
	const slide = new TextDecoder().decode(zip['ppt/slides/slide1.xml'])
	// Split on shape starts so each chunk carries exactly one shape's name and ext.
	const shapes = new Map(
		slide
			.split('<p:sp>')
			.slice(1)
			.map((sp) => [/<p:cNvPr [^>]*name="([^"]*)"/.exec(sp)?.[1], sp])
	)

	test('every case has a named box in the deck', () => {
		expect([...shapes.keys()].filter(Boolean).sort()).toEqual(oracle.cases.map((c) => c.id).sort())
	})

	for (const c of oracle.cases) {
		test(`${c.id}: baked a:ext/@cy matches the recorded height`, () => {
			const sp = shapes.get(c.id)
			expect(sp, `no shape named ${c.id}`).toBeTruthy()
			expect(sp).toContain('<a:spAutoFit/>')
			const cy = Number(/<a:ext cx="\d+" cy="(\d+)"\/>/.exec(sp)?.[1])
			expect(cy / EMU_PER_PT).toBeCloseTo(c.bakedHeightPt, 3)
		})
	}
})

describe(`CJK oracle: the wrap model reproduces PowerPoint's line breaking`, () => {
	if (!metrics) {
		test(`skipped: ${oracle.fontFace} did not resolve (expected on CI)`, () => {
			console.warn(`CJK oracle: ${oracle.fontFace} not installed — line-breaking assertions skipped.`)
			expect(true).toBe(true)
		})
		return
	}

	// Malgun Gothic has no halfwidth Katakana and no Plane 2, and PowerPoint quietly
	// falls back to another face for them while the registry has no fallback at all
	// and charges its default advance. Those cases still carry the break evidence in
	// `lines` — PowerPoint split both runs mid-run — but their *widths* come from a
	// font this model never saw, so the arithmetic cannot be reproduced here.
	const covered = (text) => [...text].every((ch) => metrics.hasCodepoint(ch.codePointAt(0)))

	for (const c of oracle.cases) {
		test(`${c.id}: ${c.lineCount} line(s)`, (ctx) => {
			const paragraphs = [{ runs: [{ text: c.text, sizePt: c.sizePt, fontFace: c.fontFace }] }]
			if (!covered(c.text)) {
				// The widths are unreproducible, but the model still has to SAY so: this is
				// the one gap that can measure short rather than tall, so it is reported
				// rather than absorbed (docs/measured-text-fit.md, "No font fallback").
				const byFace = new Map()
				collectUncoveredCodepoints(paragraphs, registry, byFace)
				expect(byFace.get(c.fontFace)?.size ?? 0).toBeGreaterThan(0)
				ctx.skip(`${c.fontFace} does not cover every code point in this case; PowerPoint fell back, the model cannot`)
				return
			}
			const innerWidthPt = c.boxWidthPt - c.insetLeftPt - c.insetRightPt
			const layout = measureLayout(paragraphs, innerWidthPt, resolve, 100, 0, WIDTH_SAFETY_FACTOR)
			expect(layout).not.toBeNull()

			// The line count is the claim: it is what a per-character break opportunity
			// changes, and what a wrong one costs (a phantom line shrinks text that fits,
			// a missing one lets text overflow). Break *positions* are deliberately not
			// asserted — WIDTH_SAFETY_FACTOR wraps the model a hair early on purpose, so
			// it can land one character short of PowerPoint's break without being wrong.
			expect(layout.lineCount).toBe(c.lineCount)

			// And the height stays on the conservative side of what PowerPoint baked,
			// which is the contract the resize bake depends on (docs/measured-text-fit.md).
			const innerBakedPt = c.bakedHeightPt - c.insetTopPt - c.insetBottomPt
			expect(layout.heightPt * HEIGHT_SAFETY_FACTOR).toBeGreaterThanOrEqual(innerBakedPt)
		})
	}

	// Pinned so it cannot change silently, in either direction: if kinsoku is ever
	// implemented this fails and the doc caveat comes out with it.
	test('known gap: no kinsoku, so the widest line is narrower than PowerPoint hangs it', () => {
		const c = oracle.cases.find((x) => x.id === 'cjk__kinsoku_hanging_comma')
		const paragraphs = [{ runs: [{ text: c.text, sizePt: c.sizePt, fontFace: c.fontFace }] }]
		const innerWidthPt = c.boxWidthPt - c.insetLeftPt - c.insetRightPt
		const layout = measureLayout(paragraphs, innerWidthPt, resolve, 100, 0, WIDTH_SAFETY_FACTOR)

		// PowerPoint refuses to start a line with U+3001 and hangs it past the inset:
		// its first line measures wider than the box it is laid out in.
		expect(c.lineWidthsPt[0]).toBeGreaterThan(innerWidthPt)
		// The model has no such rule, so it breaks before the comma and stays inside.
		expect(layout.widestLineWidthPt).toBeLessThanOrEqual(innerWidthPt)
	})
})
