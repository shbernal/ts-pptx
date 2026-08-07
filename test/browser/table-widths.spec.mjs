import fs from 'node:fs'
import path from 'node:path'
import { explodePackage } from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'
import { expect, test } from './fixtures.mjs'
import { buildTableInHarness, buildTableInNode, openTableHarness, packageBytes, tableBases } from './helpers.mjs'

/**
 * `tableToSlides` against a table a browser actually laid out.
 *
 * `pickColWidthBasis` (src/gen/table/html-dom.ts) chooses between three width bases, and
 * until this spec existed the first one — the rendered `offsetWidth` — had never executed
 * anywhere in the repo. The Node suite drives happy-dom, where `offsetWidth` is `0` for
 * every cell, so it always takes one of the two fallback arms; the unit suite reaches
 * `pickColWidthBasis` by handing it numbers directly, which proves the `if` and not the
 * pipeline behind it. The *primary* path of the feature was therefore covered only at its
 * own function boundary — including the `arrColSrc` arithmetic that fixed
 * gitbrent/PptxGenJS#1244, which is built in the same pass that reads `offsetWidth`.
 *
 * ── What is and is not claimed ─────────────────────────────────────────────────────────
 *
 * Claimed: the measured arm runs, and the grid it produces is proportional to what was
 * measured, with `data-pptx-width` still winning outright. Every assertion below is
 * checked against numbers read from the same live page, so the oracle is the function's
 * own stated contract.
 *
 * NOT claimed: that Chromium's measurement is the *right* measurement, or that another
 * engine would produce the same one. That is live-DOM layout fidelity, it has no oracle,
 * and it stays out of scope — see docs/project-target.md "Out Of Active Scope". A layout
 * difference between two browsers is not a defect in this package; a `.pptx` a browser
 * builds differently from Node *is*, which is what the last test here pins.
 */

const OUT_ROOT = path.join(ROOT, '.tmp', 'browser-table')
const ONE_IN_EMU = 914400

/** Proportions relative to the first column — the only thing a basis determines. */
const ratios = (widths) => widths.map((width) => width / widths[0])

/** `<a:gridCol w="…"/>`, in column order. The same read the Node table suites make. */
const gridColWidths = (xml) => [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((match) => Number(match[1]))

/**
 * Build one fixture in the page and return its emitted grid.
 *
 * The package is exploded to `.tmp/` rather than parsed in memory, for the same reason
 * adapter-media.spec.mjs does it: on a failure the tree is what makes the emitted XML
 * readable without re-running the lane.
 */
async function browserGrid(page, scenario) {
	const outcome = await buildTableInHarness(page, scenario)
	expect(outcome.ok, `the harness failed to convert the "${scenario}" table: ${outcome.message}`).toBe(true)
	const dir = await explodePackage(packageBytes(outcome.base64), path.join(OUT_ROOT, scenario, 'browser'))
	return gridColWidths(fs.readFileSync(path.join(dir, 'ppt', 'slides', 'slide1.xml'), 'utf8'))
}

test.beforeEach(async ({ page }) => {
	await openTableHarness(page)
})

test('the measured fixture discriminates: offsetWidth and computed CSS disagree about the columns', async ({
	page,
}) => {
	const { measured, css } = await tableBases(page, 'measured')

	// Guard, not decoration. Every other assertion in this file reads "the widths came out
	// proportional to X" — and if the two bases agreed, that sentence would be equally true
	// of a run where the measured arm never executed. The fixture earns its keep only while
	// they differ, and a browser changing how it resolves `width` on a table cell is exactly
	// the kind of thing that would quietly end that without failing anything.
	expect(measured).toHaveLength(2)
	expect(
		measured.every((width) => width > 0),
		`the table was not laid out: ${measured.join(', ')}`
	).toBe(true)

	// offsetWidth is the border box, computed `width` the content box: column B pads its way
	// to A's width while stating a narrower one. See harness/table-fixture.mjs.
	const cssPx = css.map((value) => Number(String(value).replace('px', '')))
	expect(cssPx.every(Number.isFinite), `computed widths are not px: ${css.join(', ')}`).toBe(true)

	const measuredRatio = ratios(measured)[1]
	const cssRatio = ratios(cssPx)[1]
	expect(
		Math.abs(measuredRatio - cssRatio),
		`the two bases have converged (measured ${measured.join(':')}, css ${css.join(':')}) — ` +
			'this fixture can no longer say which arm produced the grid'
	).toBeGreaterThan(0.1)
})

test('the rendered offsetWidth is what sizes the columns, not the computed CSS width', async ({ page }) => {
	const { measured, css } = await tableBases(page, 'measured')
	const cssPx = css.map((value) => Number(String(value).replace('px', '')))
	const cols = await browserGrid(page, 'measured')

	expect(cols).toHaveLength(2)

	// The proportional calc rounds to 2dp of an inch before converting to EMU, so a
	// percentage-point of tolerance is the rounding, not slack in the claim.
	const emitted = ratios(cols)[1]
	expect(
		Math.abs(emitted - ratios(measured)[1]),
		`emitted ${cols.join(':')} does not follow the measured basis ${measured.join(':')}`
	).toBeLessThan(0.01)

	// Stated as its own assertion rather than left implied by the one above: this is the
	// half that fails if `pickColWidthBasis` ever silently prefers the CSS arm.
	expect(
		Math.abs(emitted - ratios(cssPx)[1]),
		`emitted ${cols.join(':')} followed the CSS basis ${cssPx.join(':')} — the measured arm did not run`
	).toBeGreaterThan(0.1)
})

test('data-pptx-width still wins outright against a live measurement', async ({ page }) => {
	const { measured } = await tableBases(page, 'override')
	const cols = await browserGrid(page, 'override')

	expect(cols).toHaveLength(2)
	// Exact, not proportional: an override states an absolute width in inches and nothing
	// downstream rescales it. The Node suite proves this against a basis of zeroes, which is
	// the easy half — here it has a real 1:3 measurement to override.
	expect(cols[0], `expected exactly 4in; got ${cols[0]} EMU`).toBe(4 * ONE_IN_EMU)
	expect(measured[1] / measured[0], 'the fixture should measure column B three times column A').toBeCloseTo(3, 1)
	// And the un-overridden column still came from the measurement rather than inheriting
	// the override.
	expect(cols[1]).not.toBe(cols[0])
})

test('a data-pptx-width on a spanning cell divides across the columns it covers, under measurement', async ({
	page,
}) => {
	const cols = await browserGrid(page, 'spanOverride')

	// upstream gitbrent/PptxGenJS#1244. The override used to be indexed by *column* against
	// a row indexed by *cell*, so a span made the two part ways. `arrColSrc` is built in the
	// same pass that reads `offsetWidth`, so this is the first time the fix is exercised
	// with a real measurement alongside it — the combination an actual consumer has.
	expect(cols).toHaveLength(3)
	expect(cols[0], 'first spanned column takes half of the 4in span').toBe(2 * ONE_IN_EMU)
	expect(cols[1], 'second spanned column takes the other half').toBe(2 * ONE_IN_EMU)
	expect(cols[2], 'the unspanned column keeps its own 2in override').toBe(2 * ONE_IN_EMU)
})

test('cross-runtime: the same table degrades to the CSS basis on Node and is measured in the browser', async ({
	page,
}) => {
	const browserCols = await browserGrid(page, 'measured')
	const nodeDir = await explodePackage(
		packageBytes(await buildTableInNode('measured')),
		path.join(OUT_ROOT, 'measured', 'node')
	)
	const nodeCols = gridColWidths(fs.readFileSync(path.join(nodeDir, 'ppt', 'slides', 'slide1.xml'), 'utf8'))

	// This is the one place in the browser lane where the two runtimes are *supposed* to
	// disagree about a slide part, and the disagreement is the documented degradation
	// itself: without a layout engine `offsetWidth` is 0, so the widths fall back to the
	// computed CSS widths. Asserting its shape is what stops it quietly becoming a
	// different disagreement — the same discipline adapter-media.spec.mjs applies to the
	// SVG placeholder.
	expect(nodeCols).toHaveLength(2)
	expect(browserCols).toHaveLength(2)
	expect(ratios(browserCols)[1], 'the browser measures both columns at the same border-box width').toBeCloseTo(1, 2)
	expect(ratios(nodeCols)[1], 'Node falls back to the stated content-box widths, which are 2:1').toBeCloseTo(0.5, 2)
})
