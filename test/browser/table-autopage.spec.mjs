import fs from 'node:fs'
import path from 'node:path'
import { explodePackage } from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'
import { AUTOPAGE_ROWS } from './harness/table-fixture.mjs'
import { expect, test } from './fixtures.mjs'
import { buildTableInHarness, buildTableInNode, openTableHarness, packageBytes } from './helpers.mjs'

/**
 * The headless-browser `tableToSlides` auto-paging repro — upstream gitbrent/PptxGenJS#1200.
 *
 * The backlog entry for that report was closed `non-target/out-of-project-scope` on the grounds
 * that "the sizing input that drives the overflow cannot be exercised here", and it invited
 * reopening with either a DOM-free `addTable`-autoPage repro or a headless-browser `tableToSlides`
 * one. This file is the second of those. It exists because the Tier 2b harness made it
 * constructible, not because a new report arrived.
 *
 * ── What it found ──────────────────────────────────────────────────────────────────────────────
 *
 * It reproduces, and the cause turned out to be DOM-free after all: the auto-pager charged each
 * row its cells' top/bottom margins and then, on a page break, zeroed that accumulator — so the
 * first row of every continuation page was placed without paying for its margins, and a page could
 * accept one row more than it had room for. Nothing about that needed a browser; the report's
 * browser flavour is why it read as a layout question for two years. The fix and the DOM-free
 * regression are in `src/gen/table/autopage.ts` and
 * test/regression/table/table-autopage-continuation-budget.test.js.
 *
 * ── What is and is not claimed ─────────────────────────────────────────────────────────────────
 *
 * Claimed: with every row the same height and every page the same usable height, every full page
 * holds the same number of rows, and the pages together hold every row exactly once. Both are
 * arithmetic about what was emitted, checked against the pager's own stated geometry.
 *
 * NOT claimed: that the rows are the height PowerPoint will draw them, or that the table ends up
 * where the slide's bottom edge is in a renderer. That is layout fidelity, it has no oracle here,
 * and it stays out of scope (docs/project-target.md "Out Of Active Scope"). The fixture is built
 * so the weaker claim is enough — see harness/table-fixture.mjs.
 */

const OUT_ROOT = path.join(ROOT, '.tmp', 'browser-table', 'autoPage')

/** Slide parts of an exploded package, in slide order. */
const slideParts = (dir) =>
	fs
		.readdirSync(path.join(dir, 'ppt', 'slides'), { withFileTypes: true })
		.filter((entry) => entry.isFile() && /^slide\d+\.xml$/.test(entry.name))
		.map((entry) => entry.name)
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
		.map((name) => fs.readFileSync(path.join(dir, 'ppt', 'slides', name), 'utf8'))

/** `<a:tr>` count per slide — how many rows the pager put on each page. */
const rowsPerSlide = (dir) => slideParts(dir).map((xml) => (xml.match(/<a:tr /g) || []).length)

/** Every `R<n>` key cell, in emission order across the whole deck. */
const keyCells = (dir) =>
	slideParts(dir)
		.flatMap((xml) => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) => match[1]))
		.filter((text) => /^R\d+$/.test(text))

async function explodeBrowserBuild(page) {
	const outcome = await buildTableInHarness(page, 'autoPage')
	expect(outcome.ok, `the harness failed to convert the autoPage table: ${outcome.message}`).toBe(true)
	return await explodePackage(packageBytes(outcome.base64), path.join(OUT_ROOT, 'browser'))
}

test.beforeEach(async ({ page }) => {
	await openTableHarness(page)
})

test('a table too tall for one slide pages onto several, and no page takes more rows than the first', async ({
	page,
}) => {
	const counts = rowsPerSlide(await explodeBrowserBuild(page))

	expect(counts.length, `the table should have paged; got ${JSON.stringify(counts)}`).toBeGreaterThan(2)

	// The last page holds the remainder and is allowed to be short. Every page before it was
	// filled to the pager's budget, and since the rows are identical and the geometry is the same
	// on every page, one budget is the only self-consistent answer. Before the fix this read
	// `[10, 11, 11, 11, 11, 7]` — the first page charged every row its cell margins, the
	// continuation pages let their first row in free, and that row hung off the slide.
	const full = counts.slice(0, -1)
	expect(
		new Set(full).size,
		`every full page must hold the same number of identical rows; got ${JSON.stringify(counts)}`
	).toBe(1)
	// A budget of one row per page would satisfy the line above and mean the arithmetic had
	// collapsed rather than held.
	expect(full[0], `expected a real page budget; got ${full[0]} rows per page`).toBeGreaterThan(3)
})

test('every source row crosses to exactly one slide, in order', async ({ page }) => {
	const keys = keyCells(await explodeBrowserBuild(page))

	// The header is `Key`/`Value`, so only the body rows match `R<n>` — one per source row, once.
	expect(keys).toEqual(Array.from({ length: AUTOPAGE_ROWS }, (_, idx) => `R${idx + 1}`))
})

test('the page budget is the same in Chromium as on a DOM that renders nothing', async ({ page }) => {
	const browserCounts = rowsPerSlide(await explodeBrowserBuild(page))
	const nodeDir = await explodePackage(packageBytes(await buildTableInNode('autoPage')), path.join(OUT_ROOT, 'node'))

	// This is the assertion that places the defect, and it is the reason the backlog entry moved
	// out of the browser bucket rather than into it. The two runtimes size the *columns*
	// differently — Chromium measures, happy-dom falls back, which is what table-widths.spec.mjs
	// pins — but the cells are short enough that no column width makes them wrap, so the vertical
	// arithmetic has the same inputs on both sides. It therefore has to reach the same answer, and
	// a report of rows running off the bottom is not a report about a rendered page.
	expect(rowsPerSlide(nodeDir)).toEqual(browserCounts)
})
