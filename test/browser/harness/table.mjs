// The browser side of the rendered-table lane: loads the *shipped* `dist/browser.js`
// unbundled, renders one fixture table into the page so a real layout engine sizes it, and
// exposes what the specs need through `page.evaluate`.
//
// Sibling of ./harness.mjs, and separate from it on purpose. That harness answers "do the
// four `RuntimeAdapter` loaders work in a browser"; this one answers the one question no
// lane in this repo could previously ask: does `tableToSlides` behave when `offsetWidth` is
// a real number rather than the 0 every DOM outside a browser reports? Mixing them would
// put a `<table>` in the adapter fixture's DOM for no reason and blur which page proves
// what.
//
// What this is NOT is a layout-fidelity fixture. Nothing here asserts that Chromium's
// numbers are the right numbers, or that another engine would agree — see
// ./table-fixture.mjs and docs/project-target.md "Out Of Active Scope".

import TsPptx from '../../../dist/browser.js'
import { TABLE_HTML, TABLE_ID } from './table-fixture.mjs'

/**
 * Render one fixture into the page and hand back its table element.
 *
 * Reading `offsetWidth` immediately after is what forces the synchronous layout the whole
 * lane depends on, so nothing here has to wait for a frame.
 *
 * @param {string} scenario a key of `TABLE_HTML`
 * @returns {HTMLTableElement} the rendered `<table>`
 */
function render(scenario) {
	const markup = TABLE_HTML[scenario]
	if (!markup) throw new Error(`unknown table fixture: ${scenario}`)
	const host = document.getElementById('fixture')
	host.innerHTML = markup
	// Narrowed rather than left as the `HTMLElement | null` `getElementById` returns: the
	// fixture below reads `.rows`, and the two guards that follow are what make the cast
	// safe — a missing element and an unlaid-out one both throw before it is used.
	const table = /** @type {HTMLTableElement | null} */ (document.getElementById(TABLE_ID))
	if (!table) throw new Error(`fixture "${scenario}" rendered no #${TABLE_ID}`)
	if (!table.offsetWidth) throw new Error(`fixture "${scenario}" was not laid out — offsetWidth is 0`)
	return table
}

/**
 * The two width bases `pickColWidthBasis` chooses between, read off the live page exactly
 * the way `genTableToSlides` reads them: from the first row that has cells, per cell,
 * before any colspan expansion.
 *
 * Returned as data so the spec can assert the fixture still *discriminates* — if these two
 * ever agree, the emitted grid stops naming which arm produced it and the test would be
 * green for no reason. That check belongs in the spec, but the numbers can only be taken
 * here.
 *
 * @param {string} scenario a key of `TABLE_HTML`
 * @returns {{measured: number[], css: string[]}}
 */
function bases(scenario) {
	const table = render(scenario)
	const cells = [...table.rows[0].cells]
	return {
		measured: cells.map((cell) => cell.offsetWidth),
		css: cells.map((cell) => getComputedStyle(cell).getPropertyValue('width')),
	}
}

/**
 * Render a fixture and convert it, returning the package as base64.
 *
 * Errors are flattened rather than thrown for the same reason ./harness.mjs flattens them:
 * a rejection crossing `page.evaluate` arrives as a bare message, losing the class and its
 * `code`.
 *
 * @param {string} scenario a key of `TABLE_HTML`
 * @returns {Promise<{ok: true, base64: string} | {ok: false, code: string, message: string}>}
 */
async function build(scenario) {
	try {
		render(scenario)
		const pres = new TsPptx()
		// The instance method, not the `ts-pptx/html` free function: this is the call a
		// browser consumer makes, and it is the one that resolves the table id off the
		// ambient `document`.
		pres.tableToSlides(TABLE_ID)
		return { ok: true, base64: /** @type {string} */ (await pres.write({ outputType: 'base64' })) }
	} catch (err) {
		return { ok: false, code: String(err?.code ?? ''), message: String(err?.message ?? err) }
	}
}

Object.assign(window, { tableHarness: { bases, build } })
