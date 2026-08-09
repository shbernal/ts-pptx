// Table markup shared by the rendered-table harness page and the Node side of the same
// comparison, in the same spirit as ./decks.mjs: defined once, rendered twice.
//
// What these fixtures are for is narrow and worth stating precisely, because the
// neighbouring claim is out of scope. `pickColWidthBasis` (src/gen/table/html-dom.ts) picks
// between three width bases, and its *first* arm — the rendered `offsetWidth` — had never
// executed anywhere: the Node suite drives happy-dom, where `offsetWidth` is 0 for every
// cell, and the unit tests reach `pickColWidthBasis` by handing it numbers directly. So the
// primary path of the whole feature was proven only at its own function boundary.
//
// These fixtures run it end to end in a real layout engine. They assert that the measured
// arm executes and that the emitted grid is proportional to what was measured. They do NOT
// assert that the measurement matches what a browser painted, or that two engines agree —
// that is live-DOM layout fidelity, and it stays out of scope (docs/project-target.md).
//
// ── The discriminator ──────────────────────────────────────────────────────────────────
//
// A test that only proves "widths came out proportional" would pass just as green if the
// measured arm never ran, because the CSS arm would have produced the same answer. So the
// `measured` fixture is built so the two bases *disagree*:
//
//   - `offsetWidth` is the border box — content + padding + border.
//   - computed `width` is the content box (`box-sizing: content-box`).
//
// Column B therefore measures the same 320px as column A while stating a computed width of
// 160px. Measured basis 1:1, CSS basis 2:1. Only one of those can be the answer, so the
// emitted grid names which arm ran. The spec re-derives both bases from the live page and
// fails if they ever converge, rather than trusting this comment to stay true.
//
// The padding stays under 96px deliberately. Cell padding is converted to inches for
// `TableCellProps.margin`, and a margin >= 1in trips the legacy-points diagnostic — a
// fixture should not have to emit a warning to make its point.

/** The id every fixture gives its table, so one selector serves them all. */
export const TABLE_ID = 'tbl'

/** How many body rows the `autoPage` fixture below carries. Named so the specs can assert on it. */
export const AUTOPAGE_ROWS = 60

/**
 * The `autoPage` fixture's markup, built rather than written out.
 *
 * Sixty near-identical rows is not something to paste into a template literal, and the row count
 * has to be a number the spec can also reason about — it is the total the pages must add back up
 * to. See the `autoPage` entry in {@link TABLE_HTML} for what the fixture is for.
 * @returns {string} the table markup
 */
function autoPageMarkup() {
	const cell = 'padding:8px;font-size:16px;border:0'
	const rows = []
	for (let idx = 1; idx <= AUTOPAGE_ROWS; idx++) {
		rows.push(`<tr><td style="${cell}">R${idx}</td><td style="${cell}">V${idx}</td></tr>`)
	}
	return `
		<table id="${TABLE_ID}" style="table-layout:fixed;width:900px;border-collapse:separate;border-spacing:0;font-size:16px">
			<thead><tr>
				<th style="${cell};width:300px">Key</th>
				<th style="${cell};width:600px">Value</th>
			</tr></thead>
			<tbody>${rows.join('')}</tbody>
		</table>`
}

/**
 * Fixture markup, keyed by the name the harness and the specs dispatch on.
 *
 * Widths are inline `style` attributes rather than a `<style>` block on purpose: this
 * markup is parsed by two different DOMs, and an inline declaration is the one form both
 * resolve through `getComputedStyle` without a cascade to run.
 *
 * @type {Record<string, string>}
 */
export const TABLE_HTML = {
	/**
	 * The headline case: measured and CSS bases disagree, so the emitted grid says which
	 * one drove it. See "The discriminator" above for the arithmetic.
	 *
	 * `table-layout: fixed` is what makes the first row authoritative for column widths,
	 * which is also the row `tableToSlides` takes its basis from — so the fixture and the
	 * implementation are reading the same thing.
	 */
	measured: `
		<table id="${TABLE_ID}" style="table-layout:fixed;width:640px;border-collapse:separate;border-spacing:0">
			<thead><tr>
				<th style="box-sizing:content-box;width:320px;padding:0;border:0">A</th>
				<th style="box-sizing:content-box;width:160px;padding:0 80px;border:0">B</th>
			</tr></thead>
			<tbody><tr>
				<td style="box-sizing:content-box;padding:0;border:0">a1</td>
				<td style="box-sizing:content-box;padding:0 80px;border:0">b1</td>
			</tr></tbody>
		</table>`,

	/**
	 * `data-pptx-width` against a *live* measurement. The Node suite already proves the
	 * override beats the degraded basis, which is the easy half — there it is beating a
	 * vector of zeroes. Here it has a real 1:3 measurement to beat.
	 */
	override: `
		<table id="${TABLE_ID}" style="table-layout:fixed;width:600px;border-collapse:separate;border-spacing:0">
			<thead><tr>
				<th style="box-sizing:content-box;width:150px;padding:0;border:0" data-pptx-width="4">A</th>
				<th style="box-sizing:content-box;width:450px;padding:0;border:0">B</th>
			</tr></thead>
			<tbody><tr>
				<td style="padding:0;border:0">a</td>
				<td style="padding:0;border:0">b</td>
			</tr></tbody>
		</table>`,

	/**
	 * upstream gitbrent/PptxGenJS#1244, under measurement.
	 *
	 * A `data-pptx-width` on a spanning cell must divide across the columns that cell
	 * covers. The Node suite pins this on a table nothing laid out; the arithmetic that
	 * fixed it (`arrColSrc`, src/gen/table/html-dom.ts) is built in the same pass that
	 * reads `offsetWidth`, so the combination a real consumer has — a span, an override and
	 * a live measurement at once — is only reached here.
	 */
	spanOverride: `
		<table id="${TABLE_ID}" style="table-layout:fixed;width:600px;border-collapse:separate;border-spacing:0">
			<thead><tr>
				<th colspan="2" style="box-sizing:content-box;width:400px;padding:0;border:0" data-pptx-width="4">Wide</th>
				<th style="box-sizing:content-box;width:200px;padding:0;border:0" data-pptx-width="2">C</th>
			</tr></thead>
			<tbody><tr>
				<td style="padding:0;border:0">a</td>
				<td style="padding:0;border:0">b</td>
				<td style="padding:0;border:0">c</td>
			</tr></tbody>
		</table>`,

	/**
	 * upstream gitbrent/PptxGenJS#1200 — "tableToSlides autoPaging not working", the repro that
	 * report invited and this repo could not build until the browser lane existed.
	 *
	 * The reporter's shape, as plainly as it can be stated: an ordinary table of dynamic data,
	 * too tall for one slide, converted with **no options at all**. What came back paged onto
	 * several slides with rows hanging off the bottom edge of the generated ones.
	 *
	 * Deliberately ordinary. Nothing here pushes on a corner: no `data-pptx-*` hints, no colspans,
	 * padding and font-size a stylesheet would plausibly set, cells short enough that no column
	 * width can make them wrap. That last one is what gives the spec its oracle — every row is the
	 * same height, so every page with the same usable height must hold the same number of them,
	 * and no rendered-page measurement is needed to say a page took one too many. See
	 * test/browser/table-autopage.spec.mjs.
	 */
	autoPage: autoPageMarkup(),
}
