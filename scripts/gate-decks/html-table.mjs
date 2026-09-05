/**
 * Gate deck: the HTML-table conversion path.
 *
 * Corpus for `scripts/byte-identity.mjs`. No showcase deck and no other gate deck calls
 * `tableToSlides`, so `src/gen/table/html-dom.ts` — the whole computed-CSS-to-OOXML walk —
 * was outside the diff. A `check` that moved every HTML-table border width by a third
 * reported PASS, which is precisely the "unproven, not proven unchanged" case the gate decks
 * exist for. See `./README.md` for why a gate deck is a separate thing from a showcase.
 *
 * The DOM is happy-dom, the same implementation the Node regression suite drives
 * (`test/regression/html/html-to-slides-node.test.js`). It is a devDependency, which is all
 * this needs: the gate never ships and never runs from an installed package.
 *
 * **Nothing here lays the table out.** happy-dom resolves the cascade and computes styles but
 * runs no layout engine, so every `offsetWidth` is `0` and the column widths come from the
 * fallback bases rather than from measurement. That is not a limitation of the fixture — it is
 * the Node path this deck is here to freeze, and it is deterministic in a way a real browser's
 * font metrics would not be. The measured basis has its own coverage in Chromium
 * (`test/browser/table-widths.spec.mjs`), where byte-identity cannot follow.
 *
 * One table per slide. What each is here to reach:
 *
 *   styled      the per-cell style reads: computed `color`, `background-color`, `font-weight`,
 *               `font-style`, `font-size`, `text-align` and `vertical-align`, over
 *               thead/tbody/tfoot so the three section walks all run.
 *   borders     `htmlBorderToProps` on all four edges at three widths, three styles and two
 *               colours, plus the `padding` reads that share the same declaration block.
 *   spans       `colspan`/`rowspan` through the occupancy grid, including a span that opens on
 *               one row and closes two below it, which is the only way the row-span bookkeeping
 *               is exercised.
 *   widths      the two fallback bases and the two overrides: an equal split where the
 *               stylesheet states nothing, computed CSS widths where it states every column,
 *               `data-pptx-width` (exact) and `data-pptx-min-width` (a floor), with one of each
 *               on a spanning cell so the divide-across-covered-columns arm runs.
 *   paged       40 rows through the auto-pager with a repeated header, a `masterTitle` and a
 *               `slideMargin`, so the conversion's own paging options reach `getSlidesForTableRows`
 *               rather than only the `addTable` ones the showcases already cover.
 *   breaks      `<br>` as a line break, entities, and text a cell carries alongside inline
 *               markup — the text extraction, which is the one part with no styling in it.
 *
 * Every case carries its caption through the conversion's own `addText` option rather than
 * putting it on a slide of its own, because that is the only way the three companion arms
 * (`addText`, `addShape`, `addImage`) reach a slide at all: they are applied to the slides the
 * conversion creates, and nothing else in the corpus creates one. One case carries a shape and
 * a picture too, which is what leaves that whole block diffed.
 */
import { Window } from 'happy-dom'
import TsPptx, { ShapeType } from '../../dist/node.js'
import { tableToSlides } from '../../dist/html.js'

const TITLE = { x: 0.3, y: 0.15, w: 12.7, h: 0.4, fontSize: 14, bold: true }
/** Where each converted table sits. `tableToSlides` takes the same position options `addTable` does. */
const PLACE = { x: 0.3, y: 0.8, w: 12.7 }
/** A 1x1 transparent PNG, inline so this deck reads no asset off disk. */
const PNG_1PX =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

/**
 * A fresh document per table: no global DOM is installed, and no stylesheet leaks between cases.
 * @param {string} html - the fixture markup, including its `<style>`
 * @returns {import('../../dist/html.js').TableToSlidesElement} the `#t` table element
 */
function tableOf(html) {
	const win = new Window()
	win.document.body.innerHTML = html
	const table = win.document.getElementById('t')
	if (!table) throw new Error('fixture is missing #t')
	return /** @type {import('../../dist/html.js').TableToSlidesElement} */ (/** @type {unknown} */ (table))
}

const STYLED = `
	<style>
		#t th { color: #ff0000; background-color: #00ff00; font-weight: 700; text-align: center; vertical-align: top; }
		#t td { color: #112233; background-color: #ffffff; text-align: right; vertical-align: middle; font-size: 11px; }
		#t tfoot td { font-style: italic; font-weight: 900; text-align: left; vertical-align: bottom; }
	</style>
	<table id="t">
		<thead><tr><th>Region</th><th>Q1</th><th>Q2</th></tr></thead>
		<tbody>
			<tr><td>North</td><td>1,204</td><td>1,530</td></tr>
			<tr><td>South</td><td>982</td><td>1,107</td></tr>
		</tbody>
		<tfoot><tr><td>Total</td><td>2,186</td><td>2,637</td></tr></tfoot>
	</table>`

const BORDERS = `
	<style>
		#t td { padding: 4px 12px 9px 2px; }
		#t .a { border: 1px solid #000000; }
		#t .b { border-top: 3px dashed #ff8800; border-right: 2px dotted #0088ff; border-bottom: 1px solid #008800; border-left: 4px double #880088; }
		#t .c { border: 0; padding: 0; }
		#t .d { border: 6px solid #cccccc; padding: 18px; }
	</style>
	<table id="t">
		<tbody>
			<tr><td class="a">all four, one solid rule</td><td class="b">four different edges</td></tr>
			<tr><td class="c">no border, no padding</td><td class="d">thick, generously padded</td></tr>
		</tbody>
	</table>`

const SPANS = `
	<table id="t">
		<thead><tr><th colspan="3">Spanning header</th><th>Plain</th></tr></thead>
		<tbody>
			<tr><td rowspan="3">Three rows tall</td><td>b</td><td colspan="2">two wide</td></tr>
			<tr><td colspan="2">two wide</td><td>d</td></tr>
			<tr><td>e</td><td>f</td><td>g</td></tr>
			<tr><td>h</td><td>i</td><td>j</td><td>k</td></tr>
		</tbody>
	</table>`

const WIDTHS_CSS = `
	<style>#t th:nth-child(1){width:100px} #t th:nth-child(2){width:300px} #t th:nth-child(3){width:200px}</style>
	<table id="t">
		<thead><tr><th>narrow</th><th>widest</th><th>middle</th></tr></thead>
		<tbody><tr><td>a</td><td>b</td><td>c</td></tr></tbody>
	</table>`

const WIDTHS_EVEN = `
	<table id="t">
		<thead><tr><th>A</th><th>B</th><th>C</th><th>D</th></tr></thead>
		<tbody><tr><td>a</td><td>b</td><td>c</td><td>d</td></tr></tbody>
	</table>`

const WIDTHS_ATTR = `
	<table id="t">
		<thead>
			<tr>
				<th data-pptx-width="2">exactly 2in</th>
				<th data-pptx-width="4" colspan="2">4in across two columns</th>
				<th data-pptx-min-width="3">at least 3in</th>
			</tr>
		</thead>
		<tbody><tr><td>a</td><td>b</td><td>c</td><td>d</td></tr></tbody>
	</table>`

const BREAKS = `
	<table id="t">
		<tbody>
			<tr><td>first line<br>second line<br>third</td><td>R&amp;D &lt;core&gt; &quot;q&quot;</td></tr>
			<tr><td>text <b>bold</b> tail</td><td>   leading and trailing   </td></tr>
		</tbody>
	</table>`

/** 40 body rows under one header, long enough that the pager needs several slides. */
const PAGED = `
	<style>#t th { background-color: #203040; color: #ffffff; font-weight: 700; }</style>
	<table id="t">
		<thead><tr><th>Account</th><th>Note</th></tr></thead>
		<tbody>${Array.from(
			{ length: 40 },
			(_unused, idx) =>
				`<tr><td>Account ${idx + 1}</td><td>lorem ipsum dolor sit amet consectetur adipiscing elit sed do</td></tr>`
		).join('')}</tbody>
	</table>`

/** Each converted table: the caption its slide carries, the markup, and the conversion options. */
const CASES = [
	{ name: 'styled: computed colour, weight, size and alignment', html: STYLED, opts: {} },
	{ name: 'borders: four edges, three styles, and padding', html: BORDERS, opts: {} },
	{ name: 'spans: colspan, rowspan, and the occupancy grid', html: SPANS, opts: {} },
	{ name: 'widths: computed CSS proportions', html: WIDTHS_CSS, opts: {} },
	{ name: 'widths: equal split, nothing stated', html: WIDTHS_EVEN, opts: {} },
	{ name: 'widths: data-pptx-width and data-pptx-min-width', html: WIDTHS_ATTR, opts: {} },
	{
		name: 'text: <br>, entities, inline markup (plus the companion objects)',
		html: BREAKS,
		opts: {
			fontSize: 13,
			addShape: {
				shapeName: ShapeType.roundRect,
				options: { x: 11.5, y: 6.4, w: 1.2, h: 0.5, fill: { color: 'DDE5EE' } },
			},
			addImage: { image: { data: PNG_1PX }, options: { x: 0.3, y: 6.4, w: 0.5, h: 0.5 } },
		},
	},
]

async function compose() {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_WIDE'
	pptx.author = 'ts-pptx byte-identity gate'
	pptx.title = 'HTML table conversion matrix'
	// The paged case names this master, which is the only way `tableToSlides`'s own
	// `masterTitle` reaches the slides it creates.
	pptx.defineSlideMaster({ title: 'GATE', margin: 0.4, background: { color: 'F4F6F8' } })

	for (const kase of CASES) {
		tableToSlides(pptx, tableOf(kase.html), {
			...PLACE,
			autoPage: false,
			addText: { text: [{ text: kase.name }], options: TITLE },
			...kase.opts,
		})
	}

	// Paged last: it appends slides of its own, so a diff reads the fixed cases above at fixed
	// slide numbers however many pages this one takes.
	tableToSlides(pptx, tableOf(PAGED), {
		x: 0.5,
		y: 0.5,
		w: 12,
		fontSize: 14,
		autoPage: true,
		autoPageRepeatHeader: true,
		autoPageSlideStartY: 0.5,
		masterTitle: 'GATE',
		slideMargin: 0.4,
	})
	return pptx
}

/** @param {string} outFile @returns {Promise<string>} */
export async function build(outFile) {
	const pptx = await compose()
	return await pptx.writeFile({ fileName: outFile })
}

export const gateDeck = {
	slug: 'html-table',
	title: 'HTML table conversion matrix',
	description:
		'Byte-identity corpus for src/gen/table/html-dom.ts — the computed-CSS-to-OOXML walk no showcase reaches.',
	fileName: 'gate_html_table.pptx',
	build,
}
