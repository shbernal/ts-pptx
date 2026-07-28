import { Window } from 'happy-dom'
import { tableToSlides } from '../../dist/html.js'
import BrowserTsPptx from '../../dist/browser.js'
import { build, readEntry, listEntries, assert, assertEqual, defineRegressionSuite } from '../helpers.js'

// Acceptance: the `ts-pptx/html` subpath converts an HTML table to slides outside a browser.
// This is the case the whole portability effort exists for, and it is the one the pure-helper
// unit tests cannot reach: only a real DOM proves that `ownerDocument`/`defaultView` resolution,
// the scoped selectors, the degraded width basis, and the cell-style reads compose into a
// package. happy-dom is the DOM under test — it implements enough of the surface to drive the
// whole flow while (deliberately, see below) rendering nothing.
//
// The conversion is asserted through the emitted `slideN.xml`, the same way every other table
// suite in this directory asserts. `inspectPptx` does not model tables (its element kinds are
// text/image/shape/group), so it cannot see any of this.

const ONE_IN_EMU = 914400

/** A fresh window per test — no global DOM is installed, and no state leaks between cases. */
function windowWith(html) {
	const win = new Window()
	win.document.body.innerHTML = html
	return win
}

function tableOf(win, id = 't') {
	const table = win.document.getElementById(id)
	assert(table, `fixture is missing #${id}`)
	return table
}

/** Cell texts per row, in emitted order, from one slide's table. */
function cellTexts(xml) {
	return [...xml.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((row) =>
		[...row[0].matchAll(/<a:tc[\s\S]*?<\/a:tc>/g)].map((cell) =>
			[...cell[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((run) => run[1]).join('|')
		)
	)
}

function gridColWidths(xml) {
	return [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
}

function slideCount(zip) {
	return listEntries(zip).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length
}

const STYLED_TABLE = `
	<style>
		#t th { color: #ff0000; background-color: #00ff00; font-weight: 700; text-align: center; }
		#t td { color: #112233; background-color: #ffffff; text-align: right; }
	</style>
	<table id="t">
		<thead><tr><th>H1</th><th>H2</th></tr></thead>
		<tbody><tr><td>b1</td><td>b2</td></tr></tbody>
		<tfoot><tr><td>f1</td><td>f2</td></tr></tfoot>
	</table>`

defineRegressionSuite('HTML table to slides on Node (happy-dom)', [
	{
		name: 'thead / tbody / tfoot rows all land, in that order',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const rows = cellTexts(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(
				JSON.stringify(rows),
				JSON.stringify([
					['H1', 'H2'],
					['b1', 'b2'],
					['f1', 'f2'],
				]),
				'row order'
			)
		},
	},
	{
		name: 'columns split evenly when nothing laid the table out and no override is set',
		fn: async () => {
			const win = windowWith('<table id="t"><thead><tr><th>A</th><th>B</th><th>C</th></tr></thead></table>')
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const cols = gridColWidths(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(cols.length, 3, 'column count')
			// The headline Node case: `offsetWidth` is 0 for every cell here, which used to make
			// the proportional calc a 0/0 divide and emit three zero-width columns.
			cols.forEach((w) => assert(w > ONE_IN_EMU, `expected each column wider than 1in; got ${w} EMU`))
			assert(cols[0] === cols[1] && cols[1] === cols[2], `expected an equal split; got ${cols.join(', ')}`)
		},
	},
	{
		name: 'computed CSS widths drive the proportions when they are stated for every column',
		fn: async () => {
			const win = windowWith(`
				<style>#t th:nth-child(1){width:100px} #t th:nth-child(2){width:300px}</style>
				<table id="t"><thead><tr><th>A</th><th>B</th></tr></thead></table>`)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const cols = gridColWidths(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(cols.length, 2, 'column count')
			const ratio = cols[1] / cols[0]
			assert(Math.abs(ratio - 3) < 0.02, `expected a 1:3 split from the CSS widths; got ratio ${ratio}`)
		},
	},
	{
		name: 'data-pptx-width overrides the degraded basis exactly',
		fn: async () => {
			const win = windowWith(
				'<table id="t"><thead><tr><th data-pptx-width="2">A</th><th data-pptx-width="4">B</th></tr></thead></table>'
			)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const cols = gridColWidths(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(cols.length, 2, 'column count')
			assertEqual(cols[0], 2 * ONE_IN_EMU, 'first column honors data-pptx-width')
			assertEqual(cols[1], 4 * ONE_IN_EMU, 'second column honors data-pptx-width')
		},
	},
	{
		name: 'data-pptx-min-width still acts as a floor',
		fn: async () => {
			const win = windowWith(
				'<table id="t"><thead><tr><th data-pptx-min-width="8">A</th><th>B</th></tr></thead></table>'
			)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const cols = gridColWidths(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(cols[0], 8 * ONE_IN_EMU, 'the floor must raise the equal-split width')
			assert(cols[1] < cols[0], `unfloored column should stay at its equal split; got ${cols[1]}`)
		},
	},
	{
		name: 'colspan and rowspan survive into the emitted table',
		fn: async () => {
			const win = windowWith(`
				<table id="t">
					<thead><tr><th colspan="2">Wide</th><th>C</th></tr></thead>
					<tbody><tr><td rowspan="2">Tall</td><td>b</td><td>c</td></tr><tr><td>d</td><td>e</td></tr></tbody>
				</table>`)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:tc\b[^>]*\bgridSpan="2"/.test(xml), 'colspan must emit gridSpan="2"')
			assert(/<a:tc\b[^>]*\browSpan="2"/.test(xml), 'rowspan must emit rowSpan="2"')
			// A colspan of 2 in the width-source row must still yield one column per spanned cell.
			assertEqual(gridColWidths(xml).length, 3, 'a 2-span header + 1 cell is 3 columns')
		},
	},
	{
		name: 'a <br> inside a cell survives as a line break',
		fn: async () => {
			const win = windowWith('<table id="t"><tbody><tr><td>Line 1<br>Line 2</td></tr></tbody></table>')
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// A line break is a paragraph split, so the one cell must carry two <a:p>. happy-dom
			// implements `innerText` but not as *rendered* text — it returns "Line 1Line 2"
			// here — so this also guards the walk that overrides it. (Runs inside a paragraph
			// are split further by the text serializer; paragraph count is the line count.)
			const paragraphs = (xml.match(/<a:p>/g) || []).length
			assertEqual(paragraphs, 2, 'a <br> must split the cell into two paragraphs')
			assert(!/Line 1Line 2/.test(xml), 'the two lines must not be concatenated')
			assert(xml.includes('>1</a:t>') && xml.includes('>2</a:t>'), `both line texts must survive; got: ${xml}`)
		},
	},
	{
		name: 'computed colors, weight and alignment map through',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// happy-dom returns the authored `#rrggbb` rather than a browser's `rgb(...)`;
			// parsing it as rgb() used to emit the literal color "NANNANNAN".
			assert(!/NAN/i.test(xml), `no channel may fail to parse; got: ${xml}`)
			assert(xml.includes('FF0000'), 'header text color must reach the XML')
			assert(xml.includes('00FF00'), 'header fill color must reach the XML')
			assert(xml.includes('112233'), 'body text color must reach the XML')
			assert(/<a:rPr\b[^>]*\bb="1"/.test(xml), 'font-weight 700 must emit bold')
			assert(/algn="ctr"/.test(xml), 'header text-align:center must emit algn="ctr"')
			assert(/algn="r"/.test(xml), 'body text-align:right must emit algn="r"')
		},
	},
	{
		name: 'computed borders map through, including a hex border color',
		fn: async () => {
			const win = windowWith(`
				<style>#t td { border-top: 2px solid #663399; border-right-width: 1px; border-right-color: #112233; }</style>
				<table id="t"><tbody><tr><td>x</td></tr></tbody></table>`)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:lnT w="25400"[\s\S]*?663399/.test(xml), `expected a 2pt #663399 top border; got: ${xml}`)
			assert(/<a:lnR w="12700"[\s\S]*?112233/.test(xml), `expected a 1pt #112233 right border; got: ${xml}`)
			// Unstated sides compute to '' and must stay absent rather than becoming a black hairline.
			assert(/<a:lnB w="0"/.test(xml), `an unstated border must emit no line; got: ${xml}`)
		},
	},
	{
		name: 'vertical-align and the first font-family map through',
		fn: async () => {
			const win = windowWith(`
				<style>#t td { vertical-align: middle; font-family: Arial, sans-serif; }</style>
				<table id="t"><tbody><tr><td>x</td></tr></tbody></table>`)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/anchor="ctr"/.test(xml), `vertical-align:middle must emit anchor="ctr"; got: ${xml}`)
			assert(/typeface="Arial"/.test(xml), `only the first font-family entry is used; got: ${xml}`)
		},
	},
	{
		name: 'fractional computed padding rounds to 2, not to 15',
		fn: async () => {
			const win = windowWith(`
				<style>#t td { padding: 1.5px 2px 3px 4px; }</style>
				<table id="t"><tbody><tr><td>x</td></tr></tbody></table>`)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// Cell margin is inches downstream (914400 EMU each), so these are the *numbers* the
			// padding parse produced, unit conversion aside. `1.5px` must become 2: the historical
			// `.replace(/\D/g,'')` deleted the decimal point and made it 15 (=> marT 13716000).
			// NOTE the px -> inches unit mismatch here is a separate pre-existing defect, tracked
			// as dn-html-table-padding-units; this case pins the parse, not the unit.
			assert(/marT="1828800"/.test(xml), `1.5px padding must round to 2, not 15; got: ${xml}`)
			assert(!/marT="13716000"/.test(xml), 'the stripped-decimal-point value must not come back')
			assert(/marL="3657600"/.test(xml), `4px padding must be 4; got: ${xml}`)
		},
	},
	{
		name: 'an explicit slideMargin array positions the table',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win), { slideMargin: [0.25, 0.75, 0.25, 1.5] })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const off = /<p:xfrm><a:off x="(\d+)" y="(\d+)"\/>/.exec(xml)
			assert(off, `expected a positioned graphic frame; got: ${xml}`)
			assertEqual(Number(off[1]), Math.round(1.5 * ONE_IN_EMU), 'x comes from the left margin')
			assertEqual(Number(off[2]), Math.round(0.25 * ONE_IN_EMU), 'y comes from the top margin')
		},
	},
	{
		name: 'auto-page extras (text, shape, table) land on the generated slide',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win), {
					addText: { text: [{ text: 'Appendix' }], options: { x: 1, y: 6, w: 3, h: 0.4 } },
					addShape: { shapeName: 'rect', options: { x: 0, y: 0, w: 1, h: 1 } },
					addTable: { rows: [['extra']], options: { x: 1, y: 5, w: 3 } },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('Appendix'), 'addText must reach the auto-paged slide')
			assert(xml.includes('extra'), 'addTable must reach the auto-paged slide')
			assert(/prstGeom prst="rect"/.test(xml), `addShape must reach the auto-paged slide; got: ${xml}`)
		},
	},
	{
		name: 'addImage without path or data warns instead of emitting a broken image',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win), { addImage: { image: {}, options: { x: 1, y: 1, w: 1, h: 1 } } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!xml.includes('<p:pic>'), 'an image with neither path nor data must not be emitted')
		},
	},
	{
		name: 'verbose mode logs without changing the emitted table',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const logged = []
			const realLog = console.log
			console.log = (...args) => logged.push(args.join(' '))
			let quiet
			let loud
			try {
				quiet = await build((pptx) => {
					tableToSlides(pptx, tableOf(windowWith(STYLED_TABLE)))
				})
				loud = await build((pptx) => {
					tableToSlides(pptx, tableOf(win), { verbose: true })
				})
			} finally {
				console.log = realLog
			}
			assert(
				logged.some((line) => line.includes('tableToSlides')),
				`verbose must trace the layout process; got ${logged.length} lines`
			)
			assertEqual(
				await readEntry(loud.zip, 'ppt/slides/slide1.xml'),
				await readEntry(quiet.zip, 'ppt/slides/slide1.xml'),
				'verbose is a dev-only flag and must not change output'
			)
		},
	},
	{
		name: 'a tall table auto-pages onto more than one slide',
		fn: async () => {
			const rows = Array.from({ length: 60 }, (_, i) => `<tr><td>row ${i}</td><td>value ${i}</td></tr>`).join('')
			const win = windowWith(`<table id="t"><thead><tr><th>K</th><th>V</th></tr></thead><tbody>${rows}</tbody></table>`)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, tableOf(win))
			})
			assert(slideCount(zip) > 1, `expected auto-paging to produce >1 slide; got ${slideCount(zip)}`)
		},
	},
	{
		name: 'the string-id form resolves against an explicitly supplied document',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, 't', { document: win.document })
			})
			const rows = cellTexts(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(rows.length, 3, 'the id must resolve to the same three rows')
		},
	},
	{
		name: 'an id that is not a valid CSS identifier still resolves',
		fn: async () => {
			// The id is interpolated into no selector now, so a leading digit is fine. It used to
			// pass the getElementById reality-check and then match nothing, emitting an empty table.
			const win = windowWith('<table id="2024.report"><tbody><tr><td>cell</td></tr></tbody></table>')
			const { zip } = await build((pptx) => {
				tableToSlides(pptx, '2024.report', { document: win.document })
			})
			const rows = cellTexts(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(JSON.stringify(rows), JSON.stringify([['cell']]), 'the table must not be empty')
		},
	},
	{
		name: 'the string-id form names both remedies when no document is resolvable',
		fn: async () => {
			let thrown
			try {
				await build((pptx) => {
					tableToSlides(pptx, 'nowhere')
				})
			} catch (err) {
				thrown = err
			}
			assert(thrown, 'a bare id with no DOM anywhere must throw, not silently emit nothing')
			assert(
				/options\.document/.test(thrown.message) && /element/.test(thrown.message),
				`the error must name both remedies; got: ${thrown.message}`
			)
		},
	},
	{
		name: 'a missing table id still reports the historical message',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			let thrown
			try {
				await build((pptx) => {
					tableToSlides(pptx, 'absent', { document: win.document })
				})
			} catch (err) {
				thrown = err
			}
			assert(thrown, 'an unresolvable id must throw')
			assertEqual(thrown.message, 'tableToSlides: Table ID "absent" does not exist!', 'error message')
		},
	},
	{
		name: 'masterTitle picks up the named master, on the free function too',
		fn: async () => {
			const win = windowWith(STYLED_TABLE)
			const { zip } = await build((pptx) => {
				pptx.defineSlideMaster({ title: 'REPORT', margin: 1.25 })
				tableToSlides(pptx, tableOf(win), { masterTitle: 'REPORT' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// The master's 1.25in margin becomes the table's x/y origin. Reading it back proves
			// the lookup happened: without it the default 0.5in margin would be used. Anchor on
			// the graphic frame's own <p:xfrm> — the slide's group transform is an <a:off> too.
			const off = /<p:xfrm><a:off x="(\d+)" y="(\d+)"\/>/.exec(xml)
			assert(off, `expected a positioned graphic frame; got: ${xml}`)
			assertEqual(Number(off[1]), Math.round(1.25 * ONE_IN_EMU), 'x must come from the master margin')
			assertEqual(Number(off[2]), Math.round(1.25 * ONE_IN_EMU), 'y must come from the master margin')
		},
	},
	{
		name: 'the browser method delegates to the same implementation',
		fn: async () => {
			const html = STYLED_TABLE
			// The method takes an id and no element, so it can only be driven here by handing it
			// the document explicitly — which is exactly the delegation being asserted.
			const viaMethod = await build((pptx) => {
				BrowserTsPptx.prototype.tableToSlides.call(pptx, 't', { document: windowWith(html).document })
			})
			const viaFunction = await build((pptx) => {
				tableToSlides(pptx, tableOf(windowWith(html)))
			})
			assertEqual(typeof BrowserTsPptx.prototype.tableToSlides, 'function', 'the browser build must keep the method')
			assertEqual(
				await readEntry(viaMethod.zip, 'ppt/slides/slide1.xml'),
				await readEntry(viaFunction.zip, 'ppt/slides/slide1.xml'),
				'method and free function must emit the same slide'
			)
		},
	},
])
