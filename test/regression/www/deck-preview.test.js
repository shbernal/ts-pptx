import { Window } from 'happy-dom'
import { describe, expect, it } from 'vitest'
import { counted, slideList, splitDeck, summarizeNotes } from '../../../www/demos/deck-preview.ts'

/**
 * The demos page's pure helpers.
 *
 * They exist so the component is markup plus assignments — `.vue` files are read by no
 * typechecker and by no test runner here — and this file is the other half of that
 * bargain. Nothing below touches a DOM, a browser or a deck: what is asserted is the
 * grouping and the prose, which are the two things that would go wrong quietly.
 */

/** One fidelity row, with only the fields the helpers read. */
function note(slide, construct, disposition = 'dropped', cause = 'unread', detail = 'because') {
	return { slide, construct, disposition, cause, detail }
}

describe('summarizeNotes', () => {
	it('collapses one construct raised on many slides into a single row', () => {
		// The real shape: a footer field note fires on every slide carrying a footer, and
		// eleven copies of one sentence reads as eleven problems rather than one.
		const rows = summarizeNotes([note(2, 'text.field'), note(4, 'text.field'), note(5, 'text.field')])

		expect(rows).toHaveLength(1)
		expect(rows[0].construct).toBe('text.field')
		expect(rows[0].slides).toEqual([2, 4, 5])
	})

	it('keeps constructs apart when they differ in disposition or cause', () => {
		const rows = summarizeNotes([
			note(1, 'chart.workbook', 'approximated', 'unsupported'),
			note(1, 'chart.workbook', 'dropped', 'unsupported'),
			note(1, 'chart.workbook', 'dropped', 'unread'),
		])

		// Same construct three times, three different declarations about it. Grouping on the
		// construct alone would report one, and would have to pick one of the three details.
		expect(rows).toHaveLength(3)
		expect(rows.map((row) => `${row.disposition}/${row.cause}`)).toEqual([
			'approximated/unsupported',
			'dropped/unsupported',
			'dropped/unread',
		])
	})

	it('does not repeat a slide that raised the same note twice', () => {
		const rows = summarizeNotes([note(3, 'shape.placeholder'), note(3, 'shape.placeholder')])

		expect(rows[0].slides).toEqual([3])
	})

	it('returns nothing for a deck that declared nothing', () => {
		expect(summarizeNotes([])).toEqual([])
	})
})

describe('slideList', () => {
	it('reads as prose rather than as an array', () => {
		expect(slideList([3, 5, 7])).toBe('3, 5 and 7')
		expect(slideList([2, 4])).toBe('2 and 4')
		expect(slideList([9])).toBe('9')
	})

	it('sorts numerically, not lexically', () => {
		// `[2, 10].sort()` is `[10, 2]`, and a deck of eleven slides is exactly where that
		// shows up.
		expect(slideList([10, 2, 11])).toBe('2, 10 and 11')
	})

	it('is empty for no slides, rather than throwing or saying "undefined"', () => {
		expect(slideList([])).toBe('')
	})

	it('reads three or more consecutive slides as a range', () => {
		// The footer notes fire on every content slide, and "2, 4, 5, 6, 8, 9 and 10" hides
		// the one gap a reader is looking for.
		expect(slideList([2, 4, 5, 6, 8, 9, 10])).toBe('2, 4 to 6 and 8 to 10')
		expect(slideList([1, 2, 3])).toBe('1 to 3')
		expect(slideList([5, 3, 4, 4])).toBe('3 to 5')
	})
})

describe('counted', () => {
	it('agrees the noun with the count', () => {
		expect(counted(1, 'difference')).toBe('1 difference')
		expect(counted(4, 'difference')).toBe('4 differences')
		expect(counted(0, 'warning')).toBe('0 warnings')
	})
})

/** A rendered deck document in the renderer's shape, cut down to what `splitDeck` reads. */
function renderedDeck() {
	const slide = (number, body) =>
		`<section class="pxh-slide" data-pxh-slide="${number}">` +
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12192000 6858000">${body}</svg></section>`
	const html = `<!doctype html><html><head><style>.pxh-text { font-family: Calibri }</style></head><body>
		<div class="pxh-deck">
			${slide(
				1,
				'<g data-pxh-node="s1.sp2" transform="translate(100 200)">' +
					'<g data-pxh-node="s1.sp3" transform="translate(100 200)"><path d="M 0 0"/></g></g>' +
					'<g data-pxh-node="s1.sp4" transform="translate(300 400)">' +
					'<g data-pxh-node="s1.sp4.r0c0" transform="translate(0 0)"><path d="M 0 0"/></g></g>' +
					'<foreignObject><div xmlns="http://www.w3.org/1999/xhtml"><p style="margin:0;line-height:1.2">' +
					'<span contenteditable="true">Title</span></p></div></foreignObject>' +
					'<image data-pxh-asset="logo.png"/>'
			)}
			<aside class="pxh-aside"><h2>Notes</h2><p>First thought.</p><p>Second thought.</p>
				<h2>Declared differences</h2><ul><li><code>text.field</code></li></ul></aside>
			${slide(2, '<rect/>')}
		</div>
		<script type="application/json" id="pxh-ir">{"assets":[{"name":"logo.png","contentType":"image/png"}]}</script>
		<script type="application/json" id="pxh-assets">{"logo.png":"iVBORw0K"}</script>
	</body></html>`
	const window = new Window()
	window.document.write(html)
	// happy-dom's `Document` is not the DOM lib's, though it has every member `splitDeck` reads.
	return /** @type {Document} */ (/** @type {unknown} */ (window.document))
}

describe('splitDeck', () => {
	it("cuts one piece per slide, with that slide's speaker notes and nothing else", () => {
		const { slides, styles, aspectRatio } = splitDeck(renderedDeck())

		expect(slides.map((slide) => slide.number)).toEqual([1, 2])
		// The declared differences sit in the same aside, as a list; only the paragraphs
		// are notes.
		expect(slides[0].notes).toEqual(['First thought.', 'Second thought.'])
		expect(slides[1].notes).toEqual([])
		expect(styles).toContain('.pxh-text')
		expect(aspectRatio).toBeCloseTo(16 / 9, 3)
	})

	it('leaves every transform where the renderer put it', () => {
		// The page used to strip a group's transform, because `pptx-html` 0.2.0 offset a
		// group's children twice. That is fixed in 0.2.1 and the correction is gone, so
		// what the renderer drew is what arrives -- including a table's, which was never
		// the bug and would have broken had the fix been written structurally.
		const [first] = splitDeck(renderedDeck()).slides
		const doc = new Window().document
		doc.body.innerHTML = first.markup

		expect(doc.querySelector('[data-pxh-node="s1.sp2"]').getAttribute('transform')).toBe('translate(100 200)')
		expect(doc.querySelector('[data-pxh-node="s1.sp4"]').getAttribute('transform')).toBe('translate(300 400)')
	})

	it('turns editing off and inlines pictures, and rewrites nothing else', () => {
		const [first] = splitDeck(renderedDeck()).slides
		const doc = new Window().document
		doc.body.innerHTML = first.markup

		// The paragraph's own line height is the renderer's to state; 0.2.1 writes the
		// unitless multiple PowerPoint means, and the page passes it through.
		expect(doc.querySelector('p').getAttribute('style')).toBe('margin:0;line-height:1.2')
		expect(doc.querySelector('[contenteditable]')).toBeNull()
		expect(doc.querySelector('image').getAttribute('href')).toBe('data:image/png;base64,iVBORw0K')
	})
})
