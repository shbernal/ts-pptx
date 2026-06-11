import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// Regression: an autoPage table that overflows onto continuation slides should be able to
// carry the source slide's populated placeholders (e.g. a title) onto every overflow slide.
// Overflow slides otherwise inherit only the layout's empty placeholders, so a title set on
// the first slide vanishes on continuation slides (upstream gitbrent/PptxGenJS#1136).

const TITLE = 'Quarterly Report'

function deck(autoPagePlaceholder) {
	return build((p) => {
		p.defineSlideMaster({
			title: 'TEST_MASTER_1136',
			objects: [
				{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.05, w: 9, h: 0.25 }, text: '' } },
			],
		})
		const s = p.addSlide({ masterName: 'TEST_MASTER_1136' })
		s.addText(TITLE, { placeholder: 'title' })
		// Enough rows to overflow several slides by natural slide height (no `h` cap, which can
		// otherwise produce a degenerate empty overflow page in getSlidesForTableRows).
		const rows = Array.from({ length: 60 }, (_, i) => [{ text: `Row ${i} col A` }, { text: `Row ${i} col B` }])
		s.addTable(rows, {
			x: 0.5,
			y: 0.4,
			w: 9,
			colW: [4.5, 4.5],
			margin: 0,
			slideMargin: 0,
			autoPage: true,
			autoPagePlaceholder,
			fontSize: 14,
		})
	})
}

function overflowSlideFiles(zip) {
	return listEntries(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort()
}

defineRegressionSuite('Table autoPage placeholder propagation (upstream #1136)', [
	{
		name: 'autoPagePlaceholder:true copies the title onto every overflow slide',
		fn: async () => {
			const { zip } = await deck(true)
			const slides = overflowSlideFiles(zip)
			assert(slides.length >= 2, `expected ≥2 slides (source + overflow); got ${slides.length}`)
			for (const file of slides) {
				const xml = await readEntry(zip, file)
				assert(xml.includes(`<a:t>${TITLE}</a:t>`), `expected title "${TITLE}" on ${file}; got: ${xml.slice(0, 400)}`)
				assert(/<p:ph[^>]*type="title"/.test(xml), `expected a title placeholder on ${file}`)
			}
		},
	},
	{
		name: 'default (no flag) leaves the title only on the source slide',
		fn: async () => {
			const { zip } = await deck(false)
			const slides = overflowSlideFiles(zip)
			assert(slides.length >= 2, `expected ≥2 slides; got ${slides.length}`)
			const firstXml = await readEntry(zip, slides[0])
			assert(firstXml.includes(`<a:t>${TITLE}</a:t>`), 'source slide must still carry the populated title')
			for (const file of slides.slice(1)) {
				const xml = await readEntry(zip, file)
				assert(!xml.includes(`<a:t>${TITLE}</a:t>`), `overflow slide ${file} must not duplicate the title by default`)
			}
		},
	},
])
