// `docProps/app.xml` describes the deck it was built from.
//
// Two of its values were constants whatever the deck: `PresentationFormat` always said
// `On-screen Show (16:9)` and `HiddenSlides` always `0`. PowerPoint rewrites both on its first
// save, so nothing was broken for a user who opened the file -- but anything reading the package
// without opening it (a document management system, a search indexer, our own tooling) read the
// constant, and a 4:3 deck or one with hidden slides stated something false.
//
// The five labels asserted here are PowerPoint's own, read back from decks it re-saved: three
// come from `p:sldSz@type` (`screen4x3`, `screen16x9`, `screen16x10`), `Widescreen` is what it
// calls 13.333in x 7.5in from the dimensions alone, and anything else is `Custom`.

import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'

/** `docProps/app.xml` for a deck at `layout`, with `hidden` slide numbers hidden. */
async function appXml(buildFn) {
	const { zip } = await build(buildFn)
	return readEntry(zip, 'docProps/app.xml')
}

const textOf = (xml, tag) => new RegExp(`<${tag}>(.*?)</${tag}>`).exec(xml)?.[1] ?? null

defineRegressionSuite('app.xml extended properties', [
	{
		name: 'PresentationFormat names the deck size, not a constant 16:9',
		fn: async () => {
			const cases = [
				['LAYOUT_4x3', 'On-screen Show (4:3)'],
				['LAYOUT_16x9', 'On-screen Show (16:9)'],
				['LAYOUT_16x10', 'On-screen Show (16:10)'],
				['LAYOUT_WIDE', 'Widescreen'],
			]
			for (const [layout, expected] of cases) {
				const xml = await appXml((p) => {
					p.layout = layout
					p.addSlide()
				})
				assertEqual(textOf(xml, 'PresentationFormat'), expected, layout)
			}
		},
	},
	{
		// PowerPoint's label follows `@type`, so a deck that states the label without the type is
		// relabelled `Custom` on its first save. Widescreen is the one standard size PowerPoint
		// itself writes untyped.
		name: 'p:sldSz declares the type the PresentationFormat label stands for',
		fn: async () => {
			const cases = [
				['LAYOUT_4x3', ' type="screen4x3"'],
				['LAYOUT_16x9', ' type="screen16x9"'],
				['LAYOUT_16x10', ' type="screen16x10"'],
				['LAYOUT_WIDE', ''],
			]
			for (const [layout, type] of cases) {
				const { zip, pres } = await build((p) => {
					p.layout = layout
					p.addSlide()
				})
				const { width, height } = pres.presLayout
				const xml = await readEntry(zip, 'ppt/presentation.xml')
				const expected = `<p:sldSz cx="${width}" cy="${height}"${type}/>`
				assert(xml.includes(expected), `${layout}: expected ${expected}`)
			}
		},
	},
	{
		name: 'a custom layout at a standard size takes that size type, and any other none',
		fn: async () => {
			const sldSzOf = async (width, height) => {
				const { zip } = await build((p) => {
					p.defineLayout({ name: 'Mine', width, height })
					p.layout = 'Mine'
					p.addSlide()
				})
				return /<p:sldSz [^>]*>/.exec(await readEntry(zip, 'ppt/presentation.xml'))[0]
			}
			assert((await sldSzOf(10, 5.625)).includes('type="screen16x9"'), 'a 16:9 on-screen defineLayout')
			assert(!(await sldSzOf(11.7, 8.3)).includes('type='), 'an A4 defineLayout is custom, the default')
		},
	},
	{
		name: 'a size PowerPoint has no label for is Custom',
		fn: async () => {
			const xml = await appXml((p) => {
				p.defineLayout({ name: 'A4', width: 11.7, height: 8.3 })
				p.layout = 'A4'
				p.addSlide()
			})
			assertEqual(textOf(xml, 'PresentationFormat'), 'Custom', 'a custom layout')
		},
	},
	{
		name: 'HiddenSlides counts the hidden slides',
		fn: async () => {
			const xml = await appXml((p) => {
				p.addSlide()
				p.addSlide().hidden = true
				p.addSlide().hidden = true
			})
			assertEqual(textOf(xml, 'HiddenSlides'), '2', 'two of three hidden')
			assertEqual(textOf(xml, 'Slides'), '3', 'a hidden slide is still a slide')
		},
	},
	{
		name: 'a deck with nothing hidden still says zero',
		fn: async () => {
			const xml = await appXml((p) => {
				p.addSlide()
				p.addSlide()
			})
			assertEqual(textOf(xml, 'HiddenSlides'), '0', 'nothing hidden')
			assert(xml.includes('<Application>Microsoft Office PowerPoint</Application>'), 'the rest of the part is intact')
		},
	},
])
