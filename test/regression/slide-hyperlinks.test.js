import { defineRegressionSuite, build, readEntry, assertIncludes, assertNotIncludes } from '../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// upstream-issue-1165: a hyperlink run with no color configured anywhere must
// inherit the theme hyperlink color (a:schemeClr hlink, and folHlink once
// visited). PptxGenJS used to default every run's color to DEF_FONT_COLOR
// ('000000'), which then emitted a solidFill plus an `ahyp:hlinkClr val="tx"`
// override — pinning the link to black and suppressing the automatic theme
// hyperlink/visited colors. The run must now carry a bare <a:hlinkClick/> with
// no fill and no hlinkClr override so PowerPoint applies the theme colors.
defineRegressionSuite('Slide hyperlink theme colors (upstream #1165)', [
	{
		name: 'external hyperlink without color emits bare hlinkClick (no fill, no hlinkClr)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('docs', { x: 1, y: 1, w: 4, h: 0.5, hyperlink: { url: 'https://example.com' } })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, '<a:hlinkClick r:id="rId1"', 'external hlinkClick present')
			assertIncludes(xml, 'endSnd="0"/>', 'hlinkClick self-closes (no children)')
			assertNotIncludes(xml, '<a:srgbClr val="000000"/>', 'no defaulted black fill on the link run')
			assertNotIncludes(xml, 'hlinkClr', 'no hlinkClr text-color override')
		},
	},
	{
		name: 'internal slide hyperlink without color emits bare hlinkClick',
		fn: async () => {
			const { zip } = await build((p) => {
				const pres = p
				pres.addSlide()
				pres.addSlide().addText('go', { x: 1, y: 1, w: 4, h: 0.5, hyperlink: { slide: 1 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			assertIncludes(xml, 'action="ppaction://hlinksldjump"', 'internal jump action present')
			assertNotIncludes(xml, '<a:srgbClr val="000000"/>', 'no defaulted black fill on the link run')
			assertNotIncludes(xml, 'hlinkClr', 'no hlinkClr text-color override')
		},
	},
	{
		name: 'hyperlink with explicit color still pins to that color via hlinkClr="tx"',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('red link', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					color: 'FF0000',
					hyperlink: { url: 'https://example.com' },
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, '<a:srgbClr val="FF0000"/>', 'explicit color fill present')
			assertIncludes(xml, 'ahyp:hlinkClr', 'hlinkClr override present for explicit color')
			assertIncludes(xml, 'val="tx"', 'hlinkClr pins link to the text color')
		},
	},
	{
		name: 'non-hyperlink text still defaults to black fill',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('plain', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, '<a:srgbClr val="000000"/>', 'normal text keeps the DEF_FONT_COLOR default')
		},
	},
])
