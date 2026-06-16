import { defineRegressionSuite, build, readEntry, assert, assertIncludes, assertNotIncludes } from '../helpers.js'

const NOTES_XML = (n) => `ppt/notesSlides/notesSlide${n}.xml`
const NOTES_RELS = (n) => `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`
const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

defineRegressionSuite('Speaker notes hyperlinks & rich runs (upstream #1250)', [
	{
		name: 'plain string notes still emit a single text run',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addNotes('Plain speaker notes')
			})
			const xml = await readEntry(zip, NOTES_XML(1))
			assertIncludes(xml, '<a:t>Plain speaker notes</a:t>', 'notes text')
			assertNotIncludes(xml, '<a:hlinkClick', 'no hyperlink for plain notes')
			// rels: only the reserved notesMaster + slide rels, no hyperlink rel
			const rels = await readEntry(zip, NOTES_RELS(1))
			assertNotIncludes(rels, HYPERLINK_REL_TYPE, 'no hyperlink rel for plain notes')
		},
	},
	{
		name: 'hyperlink run emits hlinkClick in notes body + external rel (rId3)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addNotes([
					{ text: 'See ' },
					{
						text: 'the docs',
						options: { hyperlink: { url: 'https://gitbrent.github.io/PptxGenJS/', tooltip: 'Docs' } },
					},
				])
			})
			const xml = await readEntry(zip, NOTES_XML(1))
			// Notes hyperlink rels start at rId3 (rId1=notesMaster, rId2=slide reserved)
			assertIncludes(xml, '<a:hlinkClick r:id="rId3"', 'hlinkClick references rId3')
			assertIncludes(xml, 'tooltip="Docs"', 'tooltip preserved')
			assertIncludes(xml, '<a:t>the docs</a:t>', 'hyperlink run text')

			const rels = await readEntry(zip, NOTES_RELS(1))
			assertIncludes(rels, 'Id="rId3"', 'rId3 relationship present')
			assertIncludes(rels, HYPERLINK_REL_TYPE, 'hyperlink rel type')
			assertIncludes(rels, 'Target="https://gitbrent.github.io/PptxGenJS/"', 'rel target url')
			assertIncludes(rels, 'TargetMode="External"', 'external target mode')
		},
	},
	{
		name: 'per-run formatting (bold/italic/color) serializes into notes runs',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addNotes([
					{ text: 'bold', options: { bold: true } },
					{ text: 'red', options: { color: 'FF0000', italic: true } },
				])
			})
			const xml = await readEntry(zip, NOTES_XML(1))
			assertIncludes(xml, 'b="1"', 'bold run')
			assertIncludes(xml, 'i="1"', 'italic run')
			assertIncludes(xml, '<a:srgbClr val="FF0000"/>', 'color run')
		},
	},
	{
		name: 'newlines split notes into separate paragraphs',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addNotes('line one\nline two')
			})
			const xml = await readEntry(zip, NOTES_XML(1))
			const paragraphs = (xml.match(/<a:p>/g) || []).length
			assert(paragraphs >= 2, `expected >=2 paragraphs for newline-separated notes; got ${paragraphs}: ${xml}`)
			assertIncludes(xml, '<a:t>line one</a:t>', 'first paragraph text')
			assertIncludes(xml, '<a:t>line two</a:t>', 'second paragraph text')
		},
	},
	{
		name: 'XML entities in notes hyperlink url + text are escaped',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addNotes([{ text: 'A & B', options: { hyperlink: { url: 'https://example.com/?a=1&b=2' } } }])
			})
			const xml = await readEntry(zip, NOTES_XML(1))
			assertIncludes(xml, '<a:t>A &amp; B</a:t>', 'escaped run text')
			const rels = await readEntry(zip, NOTES_RELS(1))
			assertIncludes(rels, 'Target="https://example.com/?a=1&amp;b=2"', 'escaped url in rel')
		},
	},
	{
		name: 'notes hyperlink `slide` target is ignored (url-only support)',
		fn: async () => {
			const warnings = []
			const originalWarn = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			try {
				const { zip } = await build((p) => {
					p.addSlide()
					p.addSlide().addNotes([{ text: 'jump', options: { hyperlink: { slide: 1 } } }])
				})
				const xml = await readEntry(zip, NOTES_XML(2))
				assertNotIncludes(xml, '<a:hlinkClick', 'slide-target hyperlink not emitted in notes')
				const rels = await readEntry(zip, NOTES_RELS(2))
				assertNotIncludes(rels, HYPERLINK_REL_TYPE, 'no hyperlink rel for slide target')
				assert(
					warnings.some((w) => w.includes('notes hyperlinks support `url` only')),
					`expected a warning about url-only notes hyperlinks; got: ${JSON.stringify(warnings)}`
				)
			} finally {
				console.warn = originalWarn
			}
		},
	},
])
