import JSZip from 'jszip'
import {
	TsPptx,
	setDiagnosticHandler,
	captureDiagnostics,
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
} from '../../helpers.js'

/** A package's entry by name, from a written deck. */
async function entryOf(pres, name) {
	return readEntry(await JSZip.loadAsync(await pres.toBytes()), name)
}

/** The `code` `fn` throws, or `null`. */
function codeOf(fn) {
	try {
		fn()
		return null
	} catch (err) {
		return err.code ?? null
	}
}

const NOTES_XML = (n) => `ppt/notesSlides/notesSlide${n}.xml`
const NOTES_RELS = (n) => `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`
const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

defineRegressionSuite('Speaker notes hyperlinks & rich runs', [
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
						options: { hyperlink: { url: 'https://example.com/', tooltip: 'Docs' } },
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
			assertIncludes(rels, 'Target="https://example.com/"', 'rel target url')
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
			setDiagnosticHandler((d) => warnings.push(d.message))
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
				setDiagnosticHandler(null)
			}
		},
	},
	{
		// The notes part numbers its links in a namespace of its own. The id was stamped on the
		// hyperlink object, which a slide run can share, so the next write moved the slide run's
		// `r:id` onto whatever the slide declared under the notes id.
		name: 'one hyperlink object on a slide run and a notes run resolves in both parts, on every write',
		fn: async () => {
			const link = { url: 'https://shared.example/' }
			const pres = new TsPptx()
			const slide = pres.addSlide()
			slide.addText([{ text: 'slide run', options: { hyperlink: link } }], { x: 1, y: 1, w: 4, h: 1 })
			slide.addNotes([{ text: 'notes run', options: { hyperlink: link } }])
			for (const write of [1, 2]) {
				const zip = await JSZip.loadAsync(await pres.toBytes())
				const slideId = /<a:hlinkClick r:id="(rId\d+)"/.exec(await readEntry(zip, 'ppt/slides/slide1.xml'))?.[1]
				const slideRels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
				assert(
					new RegExp(
						`<Relationship Id="${slideId}"[^>]*relationships/hyperlink"[^>]*Target="https://shared.example/"`
					).test(slideRels),
					`write ${write}: the slide run's ${slideId} names its hyperlink relationship`
				)
				assertIncludes(await readEntry(zip, NOTES_XML(1)), '<a:hlinkClick r:id="rId3"', `write ${write}: the notes run`)
				assertIncludes(await readEntry(zip, NOTES_RELS(1)), 'Id="rId3"', `write ${write}: the notes relationship`)
			}
		},
	},
	{
		// The first write cached the notes relationships on the slide, so a link added afterwards was
		// never given an id and wrote `r:id="rIdundefined"`.
		name: 'notes added after a write are linked on the next write',
		fn: async () => {
			const pres = new TsPptx()
			const slide = pres.addSlide()
			slide.addNotes([{ text: 'early', options: { hyperlink: { url: 'https://early.example/' } } }])
			await pres.toBytes()
			slide.addNotes([{ text: 'late', options: { hyperlink: { url: 'https://late.example/' } } }])

			const xml = await entryOf(pres, NOTES_XML(1))
			assertNotIncludes(xml, 'rIdundefined', 'every notes link has an id')
			assertIncludes(xml, '<a:hlinkClick r:id="rId4"', 'the late link takes the next id')
			assertIncludes(await entryOf(pres, NOTES_RELS(1)), 'Target="https://late.example/"', 'and a relationship')
		},
	},
	{
		name: 'writing leaves the caller’s notes run options as they were',
		fn: async () => {
			const toSlide = { hyperlink: { slide: 1 } }
			const toUrl = { hyperlink: { url: 'https://u.example/' } }
			const before = JSON.stringify([toSlide, toUrl])
			const pres = new TsPptx()
			pres.addSlide().addNotes([
				{ text: 'to a slide', options: toSlide },
				{ text: 'to a url', options: toUrl },
			])
			await captureDiagnostics(() => pres.toBytes())
			assertEqual(JSON.stringify([toSlide, toUrl]), before, 'neither options object was written into')
		},
	},
	{
		// `addText` refuses all three when the text is added. `addNotes` accepted them: the first
		// then failed the write, and the other two were dropped without a word.
		name: 'addNotes refuses the hyperlinks addText refuses, when the notes are added',
		fn: async () => {
			for (const [hyperlink, code] of [
				[{ url: 'https://a.example/', slide: 1 }, 'hyperlink/conflicting-targets'],
				[{}, 'hyperlink/missing-target'],
				['https://a.example/', 'hyperlink/not-an-object'],
			]) {
				const slide = new TsPptx().addSlide()
				const options = { hyperlink: /** @type {any} */ (hyperlink) }
				assertEqual(
					codeOf(() => slide.addNotes([{ text: 'x', options }])),
					code,
					JSON.stringify(hyperlink)
				)
			}
		},
	},
])
