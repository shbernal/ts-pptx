import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertIncludes,
	assertNotIncludes,
} from '../../helpers.js'

// slide numbers defined on a master disappeared on slides that
// PowerPoint inserts from that master. Root cause: makeXmlMaster always emitted
// <p:hf sldNum="0" .../>, and CT_HeaderFooter/@sldNum defaults to true (ECMA-376), so the
// explicit "0" disabled the slide-number placeholder for inherited/new slides even though the
// master itself carried a sldNum placeholder shape. This suite pins the emitted master XML to be
// internally consistent: when a slide number is defined, the placeholder is present AND the
// header/footer element does not disable sldNum.
defineRegressionSuite('Master slide-number header/footer', [
	{
		name: 'master with slideNumber → sldNum placeholder present and <p:hf> does not disable sldNum',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'MASTER_WITH_SLDNUM',
					slideNumber: { x: 0.5, y: '90%' },
				})
			})
			const xml = await readEntry(zip, 'ppt/slideMasters/slideMaster1.xml')
			assertIncludes(xml, '<p:ph type="sldNum"', 'slide-number placeholder')
			const hf = (xml.match(/<p:hf\b[^>]*\/>/) || [])[0]
			assert(hf, 'expected a <p:hf .../> element; got: ' + xml)
			assertNotIncludes(hf, 'sldNum="0"', '<p:hf> must not disable sldNum when a slide number is defined')
		},
	},
	{
		name: 'master without slideNumber → <p:hf> keeps sldNum="0" (no inherited slide number)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'MASTER_NO_SLDNUM' })
			})
			const xml = await readEntry(zip, 'ppt/slideMasters/slideMaster1.xml')
			const hf = (xml.match(/<p:hf\b[^>]*\/>/) || [])[0]
			assert(hf, 'expected a <p:hf .../> element; got: ' + xml)
			assertIncludes(hf, 'sldNum="0"', '<p:hf> should disable sldNum when no slide number is defined')
		},
	},
	{
		// `<a:t>` inside an `a:fld` is the field's CACHED rendering. A master has no slide number
		// and a layout carries the library's internal 1000+ counter, so both used to ship a value
		// that reads as a page number and is not one -- `null` on a master, `1004` on a layout.
		// PowerPoint recomputes the field on open, so nothing looked wrong; anything reading the
		// cache without recomputing (a text extractor, a search indexer, this library's own read
		// path) saw it. What PowerPoint itself caches on a master or layout is the placeholder
		// glyph, and every en-US master and layout in `test/read/fixtures/` writes exactly this.
		name: 'a master and a layout cache the placeholder glyph, not a slide number',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'MASTER_WITH_SLDNUM',
					slideNumber: { x: 0.5, y: '90%' },
				})
			})
			const parts = listEntries(zip).filter((n) => /^ppt\/(slideMasters|slideLayouts)\/\w+\d+\.xml$/.test(n))
			let found = 0
			for (const part of parts) {
				const xml = await readEntry(zip, part)
				const fld = (xml.match(/<a:fld\b[^>]*type="slidenum"[\s\S]*?<\/a:fld>/) || [])[0]
				if (!fld) continue
				found++
				assertIncludes(fld, '<a:t>‹#›</a:t>', `${part} caches the placeholder glyph`)
			}
			assert(found >= 2, 'expected the field on both the master and its default layout; found ' + found)
		},
	},
	{
		// The other half of the same rule: a real slide DOES have a number to cache, and
		// PowerPoint caches the digits there. Pinned so the fix above cannot be over-applied.
		name: 'a slide caches its own number',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide()
				p.addSlide().slideNumber = { x: 0.5, y: '90%' }
			})
			const xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			const fld = (xml.match(/<a:fld\b[^>]*type="slidenum"[\s\S]*?<\/a:fld>/) || [])[0]
			assert(fld, 'expected a slidenum field on slide 2; got: ' + xml)
			assertIncludes(fld, '<a:t>2</a:t>', 'the second slide caches 2')
		},
	},
])
