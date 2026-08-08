import {
	setDiagnosticHandler,
	build,
	readEntry,
	listEntries,
	defineRegressionSuite,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
	assertNonVisualDrawingProperty,
	firstXmlBlock,
	selfClosingTags,
	xmlAttributes,
	xmlOpeningTags,
} from '../../helpers.js'

// The *definition* side of Insert ▸ Zoom (`gen/define/zoom.ts` + `gen/define/preview-image.ts`),
// as distinct from `zoom-links.test.js`, which byte-pins the emitter given an already-built
// `SlideObject`. Everything here goes through the public builder, because that is the only way to
// reach the resolution the definer does: turning a Slide/number/section title into the `sldId`,
// section GUID and rIds the emitter consumes, and refusing — with a warning, never an exception —
// when the caller names something that does not exist.
//
// The refusals are the bulk of it. A zoom that cannot resolve its target is dropped silently
// today; a test that only ever passes valid input would not notice the day one of those guards
// starts throwing, or starts emitting a graphicFrame pointing at nothing.
//
// Three branches across the two modules are deliberately left red, all "unreachable by
// construction" in the sense of docs/testing.md:
//   - `if (!firstSlide) return` in the Summary Zoom tile loop. The list it walks was already
//     filtered on `_slides.length > 0`; the guard is there to keep the callback total, and its
//     own comment says so. Covering it would mean reaching past the public surface.
//   - the `'sm'` and `sl-<n>` arms of the media-key ternary in `preview-image.ts`. That ternary
//     is copied from `gen/define/image.ts`, where a master (`_slideNum === null`) or a layout
//     (`_slideNum >= 1000`) really can own an image. A preview image cannot: the only two callers
//     are the zoom and OLE definers, both reachable only from `SlideBuilder`, which always carries
//     a 1-based `_slideNum`. Dead here, load-bearing where it was copied from.

/** A jpg on disk, so a cover image can be supplied by `path` with no `data` alongside it. */
const COVER_JPG = 'demos/common/images/cc_logo.jpg'

/** Build, capturing library warnings (`log.ts` routes every one through `console.warn`). */
async function buildCapturingWarnings(buildFn) {
	const warnings = []
	setDiagnosticHandler((d) => warnings.push(d.message))
	try {
		const result = await build(buildFn)
		return { ...result, warnings }
	} finally {
		setDiagnosticHandler(null)
	}
}

function assertWarned(warnings, pattern, label) {
	assert(
		warnings.some((message) => pattern.test(message)),
		`expected a warning matching ${pattern} ${label || ''}; got: ${JSON.stringify(warnings)}`
	)
}

/** @param {string} tag @returns {Record<string, string>} */
const attrs = (tag) => /** @type {Record<string, string>} */ (xmlAttributes(tag))

/**
 * The `<a:off>`/`<a:ext>` of a graphicFrame's own `<p:xfrm>` (the zoom frame, not a tile).
 * @param {string} xml
 * @returns {Record<string, string>}
 */
function frameExtent(xml) {
	const frame = firstXmlBlock(xml, 'p:graphicFrame')
	const xfrm = firstXmlBlock(frame, 'p:xfrm')
	const [off, ext] = [selfClosingTags(xfrm, 'a:off')[0], selfClosingTags(xfrm, 'a:ext')[0]]
	return { ...attrs(off), ...attrs(ext) }
}

/** Every `Relationship` in slide 1's rels part, as `{ id, type, target }`. */
async function slideRels(zip) {
	const xml = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
	return selfClosingTags(xml, 'Relationship')
		.map((tag) => xmlAttributes(tag))
		.map((attrs) => ({ id: attrs.Id, type: attrs.Type, target: attrs.Target }))
}

defineRegressionSuite('Zoom definition', [
	{
		// `!opts?.target` also swallows `target: 0`, which is what a caller reaching for a 0-based
		// slide index would pass. Zoom targets are 1-based, so 0 is not a slide either way.
		name: 'addSlideZoom without a target warns and emits nothing',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				const host = p.addSlide()
				p.addSlide()
				host.addSlideZoom(/** @type {any} */ ({ x: 1, y: 1, w: 3, h: 1.7 }))
				host.addSlideZoom(/** @type {any} */ ({ target: 0, x: 5, y: 1, w: 3, h: 1.7 }))
			})
			assertEqual(warnings.length, 2, `expected one warning per dropped zoom; got: ${JSON.stringify(warnings)}`)
			assertWarned(warnings, /addSlideZoom requires a `target` slide/, 'for a missing target')
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'slidezoom', 'slide 1')
		},
	},
	{
		// Reachable from JavaScript, where the `Slide | number` type is not enforced: anything that
		// is truthy but not a slide gets past the first guard and resolves to no ids at all.
		name: 'addSlideZoom with an unresolvable target warns and emits nothing',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				const host = p.addSlide()
				host.addSlideZoom(/** @type {any} */ ({ target: {}, x: 1, y: 1, w: 3, h: 1.7 }))
			})
			assertWarned(warnings, /addSlideZoom: could not resolve the target slide/, 'for a non-slide target')
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'slidezoom', 'slide 1')
		},
	},
	{
		// Position is optional on every zoom kind, so a caller can register one and place it later
		// in PowerPoint. The frame must still be well-formed — an omitted `w` becomes an explicit
		// `cx="0"`, not a missing attribute, which is what PowerPoint requires of `<a:ext>`.
		name: 'a zoom with no geometry gets a zero-extent frame and keeps its objectName',
		fn: async () => {
			const { zip } = await build((p) => {
				const host = p.addSlide()
				p.addSlide()
				host.addSlideZoom({ target: 2, objectName: 'Nav tile' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNonVisualDrawingProperty(xml, { name: 'Nav tile' }, 'the zoom frame')
			const extent = frameExtent(xml)
			assertEqual(extent.x, '0', 'frame x')
			assertEqual(extent.y, '0', 'frame y')
			assertEqual(extent.cx, '0', 'frame cx')
			assertEqual(extent.cy, '0', 'frame cy')
		},
	},
	{
		// Targeting by number computes `sldId` as 256 + (n - 1) without consulting the slide, so
		// the two addressing modes have to be checked to agree; a Slide object reports its own id.
		name: 'targeting by number and by Slide object resolve to the same sldId',
		fn: async () => {
			const { zip } = await build((p) => {
				const host = p.addSlide()
				const other = p.addSlide()
				host.addSlideZoom({ target: 2, x: 1, y: 1, w: 3, h: 1.7 })
				host.addSlideZoom({ target: other, x: 5, y: 1, w: 3, h: 1.7 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const ids = [...xml.matchAll(/<pslz:sldZmObj sldId="(\d+)"/g)].map((m) => m[1])
			assertEqual(ids.length, 2, 'expected two slide zooms')
			assertEqual(ids[0], ids[1], 'the numeric and object forms should resolve to the same sldId')

			// Each zoom mints its own preview-image rel and its own `.../slide` fallback rel.
			const rels = await slideRels(zip)
			const slideRelTargets = rels.filter((r) => r.type.endsWith('/slide')).map((r) => r.target)
			assertEqual(slideRelTargets.length, 2, 'expected one fallback slide rel per zoom')
			assertEqual(new Set(slideRelTargets).size, 1, 'both fallbacks point at the same slide')
			assertEqual(slideRelTargets[0], 'slide2.xml', 'fallback target')
		},
	},
	{
		name: 'addSectionZoom without a sectionTitle warns and emits nothing',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				p.addSection({ title: 'Alpha' })
				p.addSlide({ sectionTitle: 'Alpha' }).addSectionZoom(/** @type {any} */ ({ x: 1, y: 1, w: 3, h: 1.7 }))
			})
			assertWarned(warnings, /addSectionZoom requires a `sectionTitle`/, 'for a missing title')
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'sectionzoom', 'slide 1')
		},
	},
	{
		name: 'addSectionZoom naming a section that does not exist warns and emits nothing',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				p.addSection({ title: 'Alpha' })
				p.addSlide({ sectionTitle: 'Alpha' }).addSectionZoom({ sectionTitle: 'Omega', x: 1, y: 1, w: 3, h: 1.7 })
			})
			assertWarned(warnings, /no section titled "Omega"/, 'for an unknown section')
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'sectionzoom', 'slide 1')
		},
	},
	{
		// A section exists as soon as it is declared; slides join it later (or never). A zoom to an
		// empty one has no first slide to link to, so it is dropped rather than emitted with a
		// dangling `hlinksldjump`.
		name: 'addSectionZoom to a section with no slides warns and emits nothing',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				p.addSection({ title: 'Alpha' })
				p.addSection({ title: 'Empty' })
				p.addSlide({ sectionTitle: 'Alpha' }).addSectionZoom({ sectionTitle: 'Empty', x: 1, y: 1, w: 3, h: 1.7 })
			})
			assertWarned(warnings, /section "Empty" has no slides/, 'for an empty section')
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'sectionzoom', 'slide 1')
		},
	},
	{
		name: 'addSectionZoom resolves the section GUID and links to its first slide',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSection({ title: 'Alpha' })
				p.addSection({ title: 'Beta' })
				const host = p.addSlide({ sectionTitle: 'Alpha' })
				p.addSlide({ sectionTitle: 'Beta' })
				p.addSlide({ sectionTitle: 'Beta' })
				host.addSectionZoom({ sectionTitle: 'Beta', x: 1, y: 1, w: 3, h: 1.7 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const zoomed = xml.match(/<psez:sectionZmObj sectionId="(\{[0-9A-F-]+\})"/)
			assert(zoomed, `expected a section zoom carrying a GUID; got: ${xml}`)

			// The GUID must be the one `presentation.xml` declares for Beta, or PowerPoint drops the link.
			const pres = await readEntry(zip, 'ppt/presentation.xml')
			const beta = xmlOpeningTags(pres, 'p14:section')
				.map((tag) => xmlAttributes(tag))
				.find((attrs) => attrs.name === 'Beta')
			assert(beta, `expected a p14:section named Beta; got: ${pres}`)
			assertEqual(zoomed[1], beta.id, 'the zoom should carry the declared section id')

			const rels = await slideRels(zip)
			const fallback = rels.find((r) => r.type.endsWith('/slide'))
			assert(fallback, 'expected a fallback slide rel')
			assertEqual(fallback.target, 'slide2.xml', "the fallback should point at Beta's first slide")
		},
	},
	{
		// A summary excludes the host slide's own section, so a deck with only that section has
		// nothing left to summarize.
		name: 'addSummaryZoom with nothing to summarize warns and emits nothing',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				p.addSection({ title: 'Only' })
				p.addSlide({ sectionTitle: 'Only' }).addSummaryZoom({ x: 0.5, y: 1.5, w: 11, h: 4.5 })
			})
			assertWarned(warnings, /addSummaryZoom: no sections to summarize/, 'for a single-section deck')
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'summaryzoom', 'slide 1')
		},
	},
	{
		// The grid fits tiles to whichever of the two frame dimensions binds first. A frame wider
		// than the tiles need is the width-bound case: tile height comes out under the row budget
		// and is left alone, so the grid is centered vertically rather than stretched.
		name: 'summary grid lays out one tile per section, width-bound in a short-and-wide frame',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSection({ title: 'Host' })
				p.addSection({ title: 'Alpha' })
				p.addSection({ title: 'Beta' })
				const host = p.addSlide({ sectionTitle: 'Host' })
				p.addSlide({ sectionTitle: 'Alpha' })
				p.addSlide({ sectionTitle: 'Beta' })
				host.addSummaryZoom({ x: 0.5, y: 1, w: 9, h: 4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const tiles = [...xml.matchAll(/<psuz:summaryZmObj sectionId="(\{[0-9A-F-]+\})"/g)].map((m) => m[1])
			assertEqual(tiles.length, 2, "expected a tile per section other than the host's")
			assertEqual(new Set(tiles).size, 2, 'the two tiles should address different sections')

			// Two sections → a 2×1 grid: same y and size, second offset to the right of the first.
			const zm = xml.slice(xml.indexOf('<psuz:summaryZm>'), xml.indexOf('<psuz:gridLayout/>'))
			const cells = selfClosingTags(zm, 'a:off').map((tag, i) => ({
				...attrs(tag),
				...attrs(selfClosingTags(zm, 'a:ext')[i]),
			}))
			assertEqual(cells.length, 2, 'expected two tile xfrms')
			assertEqual(cells[0].y, cells[1].y, 'both tiles should sit on one row')
			assertEqual(cells[0].cx, cells[1].cx, 'tiles should be the same width')
			assert(
				Number(cells[1].x) > Number(cells[0].x),
				`expected the second tile to the right of the first; got ${cells[0].x} then ${cells[1].x}`
			)
			// Width-bound: the pair plus the gap fills the frame, and there is slack above/below.
			// Tile sizes are rounded to whole EMU, so the span can miss the frame by a unit or two —
			// an EMU is 1/914400 inch, well under anything PowerPoint renders differently.
			const frame = frameExtent(xml)
			const spanned = Number(cells[1].x) + Number(cells[1].cx)
			assert(
				Math.abs(spanned - Number(frame.cx)) <= 2,
				`expected the grid to fill the frame width (${frame.cx}); it spans ${spanned}`
			)
			assert(
				Number(cells[0].cy) < Number(frame.cy),
				`expected height slack in a width-bound grid; tile cy ${cells[0].cy} vs frame cy ${frame.cy}`
			)
		},
	},
	{
		name: 'a summary zoom with no geometry still emits a tile per section',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSection({ title: 'Host' })
				p.addSection({ title: 'Alpha' })
				const host = p.addSlide({ sectionTitle: 'Host' })
				p.addSlide({ sectionTitle: 'Alpha' })
				host.addSummaryZoom({})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, '<psuz:summaryZmObj sectionId=', 'slide 1')
			const extent = frameExtent(xml)
			assertEqual(extent.cx, '0', 'frame cx')
			assertEqual(extent.cy, '0', 'frame cy')
		},
	},
	{
		// Every tile draws from a preview raster. Supplied by `path` there is no `data:` mime to
		// sniff, so the extension has to come off the path — and a second tile pointing at the same
		// file must reuse the first rel's Target rather than shipping the bytes twice.
		name: 'a cover image supplied by path is sniffed from the path and shipped once',
		fn: async () => {
			const { zip } = await build((p) => {
				const host = p.addSlide()
				p.addSlide()
				host.addSlideZoom({ target: 2, x: 1, y: 1, w: 3, h: 1.7, coverImage: { path: COVER_JPG } })
				host.addSlideZoom({ target: 2, x: 5, y: 1, w: 3, h: 1.7, coverImage: { path: COVER_JPG } })
			})
			const rels = await slideRels(zip)
			const previews = rels.filter((r) => r.type.endsWith('/image'))
			assertEqual(previews.length, 2, 'expected a preview rel per zoom')
			assertEqual(previews[0].target, previews[1].target, 'the second cover should reuse the first rel Target')
			assert(previews[0].target.endsWith('.jpg'), `expected a .jpg target; got ${previews[0].target}`)

			const media = listEntries(zip).filter((name) => name.startsWith('ppt/media/'))
			assertEqual(media.length, 1, `the cover should be stored once; got: ${media.join(', ')}`)

			// Both tiles reference their own rel, and both rels resolve to that single part.
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const embeds = new Set([...xml.matchAll(/<a:blip r:embed="([^"]+)"/g)].map((m) => m[1]))
			assertEqual(embeds.size, 2, `expected two distinct preview rIds; got ${[...embeds].join(' ')}`)
			for (const preview of previews) {
				assert(embeds.has(preview.id), `slide 1 should reference preview rel ${preview.id}`)
			}
		},
	},
])
