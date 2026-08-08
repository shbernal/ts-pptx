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
} from '../../helpers.js'

// The option-normalization half of `gen/define/text.ts` — the work `addTextDefinition` does before
// any XML exists: defaulting an empty text array, computing line defaults for a line-shaped text
// box, rejecting out-of-range `columns`/`columnSpacing`, expanding the `underline: true` shorthand,
// mapping `align`/`valign` onto `_bodyProp`, and registering picture-bullet media rels.
//
// Everything here goes through the public builder (`addText` / `defineSlideMaster`), because the
// normalization is only observable in the emitted package: the definer mutates the caller's options
// object and the emitters read `_bodyProp` off it, so asserting on the options directly would pin
// the wrong side of the contract.
//
// The picture-bullet cases are four-way rather than repetitive: `createBulletImageRels` branches on
// SVG-vs-raster AND on data-vs-path, and each of the four combinations takes a different `path`
// fallback into `_relsMedia`. Existing coverage passed `path` and `data` together, so both fallbacks
// (`img.path || 'preencoded.' + extn` and `img.data || ''`) were unexercised in every combination.
//
// One thing to know before reading the cases, because it decides which form of `addText` each uses:
// `cleanOpts` runs once for the text object and again for every run, and `SlideBuilder.addText`'s
// string shorthand hands the SAME options object to both. So the shorthand cleans its options
// TWICE while the array form cleans the shape's options once. That is invisible for most options
// and decisive for two of them -- see the line-defaults and picture-bullet cases below.
//
// Left deliberately red, all "unreachable by construction" in the sense of docs/testing.md:
//   - `const objectOptions = opts || {}` (L36). All four callers -- `SlideBuilder.addText`,
//     `addChildDefinition`, `createSlideMaster` and `addPlaceholdersToSlideLayouts` -- pass an
//     object. It could not survive being reached anyway: four lines later `opts.shape` is read
//     unguarded, so a caller that passed `undefined` would throw before the fallback mattered.
//   - `target._slideNum == null ? 'sm' : ...` (L256), the first arm of the media-key triple. The
//     only slide with a null `_slideNum` is `presentation.ts`'s `_masterSlide`, a deliberately-partial
//     stub whose authoring methods are all `null` and which no definer is ever pointed at. The same
//     dead arm appears verbatim in define/image.ts (twice), define/ole.ts and define/preview-image.ts;
//     the layout arm below is the one a caller can actually reach.
//   - `_placeholderType ||` -> `` `Placeholder ${...}` `` (L176), the last rung of the placeholder
//     objectName ladder, and the `_placeholderIdx ?? _slideObjects.length` inside it. Reaching the
//     template needs BOTH `name` and `type` empty, and `type` is a `PLACEHOLDER_TYPE` enum with no
//     empty member, so it is schema-impossible; the `??` behind it is doubly dead, since
//     `createSlideMaster` always assigns `_placeholderIdx = 100 + idx`.
//   - `img.path || img.data || 'preencoded.svg'` (L272), third arm. The loop returns early when a
//     bullet image has neither a path nor data, so the last fallback has no input.

// 1x1 PNG (red pixel).
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='
// 1x1 JPEG, used where a second, distinct payload is needed (identical bytes dedupe onto one part).
const JPG_DATA =
	'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AL+AB//Z'
// Minimal SVG, supplied as bytes so nothing has to be read from disk.
const SVG_DATA =
	'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxIiBoZWlnaHQ9IjEiLz4='
// Real files, for the path-only halves. Both are read from disk during export, so they have to exist.
const JPG_PATH = 'demos/common/images/cc_logo.jpg'
const SVG_PATH = 'demos/common/images/lock-green.svg'

/** Build, capturing library diagnostics as `{ code, message }` pairs. */
async function buildCapturingLogs(buildFn) {
	const warnings = []
	setDiagnosticHandler((d) => warnings.push(d))
	try {
		const result = await build(buildFn)
		return { ...result, warnings }
	} finally {
		setDiagnosticHandler(null)
	}
}

/** Media part names in the package, in zip order. */
const mediaEntries = (zip) => listEntries(zip).filter((name) => name.startsWith('ppt/media/'))

/** Build a one-slide deck whose only text carries `bullet.image`, and return its media parts. */
async function bulletMedia(image) {
	const { zip } = await build((p) => {
		p.addSlide().addText('bulleted', { x: 1, y: 1, w: 4, h: 1, bullet: { image } })
	})
	return { zip, media: mediaEntries(zip) }
}

defineRegressionSuite('Text definition', [
	{
		// `addText([])` is the one caller shape that reaches the empty-text default: the string and
		// number forms are wrapped into a one-run array by `SlideBuilder.addText` before the definer
		// sees them. The result is a real (empty) text box, not a dropped object -- callers build
		// empty placeholders this way and fill them later.
		name: 'an empty text array becomes a single empty run',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText([], { x: 1, y: 1, w: 4, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNonVisualDrawingProperty(xml, { name: 'Text 0' }, 'the empty text box')
			const runs = xml.match(/<a:t>[\s\S]*?<\/a:t>/g) || []
			assertEqual(runs.length, 1, `expected exactly one run; got ${JSON.stringify(runs)}`)
			assertEqual(runs[0], '<a:t></a:t>', 'the run is empty')
		},
	},
	{
		// The ShapeLineProps defaults for a line-shaped text box are computed unconditionally but only
		// *assigned* back when `itemOpts.line` is already an object -- and on a first pass over options
		// that carry no `line` at all, it is not. STEP C's `itemOpts.line = itemOpts.line || {}` then
		// installs a bare `{}`, which is what a SECOND pass over the same object sees. So the string
		// shorthand, which cleans one shared options object twice, ends up with the full 1pt solid
		// default, while the array form -- one pass over the shape's options -- emits an empty
		// `<a:ln>`. Same call, same `shape: 'line'`, different line, decided entirely by which
		// overload the caller reached for. Pinned as-is: these are the bytes today's callers get.
		name: 'line defaults reach a line-shaped text box only on a second options pass',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('shorthand', { x: 1, y: 1, w: 4, h: 0.1, shape: 'line' })
				s.addText([{ text: 'array' }], { x: 1, y: 2, w: 4, h: 0.1, shape: 'line' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapes = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assertEqual(shapes.length, 2, 'expected both line shapes')
			// The defaults: 1pt (12700 EMU), DEF_SHAPE_LINE_COLOR, solid dash.
			assertIncludes(shapes[0], '<a:ln w="12700"', 'the twice-cleaned shorthand')
			assertIncludes(shapes[0], '<a:prstDash val="solid"/>', 'the twice-cleaned shorthand')
			assertIncludes(shapes[1], '<a:ln></a:ln>', 'the once-cleaned array form')
		},
	},
	{
		// `numCol` is ST_TextColumnCount, which ECMA-376 bounds at 1-16. Out-of-range and non-numeric
		// values warn and are dropped rather than clamped, so a typo does not silently reflow the box.
		// Three shapes: below the floor, above the ceiling, and not a number at all. The array form is
		// used so each rejection warns exactly once (the shorthand would clean the same options twice
		// and warn twice per box) -- the count is what proves each value was judged on its own.
		name: 'out-of-range text columns warn and leave bodyPr untouched',
		fn: async () => {
			const { zip, warnings } = await buildCapturingLogs((p) => {
				const s = p.addSlide()
				s.addText([{ text: 'too few' }], { x: 1, y: 1, w: 4, h: 1, columns: 0 })
				s.addText([{ text: 'too many' }], { x: 1, y: 2, w: 4, h: 1, columns: 17 })
				s.addText([{ text: 'not a number' }], { x: 1, y: 3, w: 4, h: 1, columns: /** @type {any} */ ('three') })
			})
			assertEqual(
				warnings.filter((d) => d.code === 'text/invalid-columns').length,
				3,
				`expected one warning per rejected value; got ${JSON.stringify(warnings)}`
			)
			assertNotIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'numCol', 'slide1')
		},
	},
	{
		// Same policy on the gutter: `spcCol` is a length, so a negative one is meaningless. Zero is
		// the boundary the `< 0` test sits on and it is accepted -- silently, which is the only way to
		// tell it apart from the rejection, since the emitter then omits the attribute anyway (zero is
		// already the default gutter). The third box shows a spacing that does survive to the bytes.
		name: 'a negative column spacing warns; zero is accepted and positive spacing is emitted',
		fn: async () => {
			const { zip, warnings } = await buildCapturingLogs((p) => {
				const s = p.addSlide()
				s.addText([{ text: 'negative' }], { x: 1, y: 1, w: 4, h: 1, columns: 2, columnSpacing: -1 })
				s.addText([{ text: 'zero' }], { x: 1, y: 2, w: 4, h: 1, columns: 2, columnSpacing: 0 })
				s.addText([{ text: 'positive' }], { x: 1, y: 3, w: 4, h: 1, columns: 2, columnSpacing: 20 })
			})
			assertEqual(
				warnings.filter((d) => d.code === 'text/invalid-column-spacing').length,
				1,
				`expected only the negative spacing to warn; got ${JSON.stringify(warnings)}`
			)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapes = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assertEqual(shapes.length, 3, 'expected all three text boxes')
			assertNotIncludes(shapes[0], 'spcCol', 'the rejected negative spacing')
			assertNotIncludes(shapes[1], 'spcCol', 'the accepted zero spacing')
			assertIncludes(shapes[2], 'spcCol="254000"', 'the positive spacing, 20pt in EMU')
			// Every box kept its columns: the spacing is judged on its own, not the whole block.
			for (const shape of shapes) assertIncludes(shape, 'numCol="2"', 'a columns box')
		},
	},
	{
		// `underline: true` is shorthand for the object form. It is normalized here rather than in the
		// emitter so that `u="sng"` is the only thing the emitter has to know about; `underline: false`
		// is left alone (falsy, so no attribute) and the object form passes through untouched.
		name: 'underline: true is normalized to the single-line object form',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('shorthand', { x: 1, y: 1, w: 4, h: 1, underline: true })
				s.addText('explicit', { x: 1, y: 2, w: 4, h: 1, underline: { style: 'dbl' } })
				s.addText('off', { x: 1, y: 3, w: 4, h: 1, underline: /** @type {any} */ (false) })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapes = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assertEqual(shapes.length, 3, 'expected all three text boxes')
			assertIncludes(shapes[0], 'u="sng"', 'the shorthand form')
			assertIncludes(shapes[1], 'u="dbl"', 'the explicit form')
			assertNotIncludes(shapes[2], ' u="', 'underline: false')
		},
	},
	{
		// `align` is matched by first letter, so the public spellings and any prefix of them land on
		// the same ST_TextAlignType value. Center and justify were the two arms no fixture used.
		name: 'align is resolved by first letter, including center and justify',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('c', { x: 1, y: 1, w: 4, h: 1, align: 'center' })
				s.addText('j', { x: 1, y: 2, w: 4, h: 1, align: 'justify' })
				s.addText('l', { x: 1, y: 3, w: 4, h: 1, align: 'left' })
				s.addText('r', { x: 1, y: 4, w: 4, h: 1, align: 'right' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapes = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assertEqual(shapes.length, 4, 'expected all four text boxes')
			const algn = shapes.map((shape) => (shape.match(/algn="([a-z]+)"/) || [])[1])
			assertEqual(algn.join(','), 'ctr,just,l,r', 'align resolution')
		},
	},
	{
		// A picture bullet given as bytes must carry a base64 header, exactly as `addImage()` requires.
		// Without one the definer refuses to register the rel -- and because no rel exists, the run
		// emitter falls back to a default glyph and says so. Two diagnostics saying different things
		// (why the rel was refused, and what was drawn instead), no media part, and a deck that still
		// opens -- which is why this warns where `addImage()` on the same input throws: there is a
		// sane thing to draw. The bullet sits on the run rather than the shape so that the
		// collection pass sees it once; the shorthand would present the same object twice and the
		// base64 refusal, unlike the rel registration, is not deduped.
		name: 'a bullet image whose data lacks a base64 header is refused and falls back to a glyph',
		fn: async () => {
			const { zip, warnings } = await buildCapturingLogs((p) => {
				p.addSlide().addText([{ text: 'bulleted', options: { bullet: { image: { data: 'iVBORw0KGgoAAAA==' } } } }], {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
				})
			})
			assertEqual(
				warnings.filter((d) => d.code === 'bullet/image-missing-base64-header').length,
				1,
				`expected the definer's reality-check to fire; got ${JSON.stringify(warnings)}`
			)
			assertEqual(
				warnings.filter((d) => d.code === 'bullet/image-embed-failed').length,
				1,
				`expected the emitter's fallback warning; got ${JSON.stringify(warnings)}`
			)
			assertEqual(mediaEntries(zip).length, 0, 'expected no media part for the refused bullet')
			assertIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), '<a:buChar char="&#x2022;"/>', 'slide1')
		},
	},
	{
		// Raster picture bullets, both halves of the data-vs-path fork. The media part is named from
		// the extension the definer sniffed, which comes from the `data:` mime when there are bytes and
		// from the file name when there is a path -- so these two decks disagree on the part name for
		// the same reason they disagree on the recorded source path (`preencoded.png` vs the real file).
		name: 'a raster picture bullet registers one media part from either bytes or a path',
		fn: async () => {
			const fromData = await bulletMedia({ data: PNG_DATA })
			assertEqual(fromData.media.join(','), 'ppt/media/image-1-1.png', 'bytes-only bullet')

			const fromPath = await bulletMedia({ path: JPG_PATH })
			assertEqual(fromPath.media.join(','), 'ppt/media/image-1-1.jpg', 'path-only bullet')

			// Both are referenced by the paragraph's `<a:buBlip>`, through the rel the definer assigned.
			for (const { zip } of [fromData, fromPath]) {
				const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
				const embed = (xml.match(/<a:buBlip><a:blip r:embed="(rId\d+)"/) || [])[1]
				assert(embed, `expected a picture bullet in: ${xml}`)
				assertIncludes(await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels'), `Id="${embed}"`, 'slide1 rels')
			}
		},
	},
	{
		// SVG picture bullets consume TWO rels apiece -- a PNG preview for the `<a:blip r:embed>` plus
		// the SVG itself behind the `asvg:svgBlip` extension -- so each of these decks carries two media
		// parts, in preview-then-source order. Same data-vs-path fork as the raster case above.
		name: 'an SVG picture bullet registers a PNG preview alongside the SVG, from bytes or a path',
		fn: async () => {
			const fromData = await bulletMedia({ data: SVG_DATA })
			assertEqual(fromData.media.join(','), 'ppt/media/image-1-1.png,ppt/media/image-1-2.svg', 'bytes-only SVG')

			const fromPath = await bulletMedia({ path: SVG_PATH })
			assertEqual(fromPath.media.join(','), 'ppt/media/image-1-1.png,ppt/media/image-1-2.svg', 'path-only SVG')

			// The blip embeds the PNG preview and points at the SVG via the extension, two distinct rels.
			for (const { zip } of [fromData, fromPath]) {
				const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
				const preview = (xml.match(/<a:buBlip><a:blip r:embed="(rId\d+)"/) || [])[1]
				const svg = (xml.match(/<asvg:svgBlip[^>]*r:embed="(rId\d+)"/) || [])[1]
				assert(preview && svg, `expected both bullet blips in: ${xml}`)
				assert(preview !== svg, 'the preview and the SVG take separate rels')
			}
		},
	},
	{
		// Media file names are namespaced by the slide they belong to, and layouts use an `sl-<num>`
		// key rather than a bare index so a layout's bullet cannot collide with slide 1's. This is the
		// reachable half of the media-key triple; see the header for why the master half is not.
		// The two bullets carry different payloads on purpose: identical bytes dedupe onto one part,
		// which would hide the naming this case is about.
		name: 'a picture bullet on a layout is namespaced away from the slides',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'BULLET_MASTER',
					objects: [
						{
							text: {
								text: 'layout bullet',
								options: { x: 1, y: 1, w: 4, h: 1, bullet: { image: { data: JPG_DATA } } },
							},
						},
					],
				})
				p.addSlide({ masterTitle: 'BULLET_MASTER' }).addText('slide bullet', {
					x: 1,
					y: 3,
					w: 4,
					h: 1,
					bullet: { image: { data: PNG_DATA } },
				})
			})
			// `sl-1002`, not `sl-1001`: the built-in DEFAULT layout is already in the list when
			// `defineSlideMaster` numbers this one.
			assertEqual(
				mediaEntries(zip).sort().join(','),
				'ppt/media/image-1-1.png,ppt/media/image-sl-1002-1.jpeg',
				'layout and slide bullet media'
			)
			assertIncludes(
				await readEntry(zip, 'ppt/slideLayouts/_rels/slideLayout2.xml.rels'),
				'Target="../media/image-sl-1002-1.jpeg"',
				'the layout rels'
			)
		},
	},
	{
		// A master placeholder's Selection Pane identity defaults to its declared `name`. An empty name
		// is type-legal and would put an unnamed shape in the pane, so it falls through to the
		// placeholder *type* instead -- the same ladder `addText({ placeholder })` relies on.
		name: 'a placeholder with an empty name takes its identity from its type',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'NAMELESS',
					objects: [
						{ placeholder: { options: { name: '', type: 'title', x: 1, y: 1, w: 8, h: 1 }, text: 'Untitled' } },
						{ placeholder: { options: { name: 'subhead', type: 'body', x: 1, y: 3, w: 8, h: 1 }, text: 'Sub' } },
					],
				})
				p.addSlide({ masterTitle: 'NAMELESS' })
			})
			// slideLayout2: the built-in DEFAULT layout takes slideLayout1.
			const xml = await readEntry(zip, 'ppt/slideLayouts/slideLayout2.xml')
			assertNonVisualDrawingProperty(xml, { name: 'title' }, 'the nameless placeholder')
			assertNonVisualDrawingProperty(xml, { name: 'subhead' }, 'the named placeholder')
		},
	},
])
