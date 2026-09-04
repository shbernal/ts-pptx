import JSZip from 'jszip'
import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'
import { readFixture } from '../../read/corpus.js'

// What wins when an object names a placeholder AND states options of its own.
//
// The rule: a placeholder **supplies** an option the caller left out; it never **imposes** one
// over a stated value. Three states, and the middle one had no spelling: inherit the
// placeholder's box, override it, or override part of it. The placeholder was applied last and
// unconditionally, so `addText('x', { placeholder: 'body', x: 5, y: 3, w: 2, h: 1 })` had all four
// stated values thrown away with no diagnostic — while the same object with a *partial* frame and
// no placeholder warns loudly.
//
// It is PowerPoint's own model, and `placeholder-override.pptx` is the evidence: a layout
// placeholder given a bottom anchor and a 1in left inset, with the slide's placeholder then
// re-anchored to the top, comes back as `<a:bodyPr lIns="914400" anchor="b"/>` on the layout and
// `<a:bodyPr anchor="t"/>` on the slide. The slide states only what it overrides, the inset is
// absent because it was never overridden, and the anchor that applies is the stated one.
//
// The frame was fixed first, because that half had an unambiguous reading. The cases below the
// frame ones are the rest of the same rule.

/**
 * The `<a:off>`/`<a:ext>` of the first shape on slide 1.
 *
 * Read out of the `<p:sp>` block, not the part: `p:grpSpPr` carries the spTree's own
 * all-zero transform ahead of every shape, and matching that instead makes each case here
 * pass or fail for a reason that has nothing to do with the shape.
 */
async function frameOf(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const sp = /<p:sp>[\s\S]*?<\/p:sp>/.exec(xml)
	assert(sp, 'expected a shape on the slide; got: ' + xml)
	const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(sp[0])
	const ext = /<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(sp[0])
	assert(off && ext, 'expected a transform on the shape; got: ' + sp[0])
	return { x: Number(off[1]), y: Number(off[2]), cx: Number(ext[1]), cy: Number(ext[2]) }
}

const EMU = 914400

/** A master whose `body` placeholder sits at (1, 1) and is 8 x 4 inches. */
function withBodyPlaceholder(p) {
	p.defineSlideMaster({
		title: 'PH_FRAME',
		objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 8, h: 4 }, text: '' } }],
	})
	return p.addSlide({ masterTitle: 'PH_FRAME' })
}

defineRegressionSuite('placeholder frame vs the object own coordinates', [
	{
		name: 'an object stating no frame takes the whole placeholder box',
		fn: async () => {
			const { zip } = await build((p) => {
				withBodyPlaceholder(p).addText('inherited', { placeholder: 'body' })
			})
			assertEqual(JSON.stringify(await frameOf(zip)), JSON.stringify({ x: EMU, y: EMU, cx: 8 * EMU, cy: 4 * EMU }))
		},
	},
	{
		name: 'an object stating its own frame keeps every value of it',
		fn: async () => {
			const { zip } = await build((p) => {
				withBodyPlaceholder(p).addText('own coords', { placeholder: 'body', x: 5, y: 3, w: 2, h: 1 })
			})
			assertEqual(
				JSON.stringify(await frameOf(zip)),
				JSON.stringify({ x: 5 * EMU, y: 3 * EMU, cx: 2 * EMU, cy: 1 * EMU }),
				'all four stated values survive'
			)
		},
	},
	{
		name: 'a partial frame overrides only the axes it states',
		fn: async () => {
			// The state that had no spelling at all: neither "inherit" nor "override" describes
			// an object that states an `x` and nothing else.
			const { zip } = await build((p) => {
				withBodyPlaceholder(p).addText('half', { placeholder: 'body', x: 5, w: 2 })
			})
			assertEqual(
				JSON.stringify(await frameOf(zip)),
				JSON.stringify({ x: 5 * EMU, y: EMU, cx: 2 * EMU, cy: 4 * EMU }),
				'x and w are the object own, y and h come from the placeholder'
			)
		},
	},
	{
		name: 'a zero coordinate is a stated coordinate',
		fn: async () => {
			// `0` is the value a truthiness test loses, on both sides of this decision.
			const { zip } = await build((p) => {
				withBodyPlaceholder(p).addText('origin', { placeholder: 'body', x: 0, y: 0 })
			})
			const frame = await frameOf(zip)
			assertEqual(frame.x, 0, 'x: 0 is not "said nothing"')
			assertEqual(frame.y, 0, 'nor is y: 0')
			assertEqual(frame.cx, 8 * EMU, 'and the unstated width still inherits')
		},
	},
	{
		name: "a stated valign beats the placeholder's",
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PH_VALIGN',
					objects: [
						{
							placeholder: {
								options: { name: 'body', type: 'body', x: 1, y: 1, w: 8, h: 4, valign: 'bottom' },
								text: '',
							},
						},
					],
				})
				p.addSlide({ masterTitle: 'PH_VALIGN' }).addText('own anchor', { placeholder: 'body', valign: 'top' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:bodyPr[^>]*anchor="t"/.test(xml), "the caller's valign wins; got: " + xml)
		},
	},
	{
		name: 'and an unstated one still comes from it',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PH_VALIGN2',
					objects: [
						{
							placeholder: {
								options: { name: 'body', type: 'body', x: 1, y: 1, w: 8, h: 4, valign: 'bottom' },
								text: '',
							},
						},
					],
				})
				p.addSlide({ masterTitle: 'PH_VALIGN2' }).addText('inherited anchor', { placeholder: 'body' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:bodyPr[^>]*anchor="b"/.test(xml), 'the placeholder supplies it; got: ' + xml)
		},
	},
	{
		// The trap in reading "the caller stated nothing" off the options bag: `bullet` is
		// defaulted to `false` for a placeholder-targeting object BEFORE the inheritance runs, so a
		// presence test would let that default beat the layout's bullet. Letting a default win is
		// not the same statement as letting a caller win.
		name: 'a default does not count as the caller stating something',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PH_BULLET',
					objects: [
						{
							placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 8, h: 4, bullet: true }, text: '' },
						},
					],
				})
				p.addSlide({ masterTitle: 'PH_BULLET' }).addText('bulleted', { placeholder: 'body' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:buChar/.test(xml), "the layout's bullet survives the internal default; got: " + xml)
		},
	},
	{
		name: 'but the caller stating it explicitly does',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PH_BULLET2',
					objects: [
						{
							placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 8, h: 4, bullet: true }, text: '' },
						},
					],
				})
				p.addSlide({ masterTitle: 'PH_BULLET2' }).addText('plain', { placeholder: 'body', bullet: false })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/<a:buChar/.test(xml), 'an explicit `bullet: false` suppresses it; got: ' + xml)
		},
	},
	{
		// The oracle, read rather than cited: a comment claiming what PowerPoint writes goes stale
		// silently, and the fixture is evidence only for as long as it still says what it is said to
		// say. The layout states three things and the slide overrides two of them; what makes the
		// point is the `lIns` that is NOT on the slide.
		name: 'the PowerPoint-authored oracle states only what the slide overrides',
		fn: async () => {
			const zip = await JSZip.loadAsync(await readFixture('placeholder-override'))
			const bodyPrOf = async (part) => {
				const xml = await zip.file(part).async('string')
				const sp = (xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []).find((block) => /<p:ph idx="1"/.test(block))
				assert(sp, `${part} has no body placeholder`)
				return /<a:bodyPr[^>]*\/?>/.exec(sp)[0]
			}
			assertEqual(
				await bodyPrOf('ppt/slideLayouts/slideLayout2.xml'),
				'<a:bodyPr lIns="914400" anchor="b"/>',
				'the layout states the inset and the anchor'
			)
			assertEqual(
				await bodyPrOf('ppt/slides/slide1.xml'),
				'<a:bodyPr anchor="t"/>',
				'the slide states only the anchor it overrode; the inset is absent and inherits'
			)
		},
	},
	{
		name: 'a negative extent normalizes against the frame that actually won',
		fn: async () => {
			// The override ran after the negative-extent normalization, so a placeholder's box
			// skipped it while the flip flags were derived from the object's own signs -- the two
			// halves of one decision computed from different numbers.
			const { zip } = await build((p) => {
				withBodyPlaceholder(p).addText('mirrored', { placeholder: 'body', x: 6, y: 3, w: -2, h: 1 })
			})
			const frame = await frameOf(zip)
			assertEqual(frame.x, 4 * EMU, 'a negative width moves the origin to the min corner')
			assertEqual(frame.cx, 2 * EMU, 'and the extent is positive, as ST_PositiveCoordinate requires')
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/flipH="1"/.test(xml), 'and the mirror is carried as a flip; got: ' + xml)
		},
	},
])
