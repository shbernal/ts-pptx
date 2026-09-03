import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'

// Which frame wins when an object names a placeholder AND states coordinates of its own.
//
// Three states, and the middle one had no spelling: inherit the placeholder's box, override it,
// or override part of it. The placeholder was applied last and unconditionally, so
// `addText('x', { placeholder: 'body', x: 5, y: 3, w: 2, h: 1 })` had all four stated values
// thrown away with no diagnostic — while the same object with a *partial* frame and no
// placeholder warns loudly. The decision is that an explicit option beats an inherited one, as
// it does everywhere else in this library; a placeholder's frame is an inherited one.

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
