import {
	setDiagnosticHandler,
	build,
	readEntry,
	defineRegressionSuite,
	assertEqual,
	assertIncludes,
	assertNonVisualDrawingProperty,
} from '../../helpers.js'

// The definers write their normalization back onto the options object they are handed -- assigned
// `objectName`s, defaulted `line`/`fontSize`/`margin`/`color`, the resolved `autoPage*` family --
// so each one has to copy first, or that state lands on the CALLER's object and leaks to every
// later use of it.
//
// Reusing a style literal is the ordinary way to give several objects a common look, and it used to
// be enough to corrupt them. `addShape` stamped `objectName` onto the literal, so
//
//   s.addShape('rect', STYLE); s.addShape('ellipse', { ...STYLE })
//
// spread the FIRST shape's name onto the second and every one after it -- three shapes called
// `Shape 0`, colliding in the Selection Pane. `addTable` did the same, and worse: STEP 5 hands the
// options object to every plain string cell as that cell's options, so the cell emitters wrote onto
// the caller's literal too.
//
// `addText` was fixed first and is covered by `test/regression/text/text-definition.test.js`; this
// file is the same contract for the two entry points that still had it, plus the two that never
// did (`addImage`, `addMedia`) so a regression in either is caught here rather than in the field.
//
// Asserted three ways per entry point, matching the text test: the literal itself (still exactly
// what the caller wrote -- the actual contract), the diagnostics (no name collision), and the bytes
// (each object gets its own identity). The last two are what breaks without the first.

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

/** `<p:sp>` blocks on slide 1, in document order. */
async function shapesOn(zip, part = 'ppt/slides/slide1.xml') {
	return (await readEntry(zip, part)).match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
}

// 1x1 PNG (red pixel).
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

defineRegressionSuite('Caller-owned options', [
	{
		// The reported failure, verbatim: three shapes built from one spread literal came out named
		// `['', 'Shape 0', 'Shape 0', 'Shape 0']` with a duplicate-name warning, because the first
		// call wrote its assigned name onto `STYLE` and each `{ ...STYLE }` carried it forward.
		name: 'a reused options literal is left untouched and does not leak between shapes',
		fn: async () => {
			const STYLE = { x: 1, y: 1, w: 2, h: 1 }
			const { zip, warnings } = await buildCapturingLogs((p) => {
				const s = p.addSlide()
				s.addShape('rect', STYLE)
				s.addShape('ellipse', { ...STYLE, y: 2.5 })
				s.addShape('roundRect', { ...STYLE, y: 4 })
			})
			const shapes = await shapesOn(zip)
			assertEqual(shapes.length, 3, 'expected all three shapes')
			assertNonVisualDrawingProperty(shapes[0], { name: 'Shape 0' }, 'the first shape')
			assertNonVisualDrawingProperty(shapes[1], { name: 'Shape 1' }, 'the second shape')
			assertNonVisualDrawingProperty(shapes[2], { name: 'Shape 2' }, 'the third shape')
			assertEqual(warnings.length, 0, `expected no diagnostics; got ${JSON.stringify(warnings)}`)
			assertEqual(
				JSON.stringify(STYLE),
				JSON.stringify({ x: 1, y: 1, w: 2, h: 1 }),
				'the caller-owned style literal after three addShape calls'
			)
		},
	},
	{
		// The other half of what `addShape` wrote back: the `line` defaults, and `_bodyProp`, landed
		// on the literal too, so a caller who gave ONE shape a stroke handed the normalized form of
		// it to every shape built from the same literal afterwards.
		//
		// `shadow` is the deliberate exception. `correctShadowOptions` normalizes in place -- angle
		// rounded, `_alpha` derived, a leading `#` stripped -- and the definer shares the caller's
		// object rather than copying it, following `addTextDefinition` and not `copyChartOptions`.
		// Those are idempotent normalizations of the caller's own values, not one shape's identity
		// carried to the next, and `shape/shared-shadow.test.js` pins the resulting `_alpha` because
		// two shapes sharing one shadow object have to keep emitting the same `<a:effectLst>`.
		// Asserted here so the difference between the two nested objects is on the record.
		name: 'shape line normalization does not write back onto the caller',
		fn: async () => {
			const LINE = { color: '0088CC', width: 3 }
			const SHADOW = { type: 'outer', blur: 6, transparency: 40, color: 'FF0000' }
			const STYLE = { x: 1, y: 1, w: 2, h: 1, line: LINE, shadow: SHADOW }
			const { zip } = await build((p) => {
				p.addSlide().addShape('rect', STYLE)
			})
			assertIncludes((await shapesOn(zip))[0], '<a:srgbClr val="0088CC"/>', 'the requested stroke')
			// `line` is replaced wholesale by the normalized copy, so the caller's own line object
			// keeps the two keys it was written with rather than gaining `type`/`transparency`/`dashType`.
			assertEqual(
				JSON.stringify(LINE),
				JSON.stringify({ color: '0088CC', width: 3 }),
				'the caller-owned line literal after addShape'
			)
			assertEqual(
				JSON.stringify(Object.keys(STYLE)),
				JSON.stringify(['x', 'y', 'w', 'h', 'line', 'shadow']),
				'the caller-owned style literal after addShape'
			)
			assertEqual(SHADOW._alpha, 0.6, 'the derived alpha, stamped onto the shared shadow object by design')
		},
	},
	{
		// `addTable` normalized nine keys onto the caller's literal (`objectName`, `fontSize`,
		// `margin`, `color`, and the five `autoPage*` ones). The `objectName` leak is the visible
		// one -- two tables named `Table 0` -- but `autoPage` is the dangerous one: the auto-pager
		// sets `opt.autoPage = false` once it has shredded the rows, so a literal reused for a
		// second table silently lost its paging.
		name: 'a reused options literal is left untouched and does not leak between tables',
		fn: async () => {
			const STYLE = { x: 0.5, y: 0.5, w: 6, colW: [3, 3] }
			const { zip, warnings } = await buildCapturingLogs((p) => {
				const s = p.addSlide()
				s.addTable([['A1', 'B1']], STYLE)
				s.addTable([['A2', 'B2']], { ...STYLE, y: 3 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const frames = xml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g) || []
			assertEqual(frames.length, 2, 'expected both tables')
			assertNonVisualDrawingProperty(frames[0], { name: 'Table 0' }, 'the first table')
			assertNonVisualDrawingProperty(frames[1], { name: 'Table 1' }, 'the second table')
			assertEqual(warnings.length, 0, `expected no diagnostics; got ${JSON.stringify(warnings)}`)
			assertEqual(
				JSON.stringify(STYLE),
				JSON.stringify({ x: 0.5, y: 0.5, w: 6, colW: [3, 3] }),
				'the caller-owned style literal after two addTable calls'
			)
		},
	},
	{
		// The worst of the table sites: STEP 5 gives every plain string cell the table's options
		// object as that cell's own options, so with no copy the cell emitters wrote onto the
		// caller's literal as well. Identity WITHIN one call is kept on purpose -- all the string
		// cells share the one object the definer now owns -- so this asserts the caller's literal,
		// not the cells. `border` is copied too, because the array form is normalized slot by slot.
		name: 'string cells and border defaults do not write back onto the caller',
		fn: async () => {
			const BORDER = [{ type: 'solid', color: 'FF0000' }, undefined, { type: 'solid' }, undefined]
			const STYLE = { x: 0.5, y: 0.5, w: 6, border: BORDER }
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A1', 'B1']], STYLE)
			})
			assertIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), 'FF0000', 'the table border colour')
			assertEqual(
				JSON.stringify(BORDER),
				JSON.stringify([{ type: 'solid', color: 'FF0000' }, undefined, { type: 'solid' }, undefined]),
				'the caller-owned border array after addTable'
			)
			assertEqual(
				JSON.stringify(Object.keys(STYLE)),
				JSON.stringify(['x', 'y', 'w', 'border']),
				'the caller-owned style literal after addTable'
			)
		},
	},
	{
		// `addImage` and `addMedia` never had the bug -- both build a fresh options object rather
		// than normalizing the caller's -- but nothing pinned that, so it was one refactor away from
		// regressing into the same shape as the other three. Pinned here alongside them.
		name: 'image and media options are left untouched',
		fn: async () => {
			const IMG = { data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 }
			const VID = { type: 'video', data: 'video/mp4;base64,AAAA', x: 3, y: 1, w: 2, h: 1 }
			await build((p) => {
				const s = p.addSlide()
				s.addImage(IMG)
				s.addMedia(VID)
			})
			assertEqual(
				JSON.stringify(IMG),
				JSON.stringify({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 }),
				'the caller-owned addImage literal'
			)
			assertEqual(
				JSON.stringify(VID),
				JSON.stringify({ type: 'video', data: 'video/mp4;base64,AAAA', x: 3, y: 1, w: 2, h: 1 }),
				'the caller-owned addMedia literal'
			)
		},
	},
])
