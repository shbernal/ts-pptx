import { PNG_1X1, setDiagnosticHandler, defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

// Object lock flags (upstream-issue-438): user-facing spLocks / picLocks /
// graphicFrameLocks. Each flag maps 1:1 to the OOXML attribute of the same name;
// only flags set to true are emitted, and flags invalid for an element type are
// dropped with a warning rather than silently producing illegal XML.
defineRegressionSuite('Object locks', [
	{
		name: 'shape with no objectLock emits a bare cNvSpPr (output unchanged)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:cNvSpPr\/>/.test(xml), 'expected bare <p:cNvSpPr/>; got: ' + xml)
			assert(!/<a:spLocks\b/.test(xml), 'no objectLock should emit no spLocks; got: ' + xml)
		},
	},
	{
		name: 'textbox lock nests spLocks inside the txBox cNvSpPr',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('hi', { x: 1, y: 1, w: 2, h: 1, isTextBox: true, objectLock: { noResize: true } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<p:cNvSpPr txBox="1"><a:spLocks noResize="1"\/><\/p:cNvSpPr>/.test(xml),
				'expected txBox cNvSpPr wrapping spLocks; got: ' + xml
			)
		},
	},
	{
		name: 'shape objectLock emits spLocks with only the set flags, in canonical order',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('rect', {
					x: 1,
					y: 1,
					w: 2,
					h: 1,
					objectLock: { noMove: true, noRot: true, noResize: false, noChangeShapeType: true },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<p:cNvSpPr><a:spLocks noRot="1" noMove="1" noChangeShapeType="1"\/><\/p:cNvSpPr>/.test(xml),
				'expected spLocks with noRot/noMove/noChangeShapeType only; got: ' + xml
			)
			assert(!/noResize/.test(xml), 'false flags must not be emitted; got: ' + xml)
		},
	},
	{
		name: 'text box objectLock supports noTextEdit (spLocks)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('locked', { x: 1, y: 1, w: 2, h: 1, objectLock: { noTextEdit: true } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:spLocks noTextEdit="1"\/>/.test(xml), 'expected spLocks noTextEdit; got: ' + xml)
		},
	},
	{
		name: 'image defaults to picLocks noChangeAspect="1" (output unchanged)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addImage({ data: PNG_1X1, x: 1, y: 1, w: 2, h: 2 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:picLocks noChangeAspect="1"\/>/.test(xml), 'expected default picLocks noChangeAspect; got: ' + xml)
		},
	},
	{
		name: 'image objectLock merges with the noChangeAspect default and can override it',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addImage({
					data: PNG_1X1,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					objectLock: { noChangeAspect: false, noCrop: true, noMove: true },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:picLocks noMove="1" noCrop="1"\/>/.test(xml),
				'expected picLocks noMove/noCrop without noChangeAspect; got: ' + xml
			)
			assert(!/noChangeAspect/.test(xml), 'noChangeAspect:false must drop the default; got: ' + xml)
		},
	},
	{
		name: 'table defaults to graphicFrameLocks noGrp="1" (output unchanged)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }]], { x: 1, y: 1, w: 4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:graphicFrameLocks noGrp="1"\/>/.test(xml), 'expected default graphicFrameLocks noGrp; got: ' + xml)
		},
	},
	{
		name: 'table objectLock adds graphicFrame-specific flags (noDrilldown)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }]], {
					x: 1,
					y: 1,
					w: 4,
					objectLock: { noSelect: true, noDrilldown: true },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:graphicFrameLocks noGrp="1" noDrilldown="1" noSelect="1"\/>/.test(xml),
				'expected graphicFrameLocks with noGrp default + noDrilldown/noSelect; got: ' + xml
			)
		},
	},
	{
		name: 'flag invalid for the element type is dropped with a warning',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let xml
			try {
				const { zip } = await build((p) => {
					// noCrop is a picLocks-only flag; it is not valid on a shape's spLocks
					p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, objectLock: { noMove: true, noCrop: true } })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				setDiagnosticHandler(null)
			}
			assert(/<a:spLocks noMove="1"\/>/.test(xml), 'expected only the valid flag emitted; got: ' + xml)
			assert(!/noCrop/.test(xml), 'noCrop must not appear on spLocks; got: ' + xml)
			assert(
				warnings.some((w) => /noCrop/.test(w) && /a:spLocks/.test(w)),
				'expected a warning about noCrop; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		// Zoom tiles and 3D models emit a FIXED lock set rather than the caller's — neither
		// `SlideZoomProps` nor `Model3dProps` has an `objectLock` field. Both used to write the
		// element by hand, bypassing `genXmlObjectLock`; they now go through it like every other
		// renderer, which puts the attribute order under the shared `PICTURE_LOCK_ATTRS` table
		// instead of under two literals.
		//
		// Byte-pinned here because the byte-identity corpus emits neither construct, so the gate
		// that would normally catch a reordering is silent on exactly these two elements. The
		// strings below are what both hand-written literals produced.
		name: 'zoom and model-3d picLocks are emitted in the shared table order',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('target', { x: 1, y: 1, w: 3, h: 1 })
				const s = p.addSlide()
				s.addSlideZoom({ target: 1, x: 1, y: 1, w: 2, h: 1.5 })
				s.addModel3d({ data: 'model/gltf-binary;base64,AAAA', x: 4, y: 1, w: 2, h: 2 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			const locks = xml.match(/<a:picLocks[^>]*\/>/g) || []
			assert(locks.length === 2, `expected the zoom tile's locks and the model's; got ${JSON.stringify(locks)}`)
			const PICTURE_SET =
				'noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" ' +
				'noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1"'
			assert(locks[0] === `<a:picLocks ${PICTURE_SET}/>`, 'zoom tile locks; got: ' + locks[0])
			// A 3D model adds `noCrop`: it is reframed by its camera, never by cropping the cached
			// raster. It is last because that is where the table puts it.
			assert(locks[1] === `<a:picLocks ${PICTURE_SET} noCrop="1"/>`, 'model-3d locks; got: ' + locks[1])
		},
	},
])
