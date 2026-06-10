import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

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
				p.addSlide().addImage({ data: PNG_DATA, x: 1, y: 1, w: 2, h: 2 })
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
					data: PNG_DATA,
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
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					// noCrop is a picLocks-only flag; it is not valid on a shape's spLocks
					p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, objectLock: { noMove: true, noCrop: true } })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assert(/<a:spLocks noMove="1"\/>/.test(xml), 'expected only the valid flag emitted; got: ' + xml)
			assert(!/noCrop/.test(xml), 'noCrop must not appear on spLocks; got: ' + xml)
			assert(
				warnings.some((w) => /noCrop/.test(w) && /a:spLocks/.test(w)),
				'expected a warning about noCrop; got: ' + JSON.stringify(warnings)
			)
		},
	},
])
