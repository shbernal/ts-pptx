import { defineRegressionSuite, build, readEntry, assert, xmlBlocks, firstXmlBlock } from '../helpers.js'

// Regression (dn-negative-extent-normalization): a negative `w`/`h` must never reach
// `<a:ext cx>`/`<a:ext cy>`. Both are ST_PositiveCoordinate, so a negative value is out of range and
// PowerPoint rejects the whole package (0x80070570) without naming the shape — while LibreOffice
// renders it happily, hiding the defect. The writer normalizes to a min-corner origin + absolute
// extent + a flip, the same encoding `addConnector` derives from its endpoints.

const IN = 914400

async function slideXml(buildFn) {
	const { zip } = await build(buildFn)
	return readEntry(zip, 'ppt/slides/slide1.xml')
}

/** Every `<a:ext>`/`<a:chExt>` in the part, as `{ cx, cy }` numbers. */
function extents(xml) {
	return (xml.match(/<a:(?:ch)?ext\b[^>]*\/>/g) || []).map((tag) => ({
		cx: Number(/\bcx="(-?\d+)"/.exec(tag)?.[1]),
		cy: Number(/\bcy="(-?\d+)"/.exec(tag)?.[1]),
	}))
}

function assertNoNegativeExtents(xml) {
	const bad = extents(xml).filter((e) => e.cx < 0 || e.cy < 0)
	assert(bad.length === 0, `expected no negative extents; got ${JSON.stringify(bad)}`)
}

defineRegressionSuite('Negative extent normalization', [
	{
		name: 'a line drawn upward normalizes to a min-corner box plus flipV',
		fn: async () => {
			// The natural expression of "draw from (1,3) to (2.5,1)" is a signed delta.
			const xml = await slideXml((p) => {
				p.addSlide().addShape('line', { x: 1, y: 3, w: 1.5, h: -2, objectName: 'Leader' })
			})
			const sp = xmlBlocks(xml, 'p:sp')[0]
			assert(sp, 'expected a <p:sp>')
			// origin moves to the min corner (y: 3 - 2 = 1in); extent is |h|.
			assert(sp.includes(`<a:off x="${IN}" y="${IN}"/>`), `expected min-corner origin; got: ${sp}`)
			assert(sp.includes(`<a:ext cx="${1.5 * IN}" cy="${2 * IN}"/>`), `expected absolute extent; got: ${sp}`)
			assert(/<a:xfrm flipV="1">/.test(sp), 'an upward line must set flipV')
			assert(!/flipH="1"/.test(sp), 'a rightward line must not set flipH')
			assertNoNegativeExtents(xml)
		},
	},
	{
		name: 'both axes negative flip both axes; a positive box is untouched',
		fn: async () => {
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addShape('line', { x: 4, y: 3, w: -3, h: -2 })
				s.addShape('rect', { x: 1, y: 1, w: 2, h: 1 })
			})
			const sps = xmlBlocks(xml, 'p:sp')
			assert(sps.length === 2, `expected 2 shapes; got ${sps.length}`)
			assert(sps[0].includes(`<a:off x="${IN}" y="${IN}"/>`), `expected origin at (1in,1in); got: ${sps[0]}`)
			assert(sps[0].includes(`<a:ext cx="${3 * IN}" cy="${2 * IN}"/>`), `expected 3x2in extent; got: ${sps[0]}`)
			assert(/<a:xfrm flipH="1" flipV="1">/.test(sps[0]), 'expected both flips, in schema order')
			// A shape with no negative extent keeps its bare <a:xfrm> — normalization is a no-op there.
			assert(sps[1].includes('<a:xfrm>'), `expected an unflipped xfrm; got: ${sps[1]}`)
			assert(sps[1].includes(`<a:ext cx="${2 * IN}" cy="${IN}"/>`), `expected 2x1in extent; got: ${sps[1]}`)
		},
	},
	{
		name: 'an explicit flip and a negative extent cancel rather than double-apply',
		fn: async () => {
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addShape('rightArrow', { x: 4, y: 1, w: -3, h: 1, flipH: true })
				s.addShape('rightArrow', { x: 4, y: 3, w: -3, h: 1 })
			})
			const sps = xmlBlocks(xml, 'p:sp')
			// Same box either way; the flip differs because one is mirrored twice.
			for (const sp of sps) {
				assert(sp.includes(`<a:off x="${IN}"`), `expected origin x at 1in; got: ${sp}`)
				assert(sp.includes(`cx="${3 * IN}"`), `expected 3in extent; got: ${sp}`)
			}
			assert(!/flipH="1"/.test(sps[0]), 'flipH + negative w is not mirrored')
			assert(/<a:xfrm flipH="1">/.test(sps[1]), 'negative w alone is mirrored')
		},
	},
	{
		name: 'negative string coords normalize after unit/percent resolution',
		fn: async () => {
			// '-25%' of a 10in-wide layout is -2.5in; '-1in' is a negative absolute coord.
			const xml = await slideXml((p) => {
				p.defineLayout({ name: 'TEST', width: 10, height: 7.5 })
				p.layout = 'TEST'
				p.addSlide().addShape('rect', { x: '50%', y: 4, w: '-25%', h: '-1in' })
			})
			const sp = xmlBlocks(xml, 'p:sp')[0]
			assert(sp.includes(`<a:off x="${2.5 * IN}" y="${3 * IN}"/>`), `expected min-corner origin; got: ${sp}`)
			assert(sp.includes(`<a:ext cx="${2.5 * IN}" cy="${IN}"/>`), `expected absolute extents; got: ${sp}`)
			assert(/<a:xfrm flipH="1" flipV="1">/.test(sp), 'expected both flips')
			assertNoNegativeExtents(xml)
		},
	},
	{
		name: 'text boxes and images share the normalized placement path',
		fn: async () => {
			// 1x1 transparent PNG
			const png =
				'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addText('up', { x: 1, y: 4, w: 2, h: -3 })
				s.addImage({ data: png, x: 5, y: 4, w: -2, h: -1 })
			})
			const sp = xmlBlocks(xml, 'p:sp')[0]
			assert(sp.includes(`<a:off x="${IN}" y="${IN}"/>`), `expected text origin at (1in,1in); got: ${sp}`)
			assert(sp.includes(`<a:ext cx="${2 * IN}" cy="${3 * IN}"/>`), `expected 2x3in text extent; got: ${sp}`)
			const pic = firstXmlBlock(xml, 'p:pic')
			assert(pic.includes(`<a:off x="${3 * IN}" y="${3 * IN}"/>`), `expected image origin at (3in,3in); got: ${pic}`)
			assert(pic.includes(`<a:ext cx="${2 * IN}" cy="${IN}"/>`), `expected 2x1in image extent; got: ${pic}`)
			assertNoNegativeExtents(xml)
		},
	},
	{
		name: 'group auto-bounds size around a child normalized box, not its signed one',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addGroup([
					{ rect: { x: 1, y: 1, w: 1, h: 1 } },
					// Runs back to (1in, 1in): un-normalized this child would report a maxX left of
					// the group's minX and collapse the auto-bounds box.
					{ line: { x: 4, y: 4, w: -3, h: -3 } },
				])
			})
			const grp = firstXmlBlock(xml, 'p:grpSp')
			const frame = `<a:off x="${IN}" y="${IN}"/><a:ext cx="${3 * IN}" cy="${3 * IN}"/>`
			assert(grp.includes(frame), `expected the group to span (1in,1in)-(4in,4in); got: ${grp}`)
			assertNoNegativeExtents(xml)
		},
	},
])
