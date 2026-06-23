import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

// Flat group shapes (upstream-issue-307): slide.addGroup() wraps child objects in a single
// PowerPoint group (<p:grpSp>) with an identity child coordinate space (chOff/chExt == off/ext),
// so children keep their slide-absolute coordinates.
defineRegressionSuite('Group shapes', [
	{
		name: 'addGroup emits one p:grpSp wrapping its children with identity chOff/chExt',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addGroup(
					[
						{ rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' } } },
						{ text: { text: 'Hi', options: { x: 1, y: 1, w: 2, h: 1 } } },
					],
					{ objectName: 'MyGroup' }
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// exactly one group wrapper
			assert((xml.match(/<p:grpSp>/g) || []).length === 1, 'expected exactly one <p:grpSp>; got: ' + xml)
			assert(/name="MyGroup"/.test(xml), 'expected group objectName; got: ' + xml)
			// identity child transform: chOff == off and chExt == ext
			const m = xml.match(
				/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(\d+)" y="(\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/
			)
			assert(m, 'expected group xfrm with off/ext/chOff/chExt; got: ' + xml)
			assert(
				m[1] === m[5] && m[2] === m[6] && m[3] === m[7] && m[4] === m[8],
				'expected identity chOff/chExt == off/ext; got: ' + m.slice(1).join(',')
			)
			// both children rendered inside the group
			assert(/<a:srgbClr val="CC0000"\/>/.test(xml), 'expected rect child inside group; got: ' + xml)
			assert(/<a:t>Hi<\/a:t>/.test(xml), 'expected text child inside group; got: ' + xml)
		},
	},
	{
		name: 'group bounds auto-compute as the bounding box of its children',
		fn: async () => {
			const { zip } = await build((p) => {
				// rect at (1,1) 2x1in; image at (3.5,1) 1x1in -> bbox x=1..4.5in, y=1..2in
				p.addSlide().addGroup([
					{ rect: { x: 1, y: 1, w: 2, h: 1 } },
					{ image: { data: PNG_DATA, x: 3.5, y: 1, w: 1, h: 1 } },
				])
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// 1in == 914400 EMU; bbox off=(914400,914400) ext=(3200400,914400)
			assert(
				/<a:off x="914400" y="914400"\/><a:ext cx="3200400" cy="914400"\/>/.test(xml),
				'expected auto-computed group bounds; got: ' + xml
			)
		},
	},
	{
		name: 'group children get cNvPr ids unique from top-level objects',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 0.2, y: 0.2, w: 1, h: 1 }) // top-level idx 0 -> id 2
				s.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1 } }, { rect: { x: 2, y: 1, w: 1, h: 1 } }]) // group is idx 1 -> id 3; children seeded past length -> ids 4,5
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const ids = (xml.match(/<p:cNvPr id="(\d+)"/g) || []).map((s) => Number(s.match(/"(\d+)"/)[1]))
			const uniq = new Set(ids)
			assert(ids.length === uniq.size, 'expected unique cNvPr ids; got: ' + ids.join(','))
		},
	},
	{
		name: 'unsupported child types are skipped with a warning',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					p.addSlide().addGroup([
						{ rect: { x: 1, y: 1, w: 1, h: 1 } },
						{ table: { rows: [[{ text: 'x' }]] } }, // unsupported in MVP
					])
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assert(/<p:grpSp>/.test(xml), 'expected group still emitted; got: ' + xml)
			assert(!/<a:tbl>/.test(xml), 'table child must be skipped; got: ' + xml)
			assert(
				warnings.some((w) => /addGroup/.test(w) && /table/.test(w)),
				'expected warning about table child; got: ' + JSON.stringify(warnings)
			)
		},
	},
])
