import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

// Group shapes (upstream-issue-307): slide.addGroup() wraps child objects in a PowerPoint group
// (<p:grpSp>) with an identity child coordinate space (chOff/chExt == off/ext) at every depth, so
// children — including nested groups — keep their slide-absolute coordinates.
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
		name: 'nested group emits a group inside a group with identity chOff/chExt at both levels',
		fn: async () => {
			const { zip } = await build((p) => {
				// outer: rect at (1,1) 2x1in (x 1..3); nested auto-sized group of rect+text at (4,1) 1x1in (x 4..5)
				// -> outer bbox x=1..5in (w=4in), y=1..2in (h=1in)
				p.addSlide().addGroup([
					{ rect: { x: 1, y: 1, w: 2, h: 1 } },
					{
						group: {
							children: [
								{ rect: { x: 4, y: 1, w: 1, h: 1, fill: { color: '00CC00' } } },
								{ text: { text: 'Nested', options: { x: 4, y: 1, w: 1, h: 1 } } },
							],
						},
					},
				])
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// two group wrappers: outer + nested
			assert((xml.match(/<p:grpSp>/g) || []).length === 2, 'expected two <p:grpSp> (outer + nested); got: ' + xml)
			// identity child transform at EVERY group level (3 xfrms: spTree root + outer + nested)
			const xfrms = xml.match(
				/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(\d+)" y="(\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/g
			)
			assert(
				xfrms && xfrms.length === 3,
				'expected three group xfrms with chOff/chExt (root + outer + nested); got: ' + xml
			)
			xfrms.forEach((frag) => {
				const m = frag.match(
					/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(\d+)" y="(\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/
				)
				assert(
					m[1] === m[5] && m[2] === m[6] && m[3] === m[7] && m[4] === m[8],
					'expected identity chOff/chExt == off/ext at each level; got: ' + frag
				)
			})
			// outer auto-bbox includes the nested group's children: off=(914400,914400) ext=(3657600,914400)
			assert(
				/<a:off x="914400" y="914400"\/><a:ext cx="3657600" cy="914400"\/>/.test(xml),
				'expected outer group bbox to include nested children; got: ' + xml
			)
			// nested group's own auto-bbox: off=(3657600,914400) ext=(914400,914400)
			assert(
				/<a:off x="3657600" y="914400"\/><a:ext cx="914400" cy="914400"\/>/.test(xml),
				'expected nested group bbox; got: ' + xml
			)
			// nested children rendered
			assert(/<a:srgbClr val="00CC00"\/>/.test(xml), 'expected nested rect; got: ' + xml)
			assert(/<a:t>Nested<\/a:t>/.test(xml), 'expected nested text; got: ' + xml)
			// all cNvPr ids unique across nesting depth
			const ids = (xml.match(/<p:cNvPr id="(\d+)"/g) || []).map((s) => Number(s.match(/"(\d+)"/)[1]))
			assert(ids.length === new Set(ids).size, 'expected unique cNvPr ids across nesting; got: ' + ids.join(','))
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
