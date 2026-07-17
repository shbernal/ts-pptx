import { Presentation } from '../../dist/read.js'
import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../helpers.js'

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
		name: 'default object names stay unique across the group boundary',
		fn: async () => {
			// Group children are spliced out of the slide's object list, so a name counter derived
			// from that list never advanced past them and the later top-level shape reused the
			// grouped child's `Shape 0` in the Selection Pane.
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1 } }])
					s.addShape('rect', { x: 3, y: 1, w: 1, h: 1 })
					s.addText('Hi', { x: 5, y: 1, w: 1, h: 1 })
					s.addImage({ data: PNG_DATA, x: 7, y: 1, w: 1, h: 1 })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			const names = (xml.match(/<p:cNvPr id="\d+" name="([^"]*)"/g) || []).map((s) => s.match(/name="([^"]*)"/)[1])
			assert(names.length === new Set(names).size, 'expected unique objectNames slide-wide; got: ' + names.join(','))
			assert(
				!warnings.some((w) => /duplicate objectName/.test(w)),
				'expected no duplicate-objectName warning; got: ' + JSON.stringify(warnings)
			)
			// grouped child takes `Shape 0`, so the later top-level shape must take `Shape 1`
			assert(names.includes('Shape 0') && names.includes('Shape 1'), 'expected Shape 0 + Shape 1; got: ' + names)
		},
	},
	{
		name: 'the duplicate-objectName warning sees names inside groups',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			try {
				await build((p) => {
					const s = p.addSlide()
					s.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1, objectName: 'Dupe' } }])
					s.addShape('rect', { x: 3, y: 1, w: 1, h: 1, objectName: 'Dupe' })
				})
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => /duplicate objectName/.test(w) && /Dupe/.test(w)),
				'expected a duplicate warning for a name colliding across the group boundary; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		name: 'identical presentations built in one process get identical group names',
		fn: async () => {
			// A module-global group counter made `Group N` depend on how many groups the *process*
			// had built, so the same deck emitted different bytes on the second build.
			const names = []
			for (let run = 0; run < 3; run++) {
				const { zip } = await build((p) => {
					p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1 } }])
				})
				const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
				names.push((xml.match(/name="(Group \d+)"/) || [])[1])
			}
			assert(
				names.every((n) => n === 'Group 1'),
				'expected every independent presentation to name its first group "Group 1"; got: ' + names.join(',')
			)
		},
	},
	{
		name: 'group names number per slide and inside-out across nesting',
		fn: async () => {
			const { zip } = await build((p) => {
				const s1 = p.addSlide()
				s1.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1 } }]) // Group 1
				s1.addGroup([{ group: { children: [{ rect: { x: 3, y: 1, w: 1, h: 1 } }] } }]) // nested -> Group 2, outer -> Group 3
				p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1 } }]) // slide 2 restarts at Group 1
			})
			const groupNames = (xml) => (xml.match(/name="Group \d+"/g) || []).map((s) => s.match(/"(.*)"/)[1])
			const slide1 = groupNames(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assertEqual(slide1.join(','), 'Group 1,Group 3,Group 2', 'slide 1 group names (outer emitted before nested)')
			assertEqual(groupNames(await readEntry(zip, 'ppt/slides/slide2.xml')).join(','), 'Group 1', 'slide 2 restarts')
		},
	},
	{
		name: 'write -> read round-trip: every object in the tree has a unique id and name',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 0.2, y: 0.2, w: 1, h: 1 })
				s.addGroup([
					{ rect: { x: 1, y: 1, w: 1, h: 1 } },
					{ text: { text: 'Hi', options: { x: 2, y: 1, w: 1, h: 1 } } },
					{ group: { children: [{ rect: { x: 3, y: 1, w: 1, h: 1 } }] } },
				])
				s.addText('After', { x: 5, y: 1, w: 1, h: 1 })
			})
			const [slide] = (await Presentation.load(buf)).slides
			const flatten = (shapes) => shapes.flatMap((sh) => [sh, ...(sh.shapes ? flatten(sh.shapes) : [])])
			const all = flatten(slide.shapes)
			assertEqual(all.length, 7, 'expected 2 top-level + group + 2 children + nested group + its child')
			const names = all.map((sh) => sh.name)
			const ids = all.map((sh) => sh.id)
			assert(
				names.every((n) => n),
				'expected every shape to read back with a name; got: ' + names.join(',')
			)
			assertEqual(new Set(names).size, names.length, 'objectNames unique through a real read: ' + names.join(','))
			assertEqual(new Set(ids).size, ids.length, 'drawing ids unique through a real read: ' + ids.join(','))
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
	{
		name: 'a partial group frame warns and falls back to auto-bounds',
		fn: async () => {
			// A partial frame used to take the shared per-object defaults on the unset axes, emitting
			// `cy="0"` and a `cx` that was silently 75% of the layout width.
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					// rect at (1,1) 2x2in -> bbox off=(914400,914400) ext=(1828800,1828800)
					p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 2, h: 2 } }], { x: 5, y: 2, objectName: 'Partial' })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assert(
				/<a:off x="914400" y="914400"\/><a:ext cx="1828800" cy="1828800"\/>/.test(xml),
				'expected auto-bounds, not the partial frame; got: ' + xml
			)
			assert(
				warnings.some((w) => /addGroup/.test(w) && /Partial/.test(w) && /partial frame/.test(w)),
				'expected a partial-frame warning naming the group; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		name: 'a complete group frame is honored verbatim and warns nothing',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 2, h: 2 } }], { x: 5, y: 2, w: 3, h: 1 })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			// all four given -> used as-is (5,2) 3x1in, and chOff/chExt still track off/ext
			assert(
				/<a:off x="4572000" y="1828800"\/><a:ext cx="2743200" cy="914400"\/><a:chOff x="4572000" y="1828800"\/><a:chExt cx="2743200" cy="914400"\/>/.test(
					xml
				),
				'expected the explicit frame verbatim with identity child space; got: ' + xml
			)
			assert(
				!warnings.some((w) => /partial frame/.test(w)),
				'expected no partial-frame warning; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		name: 'a partial frame on a nested group falls back once, and its parent sizes around the fallback',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					// inner group: partial frame -> auto-bounds of its rect at (3,1) 1x1in
					// outer group: auto-bounds over rect (1,1) 1x1in + the inner group -> (1,1) 3x1in
					p.addSlide().addGroup([
						{ rect: { x: 1, y: 1, w: 1, h: 1 } },
						{ group: { children: [{ rect: { x: 3, y: 1, w: 1, h: 1 } }], options: { w: 9 } } },
					])
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			const partialWarnings = warnings.filter((w) => /partial frame/.test(w))
			assertEqual(
				partialWarnings.length,
				1,
				'expected exactly one partial-frame warning; got: ' + JSON.stringify(warnings)
			)
			// outer bbox must be the inner group's fallback bounds, not its bogus w=9in
			assert(
				/<a:off x="914400" y="914400"\/><a:ext cx="2743200" cy="914400"\/>/.test(xml),
				'expected the parent to size around the resolved inner bounds; got: ' + xml
			)
		},
	},
	{
		name: 'write -> read round-trip: a partial-frame group resolves its children',
		fn: async () => {
			// The degenerate `cy="0"` group this used to emit made every child re-read as `null`
			// through the read path's degenerate-chExt guard.
			const origWarn = console.warn
			console.warn = () => {}
			let buf
			try {
				;({ buf } = await build((p) => {
					p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 2, h: 2 } }], { x: 5, y: 2 })
				}))
			} finally {
				console.warn = origWarn
			}
			const [slide] = (await Presentation.load(buf)).slides
			const [group] = slide.shapes
			assertEqual(group.shapes.length, 1, 'expected the group to have one child')
			const frame = group.shapes[0].absoluteFrame
			assert(frame, 'expected a resolvable absoluteFrame, not null (degenerate chExt)')
			assertEqual(frame.width, 1828800, 'expected the child to keep its 2in width')
			assertEqual(frame.left, 914400, 'expected the child to keep its 1in x — a group frame never moves children')
		},
	},
])
