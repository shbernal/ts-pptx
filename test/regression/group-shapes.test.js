import { Presentation, isGroupShape } from '../../dist/read.js'
import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

// Group shapes: slide.addGroup() wraps child objects in a PowerPoint group (<p:grpSp>) with an
// identity child coordinate space (chOff/chExt == off/ext) at every depth, so children — including
// nested groups — keep their slide-absolute coordinates.

/**
 * The `<p:cNvPr>` id the writer actually emitted for the object named `name`.
 * The cross-boundary reference tests below compare against this rather than a hardcoded number:
 * ids for group children are allocated by a walk in `slideObjectToXml`, and the references to them
 * are resolved up front by `collectSlideShapeIds`, so a hardcoded id would let those two drift apart
 * while every assertion still passed.
 */
const cNvPrIdOf = (xml, name) => {
	const m = xml.match(new RegExp(`<p:cNvPr id="(\\d+)" name="${name}"`))
	return m ? Number(m[1]) : null
}

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
			assert(isGroupShape(group), 'expected the top-level shape to read back as a group')
			assertEqual(group.shapes.length, 1, 'expected the group to have one child')
			const frame = group.shapes[0].absoluteFrame
			assert(frame, 'expected a resolvable absoluteFrame, not null (degenerate chExt)')
			assertEqual(frame.width, 1828800, 'expected the child to keep its 2in width')
			assertEqual(frame.left, 914400, 'expected the child to keep its 1in x — a group frame never moves children')
		},
	},
	{
		name: 'a connector binds to a shape inside a group',
		fn: async () => {
			// Binding used to resolve only against `_slideObjects`, which group children are spliced out
			// of, so this fell back to static endpoints and warned that the shape did not exist.
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, objectName: 'boxInGroup' } }], { objectName: 'Grp' })
					s.addConnector({ type: 'elbow', x1: 3, y1: 1.5, x2: 6, y2: 4.5, startShape: 'boxInGroup', startShapeIdx: 3 })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assertEqual(warnings.length, 0, 'a resolvable binding must not warn; got: ' + JSON.stringify(warnings))
			const childId = cNvPrIdOf(xml, 'boxInGroup')
			assert(childId !== null, 'expected the grouped child to be emitted; got: ' + xml)
			const cxn = (xml.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) || [])[0]
			assert(
				cxn.includes(`<a:stCxn id="${childId}" idx="3"/>`),
				`expected stCxn to point at the grouped child's cNvPr id (${childId}); got: ${cxn}`
			)
		},
	},
	{
		name: 'an animation targets a shape inside a nested group by objectName',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addGroup([
					{ rect: { x: 1, y: 1, w: 1, h: 1 } },
					{ group: { children: [{ rect: { x: 3, y: 1, w: 1, h: 1, objectName: 'deepBox' } }] } },
				])
				s.addAnimation({ preset: 'fadeIn', objectName: 'deepBox' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const deepId = cNvPrIdOf(xml, 'deepBox')
			assert(deepId !== null, 'expected the nested child to be emitted; got: ' + xml)
			const timing = (xml.match(/<p:timing>[\s\S]*<\/p:timing>/) || [])[0]
			assert(timing, 'expected a <p:timing> tree — the effect was dropped; got: ' + xml)
			const spids = [...timing.matchAll(/<p:spTgt spid="(\d+)"\/>/g)].map((m) => Number(m[1]))
			assert(spids.length > 0, 'expected the effect to target a shape; got: ' + timing)
			assert(
				spids.every((spid) => spid === deepId),
				`expected every spid to be the nested child's cNvPr id (${deepId}); got: ${spids.join(',')}`
			)
		},
	},
	{
		name: 'an animation naming no object on the slide warns instead of vanishing',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'real' })
					s.addAnimation({ preset: 'fadeIn', objectName: 'ghost' })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => /no object named "ghost"/.test(w)),
				'expected an unresolved-target warning; got: ' + JSON.stringify(warnings)
			)
			assert(!/<p:spTgt/.test(xml), 'an unresolvable effect must not emit a dangling spid; got: ' + xml)
		},
	},
	{
		name: 'a group child and a top-level object of the same name resolve to the top-level one',
		fn: async () => {
			// Duplicate names are warned about separately; resolution must stay as it was before group
			// children were searched at all, so no existing deck changes its bindings.
			const origWarn = console.warn
			console.warn = () => {}
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'dupe' })
					s.addGroup([{ rect: { x: 4, y: 1, w: 1, h: 1, objectName: 'dupe' } }])
					s.addConnector({ type: 'straight', x1: 2, y1: 1, x2: 4, y2: 1, endShape: 'dupe' })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			const cxn = (xml.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) || [])[0]
			assert(cxn.includes('<a:endCxn id="2" idx="0"/>'), `expected the top-level "dupe" (id 2) to win; got: ${cxn}`)
		},
	},

	// --- write-side group frame/identity options (rotate/flip/lock/altText/empty) ---
	{
		name: 'group rotate/flipH/flipV land on the group xfrm',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1 } }], {
					rotate: 45,
					flipH: true,
					flipV: true,
					objectName: 'Rotated',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// The transform is applied to the whole group: flipH/flipV then rot on the grpSpPr xfrm.
			// 45deg -> 45 * 60000 = 2700000 (convertRotationDegrees). chOff/chExt still track off/ext.
			assert(
				/<p:grpSpPr><a:xfrm flipH="1" flipV="1" rot="2700000"><a:off x="914400" y="914400"\/><a:ext cx="1828800" cy="914400"\/><a:chOff x="914400" y="914400"\/><a:chExt cx="1828800" cy="914400"\/>/.test(
					xml
				),
				'expected flipH/flipV/rot on the group xfrm with identity child space; got: ' + xml
			)
		},
	},
	{
		name: 'group objectLock emits a:grpSpLocks; a flag it does not support warns',
		fn: async () => {
			// A supported flag (noMove) is emitted; an unsupported one (noCrop, valid only on shapes/pics)
			// is dropped with a warning rather than silently coerced.
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1 } }], {
						objectLock: { noMove: true, noResize: true },
						objectName: 'Locked',
					})
					s.addGroup([{ rect: { x: 4, y: 1, w: 1, h: 1 } }], { objectLock: { noCrop: true }, objectName: 'BadLock' })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assert(
				/<p:cNvGrpSpPr><a:grpSpLocks noMove="1" noResize="1"\/><\/p:cNvGrpSpPr>/.test(xml),
				'expected the supported lock flags on a:grpSpLocks; got: ' + xml
			)
			// The unsupported flag yields no grpSpLocks element at all (empty), and warns.
			assert(/name="BadLock"/.test(xml), 'expected the second group emitted; got: ' + xml)
			assert(
				warnings.some((w) => /objectLock\.noCrop is not supported/.test(w) && /a:grpSpLocks/.test(w)),
				'expected a warning for the unsupported lock flag; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		name: 'group altText is written to the group cNvPr descr, entity-encoded',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1 } }], {
					altText: 'Logo & wordmark',
					objectName: 'Described',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<p:cNvPr id="\d+" name="Described" descr="Logo &amp; wordmark"\/>/.test(xml),
				'expected altText on the group cNvPr descr, entity-encoded; got: ' + xml
			)
		},
	},
	{
		name: 'an empty group warns rather than silently emitting a zero-size group',
		fn: async () => {
			// Auto-bounds over no children is a 0x0 box — the degenerate result AGENTS.md says to warn on.
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			let xml
			try {
				const { zip } = await build((p) => {
					p.addSlide().addGroup([], { objectName: 'Empty' })
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => /addGroup/.test(w) && /Empty/.test(w) && /no renderable children/.test(w)),
				'expected an empty-group warning naming the group; got: ' + JSON.stringify(warnings)
			)
			// It still emits the requested (degenerate) group rather than dropping it silently.
			assert(
				/<p:grpSp>[\s\S]*name="Empty"[\s\S]*<a:ext cx="0" cy="0"\/>/.test(xml),
				'expected the empty group still emitted with a zero-size box; got: ' + xml
			)
		},
	},
	{
		name: 'a group whose only children are unsupported kinds warns about both the child and the empty result',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (msg) => warnings.push(String(msg))
			try {
				await build((p) => {
					p.addSlide().addGroup([{ table: { rows: [[{ text: 'x' }]] } }], { objectName: 'AllSkipped' })
				})
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => /addGroup/.test(w) && /table/.test(w)),
				'expected the unsupported-child warning; got: ' + JSON.stringify(warnings)
			)
			assert(
				warnings.some((w) => /no renderable children/.test(w) && /AllSkipped/.test(w)),
				'expected the empty-result warning after every child was skipped; got: ' + JSON.stringify(warnings)
			)
		},
	},

	// --- groupObjects(): group objects already on the slide, addressed by objectName ---
	{
		name: 'groupObjects wraps the named objects in one group, leaving unnamed ones top-level',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' }, objectName: 'Box' })
				s.addText('Hi', { x: 3.5, y: 1, w: 1, h: 1, objectName: 'Caption' })
				s.addText('Loose', { x: 6, y: 1, w: 1, h: 1, objectName: 'Loose' })
				s.groupObjects(['Box', 'Caption'], { objectName: 'Branding' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert((xml.match(/<p:grpSp>/g) || []).length === 1, 'expected exactly one <p:grpSp>; got: ' + xml)
			const grp = xml.match(/<p:grpSp>[\s\S]*?<\/p:grpSp>/)[0]
			assert(/name="Branding"/.test(grp), 'expected the group objectName; got: ' + grp)
			assert(
				/name="Box"/.test(grp) && /name="Caption"/.test(grp),
				'expected both named objects inside the group; got: ' + grp
			)
			assert(!/name="Loose"/.test(grp), 'expected the unnamed object to stay outside the group; got: ' + grp)
			assert(/name="Loose"/.test(xml), 'expected the unnamed object to still be on the slide; got: ' + xml)
			// Frame omitted -> auto-bounds over the members: x=1..4.5in, y=1..2in (same rule as addGroup).
			const m = grp.match(
				/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(\d+)" y="(\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/
			)
			assert(m, 'expected the group xfrm; got: ' + grp)
			assertEqual(m.slice(1, 5).join(','), '914400,914400,3200400,914400', 'auto-bounds over the grouped members')
			assertEqual(m.slice(5, 9).join(','), m.slice(1, 5).join(','), 'identity chOff/chExt == off/ext')
		},
	},
	{
		name: 'groupObjects keeps slide z-order regardless of the order names are passed',
		fn: async () => {
			// Naming order is a selection, not a restack: PowerPoint grouping never reorders shapes,
			// so passing ['Top','Bottom'] must not lift Top above Bottom inside the group.
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Bottom' })
				s.addShape('rect', { x: 1.5, y: 1, w: 1, h: 1, objectName: 'Top' })
				s.groupObjects(['Top', 'Bottom'])
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const order = (xml.match(/name="(Bottom|Top)"/g) || []).map((s) => s.match(/"(.*)"/)[1])
			assertEqual(order.join(','), 'Bottom,Top', 'children must keep their existing z-order, not the naming order')
		},
	},
	{
		name: 'the group takes the topmost member former slot in the slide z-order',
		fn: async () => {
			// Under, [A, C] grouped, Over: the wrapper belongs where C was — above Under and below
			// Over. B sits between the members and must surface above the group, not vanish under it.
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Under' })
				s.addShape('rect', { x: 2, y: 1, w: 1, h: 1, objectName: 'A' })
				s.addShape('rect', { x: 3, y: 1, w: 1, h: 1, objectName: 'B' })
				s.addShape('rect', { x: 4, y: 1, w: 1, h: 1, objectName: 'C' })
				s.addShape('rect', { x: 5, y: 1, w: 1, h: 1, objectName: 'Over' })
				s.groupObjects(['A', 'C'], { objectName: 'Wrapper' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const order = (xml.match(/name="(Under|A|B|C|Over|Wrapper)"/g) || []).map((s) => s.match(/"(.*)"/)[1])
			// Wrapper is emitted before its own children, so it stands in for the A,C pair here.
			assertEqual(order.join(','), 'Under,B,Wrapper,A,C,Over', 'wrapper sits at the topmost member (C) former slot')
		},
	},
	{
		name: 'groupObjects can nest an existing group into a larger logical group',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1 } }], { objectName: 'Inner' })
				s.addText('Label', { x: 3, y: 1, w: 1, h: 1, objectName: 'Label' })
				s.groupObjects(['Inner', 'Label'], { objectName: 'Outer' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const outer = xml.match(/<p:grpSp>[\s\S]*<\/p:grpSp>/)[0]
			assert(
				(xml.match(/<p:grpSp>/g) || []).length === 2,
				'expected the outer group to wrap the inner one; got: ' + xml
			)
			assert(
				outer.indexOf('name="Outer"') < outer.indexOf('name="Inner"'),
				'expected Inner nested inside Outer; got: ' + outer
			)
			// Every id in the tree must still be unique once the wrapper joins the walk.
			const ids = (xml.match(/<p:cNvPr id="(\d+)"/g) || []).map((s) => s.match(/"(\d+)"/)[1])
			assertEqual(
				new Set(ids).size,
				ids.length,
				'expected unique cNvPr ids across the nested tree; got: ' + ids.join(',')
			)
		},
	},
	{
		name: 'groupObjects throws rather than silently leaving an object loose on the slide',
		fn: async () => {
			// Each of these leaves the caller believing an object was grouped when it was not — the
			// exact footgun the throw exists to prevent. The messages must tell the cases apart.
			const grouped = (fn) => {
				const origWarn = console.warn
				console.warn = () => {}
				try {
					return build((p) => fn(p.addSlide()))
				} finally {
					console.warn = origWarn
				}
			}
			const rejects = async (fn, re, label) => {
				let err = null
				try {
					await grouped(fn)
				} catch (ex) {
					err = ex
				}
				assert(err, `expected ${label} to throw`)
				assert(re.test(err.message), `expected ${label} message to match ${re}; got: ${err.message}`)
			}

			await rejects(
				(s) => {
					s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Real' })
					s.groupObjects(['Ghost'])
				},
				/no top-level object on this slide has that objectName/,
				'an unmatched name'
			)
			await rejects(
				(s) => {
					s.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1, objectName: 'Child' } }])
					s.groupObjects(['Child'])
				},
				/already inside a group/,
				'a name that is already grouped'
			)
			await rejects(
				(s) => {
					s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Dupe' })
					s.addShape('rect', { x: 2, y: 1, w: 1, h: 1, objectName: 'Dupe' })
					s.groupObjects(['Dupe'])
				},
				/ambiguous/,
				'an ambiguous name'
			)
			await rejects(
				(s) => {
					s.addTable([[{ text: 'a' }]], { x: 1, y: 1, w: 2, h: 1, objectName: 'Grid' })
					s.groupObjects(['Grid'])
				},
				/grouping a table is not supported yet/,
				'an ungroupable kind'
			)
			await rejects((s) => s.groupObjects([]), /non-empty array/, 'an empty selection')
		},
	},
	{
		name: 'a failed groupObjects leaves the slide exactly as it was',
		fn: async () => {
			// Resolution happens before any move, so a bad name in the list cannot half-group the rest.
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Good' })
				try {
					s.groupObjects(['Good', 'Ghost'])
				} catch {
					/* expected */
				}
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/<p:grpSp>/.test(xml), 'expected no partial group to be emitted; got: ' + xml)
			assert(/name="Good"/.test(xml), 'expected the resolvable object to stay on the slide; got: ' + xml)
		},
	},
	{
		// `objectName` is stored attribute-escaped, so an unescaped lookup key never matches a name
		// containing `&`, `<`, `>`, `"`, `'`, a tab or a newline. groupObjects() therefore threw
		// "no top-level object on this slide has that objectName" for objects plainly on the slide.
		name: 'groupObjects resolves names containing XML metacharacters',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Q&A' })
				s.addText('Hi', { x: 2, y: 1, w: 1, h: 1, objectName: 'Risk <high> "1" \'2\'\ttabbed\nwrapped' })
				s.addText('Loose', { x: 3, y: 1, w: 1, h: 1, objectName: 'Loose' })
				s.groupObjects(['Q&A', 'Risk <high> "1" \'2\'\ttabbed\nwrapped'], { objectName: 'R&D' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert((xml.match(/<p:grpSp>/g) || []).length === 1, 'expected exactly one <p:grpSp>; got: ' + xml)
			const grp = xml.match(/<p:grpSp>[\s\S]*?<\/p:grpSp>/)[0]
			assert(/name="R&amp;D"/.test(grp), 'expected the escaped group objectName; got: ' + grp)
			assert(/name="Q&amp;A"/.test(grp), 'expected the `&` member inside the group; got: ' + grp)
			assert(
				/name="Risk &lt;high&gt; &quot;1&quot; &apos;2&apos;&#9;tabbed&#10;wrapped"/.test(grp),
				'expected the full-metacharacter member inside the group; got: ' + grp
			)
			assert(!/name="Loose"/.test(grp), 'expected the unnamed object to stay outside the group; got: ' + grp)
		},
	},
	{
		// The "already inside a group" hint keys off the same name, so it went stale for the same
		// reason: an escaped-name child was reported as "no top-level object has that objectName",
		// pointing the caller at a typo that is not there. Errors quote the caller's raw spelling.
		name: 'groupObjects tells apart missing and already-grouped for a name with metacharacters',
		fn: async () => {
			const origWarn = console.warn
			console.warn = () => {}
			let err
			try {
				await build((p) => {
					const s = p.addSlide()
					s.addGroup([{ rect: { x: 1, y: 1, w: 1, h: 1, objectName: 'Q&A' } }])
					s.groupObjects(['Q&A'])
				})
			} catch (ex) {
				err = ex
			} finally {
				console.warn = origWarn
			}
			assert(err, 'expected grouping an already-grouped name to throw')
			assert(/already inside a group/.test(err.message), 'expected the already-grouped hint; got: ' + err.message)
			assert(err.message.includes('"Q&A"'), 'expected the error to quote the raw name; got: ' + err.message)
		},
	},
	{
		name: 'write -> read round-trip: grouped-after-the-fact objects read back inside the group',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rect', { x: 1, y: 1, w: 1, h: 1, objectName: 'Box' })
				s.addText('Hi', { x: 2, y: 1, w: 1, h: 1, objectName: 'Caption' })
				s.groupObjects(['Box', 'Caption'], { objectName: 'Branding' })
			})
			const [slide] = (await Presentation.load(buf)).slides
			assertEqual(slide.shapes.length, 1, 'expected a single top-level group after grouping')
			const [group] = slide.shapes
			assert(isGroupShape(group), 'expected the top-level shape to read back as a group')
			assertEqual(group.name, 'Branding', 'group name')
			assertEqual(
				(group.shapes || []).map((sh) => sh.name).join(','),
				'Box,Caption',
				'expected both members to read back as children of the group'
			)
			// (The members' geometry is unchanged by grouping; that is asserted against the emitted
			// xfrm above, since the read model exposes `xfrm()` as a raw element rather than as x/y/w/h.)
		},
	},
])
