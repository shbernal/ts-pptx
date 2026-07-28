import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertEqual,
	assertIncludes,
} from '../helpers.js'

// `createHyperlinkRels` walks the text/table-cell tree and mints one slide relationship per
// hyperlink, stamping the resolved `_rId` back onto the hyperlink so the emitter can write
// `r:id`. Its shape-level behaviour is pinned by slide-hyperlinks.test.js (theme colors) and
// the notes/zoom suites; what those never reach is the *table* half — hyperlinks that arrive
// through `addTable`, where the cell's own `options` are re-applied after the recursion loses
// them, and where auto-paging re-runs the whole walk once per emitted slide.
//
// That re-run is the interesting part: a hyperlink already carrying an `_rId` is skipped on
// its second visit, which is right when the second visit is the same slide (a reused cell
// object) and wrong when it is a new slide (a repeated header row), because a relationship
// id only means anything inside the part that declares it. Both directions are pinned below.
//
// Four arms of the walk stay uncovered on purpose, because no caller reaches them:
//   - the bail-out for a plain string/number `text`, and the fall-through for a `text` that
//     is neither array nor object. Every entry point (`addShape`, `addText`, `addTable`)
//     hands the walker an object or an array of them, and both recursive calls are guarded
//     by `Array.isArray`.
//   - the `if (tablecell.options)` skip, because `addTable` gives every cell an `options`
//     before the walk runs (gen/define/table.ts, "ARG1: ensure options exists").
//   - the no-options arm of the recursion into a nested run array. Reaching it needs a text
//     object whose own `.text` is an array of runs, which `TextProps.text` types as
//     `string | number` and the run emitter renders as `[object Object]`. Covering it would
//     pin a defect rather than a behaviour.
// Unreachable by construction per docs/testing.md, so left red rather than fenced.

const relsPath = (n) => `ppt/slides/_rels/slide${n}.xml.rels`

/** Slide numbers present in the package, in order. */
function slideNumbers(zip) {
	return listEntries(zip)
		.map((f) => /^ppt\/slides\/slide(\d+)\.xml$/.exec(f))
		.filter(Boolean)
		.map((m) => Number(m[1]))
		.sort((a, b) => a - b)
}

/** `[{ id, type, target }]` for one slide's relationship part, in document order. */
function relationships(xml) {
	return [...xml.matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"/g)].map((m) => ({
		id: m[1],
		type: m[2].split('/').pop(),
		target: m[3],
	}))
}

/**
 * Build with `console.log` captured, returning `{ lines, error }`. These two cases log a
 * diagnostic while defining and then throw while emitting, so the build never completes;
 * restoring in a `finally` keeps a throwing build from leaving the rest of the suite stubbed.
 * (Same shape as the capture in table-autopage-options.test.js.)
 */
async function captureFailedBuild(buildFn) {
	const orig = console.log
	const lines = []
	console.log = (...args) => lines.push(args.map(String).join(' '))
	let error
	try {
		await build(buildFn)
	} catch (err) {
		error = err
	} finally {
		console.log = orig
	}
	return { lines, error }
}

defineRegressionSuite('Hyperlink relationship registration', [
	{
		name: 'a hyperlink on a table cell registers a rel the cell then references',
		fn: async () => {
			// The row recursion drops each cell's `options` on the way in and re-applies them from
			// the collected `cellOpts` by index, so a cell hyperlink only survives if that index
			// still lines up. Nothing else in the suite puts a hyperlink on a table cell.
			const { zip } = await build((p) => {
				p.addSlide().addTable(
					[[{ text: 'Docs', options: { hyperlink: { url: 'https://cell.example.com' } } }, { text: 'plain' }]],
					{ x: 0.5, y: 0.5, w: 9, colW: [4.5, 4.5] }
				)
			})

			const rels = relationships(await readEntry(zip, relsPath(1)))
			const hyper = rels.filter((r) => r.type === 'hyperlink')
			assertEqual(hyper.length, 1, 'one cell hyperlink should mint exactly one rel')
			assertEqual(hyper[0].target, 'https://cell.example.com', 'the rel should carry the cell URL')

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, `<a:hlinkClick r:id="${hyper[0].id}"`, 'the cell run')
		},
	},
	{
		name: 'the same cell object used in two rows registers its hyperlink once',
		fn: async () => {
			// Reusing one cell object across rows is an ordinary way to build a table in a loop.
			// The second visit finds `_rId` already stamped and the rel already on this slide, so
			// it must add nothing: a second Relationship with the same Id would be an invalid part.
			const cell = { text: 'Docs', options: { hyperlink: { url: 'https://reuse.example.com' } } }
			const { zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[cell, { text: 'one' }],
						[cell, { text: 'two' }],
					],
					{
						x: 0.5,
						y: 0.5,
						w: 9,
						colW: [4.5, 4.5],
					}
				)
			})

			const rels = relationships(await readEntry(zip, relsPath(1)))
			assertEqual(
				rels.filter((r) => r.type === 'hyperlink').length,
				1,
				'a reused cell object should not mint a rel per occurrence'
			)
			const ids = rels.map((r) => r.id)
			assertEqual(ids.length, new Set(ids).size, 'relationship ids must be unique within the part')

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const linked = [...xml.matchAll(/<a:hlinkClick r:id="([^"]+)"/g)].map((m) => m[1])
			assertEqual(linked.length, 2, 'both occurrences of the cell should still be linked')
			assertEqual(new Set(linked).size, 1, 'both should point at the one rel')
		},
	},
	{
		name: 'a repeated header re-registers its hyperlink rels on every auto-paged slide',
		fn: async () => {
			// `autoPageRepeatHeader` re-emits the *same* header cell objects on each overflow
			// slide. Their hyperlinks already carry an `_rId` from slide 1, and the skip-if-stamped
			// path would leave slides 2..n referencing an `r:id` their own rels part never
			// declares — a dangling reference PowerPoint repairs by dropping the link. Both rel
			// kinds are here because they take different arms on the way back out: an external URL
			// keeps its `Target`, an internal jump stores the target slide number instead.
			// Annotated because the header row would otherwise fix the element type to
			// "cell that has `options`", which the plain body cells pushed below do not match.
			/** @type {Array<Array<{ text: string, options?: import('../../dist/node.js').TableCellProps }>>} */
			const rows = [
				[
					{ text: 'Docs', options: { hyperlink: { url: 'https://header.example.com' } } },
					{ text: 'Home', options: { hyperlink: { slide: 1 } } },
				],
			]
			for (let i = 0; i < 24; i++) rows.push([{ text: `A${i}` }, { text: `B${i}` }])

			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, {
					x: 0.5,
					y: 0.5,
					w: 9,
					h: 1.5,
					colW: [4.5, 4.5],
					autoPage: true,
					autoPageRepeatHeader: true,
					fontSize: 14,
				})
			})

			const slides = slideNumbers(zip)
			assert(slides.length > 1, `the table must actually page; got ${slides.length} slide(s)`)

			let firstIds
			for (const n of slides) {
				const rels = relationships(await readEntry(zip, relsPath(n)))
				const external = rels.find((r) => r.target === 'https://header.example.com')
				const internal = rels.find((r) => r.type === 'slide')
				assert(external, `slide ${n} is missing the repeated header's external hyperlink rel`)
				assert(internal, `slide ${n} is missing the repeated header's internal slide rel`)
				assertEqual(internal.target, 'slide1.xml', `slide ${n}'s internal jump should target slide 1`)

				const ids = rels.map((r) => r.id)
				assertEqual(ids.length, new Set(ids).size, `slide ${n} has a duplicate relationship id`)

				// The id is carried over rather than re-minted: it is stamped on the shared
				// hyperlink object, so every page's XML spells the same `r:id`.
				const pageIds = [external.id, internal.id]
				if (!firstIds) firstIds = pageIds
				else assertEqual(pageIds.join(','), firstIds.join(','), `slide ${n} renumbered the header's rels`)

				const xml = await readEntry(zip, `ppt/slides/slide${n}.xml`)
				for (const id of pageIds) assertIncludes(xml, `r:id="${id}"`, `slide ${n}`)
			}
		},
	},
	{
		name: 'a `hyperlink` that is not an object is reported while defining, then refused while emitting',
		fn: async () => {
			// `hyperlink: 'https://…'` is the shape people reach for first. Registration logs and
			// declines to mint a rel; the run emitter then throws on the same input. Pinned
			// together because the log is the only part that names the correct shape, and it is
			// worth knowing it still arrives ahead of the throw.
			const { lines, error } = await captureFailedBuild((p) => {
				p.addSlide().addText('link', { x: 1, y: 1, w: 4, h: 0.5, hyperlink: 'https://not-an-object.example.com' })
			})

			assert(
				lines.some((l) => l.startsWith('ERROR: text `hyperlink` option should be an object')),
				`expected the shape diagnostic; got ${JSON.stringify(lines)}`
			)
			assert(error, 'a non-object hyperlink must not silently produce a deck')
			assertIncludes(error.message, 'should be an object', 'the emitter error')
		},
	},
	{
		name: 'a `hyperlink` with none of url/slide/action is reported while defining, then refused while emitting',
		fn: async () => {
			// An empty (or misspelled-key) hyperlink object has nothing to point a rel at.
			const { lines, error } = await captureFailedBuild((p) => {
				p.addSlide().addText('link', { x: 1, y: 1, w: 4, h: 0.5, hyperlink: {} })
			})

			assert(
				lines.some((l) => l.includes('hyperlink requires either')),
				`expected the missing-target diagnostic; got ${JSON.stringify(lines)}`
			)
			assert(error, 'a targetless hyperlink must not silently produce a deck')
			assertIncludes(error.message, 'requires either', 'the emitter error')
		},
	},
])
