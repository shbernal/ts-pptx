import JSZip from 'jszip'
import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'
import { readFixture } from '../../read/corpus.js'

// Acceptance: `shadow` on a SHAPE's options is the shape's shadow, and `shadow` on a RUN's options
// is the text shadow. They are two different effects in two different places, and one option bag
// must never produce both.
//
// PowerPoint treats them as two separate gestures -- Shape Effects vs Text Effects -- and writes
// them to two elements: `p:spPr/a:effectLst` and `a:rPr/a:effectLst`. A deck it authored with one
// box per state (`test/read/fixtures/shadow-shape-vs-text.pptx`) carries exactly that: the
// shape-shadow box has the first and not the second, the text-shadow box the second and not the
// first, and only a box given BOTH gestures has both.
//
// `addText('hi', { shadow })` used to emit both, because the bare-string form handed one options
// object to the shape and to its lone run, and `shadow` was on the list of options a run inherits
// from its shape. Two shadows over one string is a state no single PowerPoint action produces.

const SHADOW = { type: 'outer', color: '000000', blur: 3, offset: 2, angle: 45, opacity: 0.5 }

/** Whether the first `<p:sp>` on slide 1 carries an `<a:effectLst>` in its spPr / in a run's rPr. */
async function shadowSites(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const sp = /<p:sp>[\s\S]*?<\/p:sp>/.exec(xml)
	assert(sp, 'expected a shape on the slide; got: ' + xml)
	const spPr = /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(sp[0])
	const rPr = /<a:rPr[\s\S]*?<\/a:rPr>/.exec(sp[0])
	return {
		shape: !!spPr && /<a:effectLst>/.test(spPr[0]),
		run: !!rPr && /<a:effectLst>/.test(rPr[0]),
	}
}

defineRegressionSuite('shadow: the shape and the glyphs are two effects', [
	{
		name: "a shadow on the shape's bag shadows the shape and not the glyphs",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('shape bag', { x: 1, y: 1, w: 4, h: 1, shadow: SHADOW })
			})
			assertEqual(JSON.stringify(await shadowSites(zip)), JSON.stringify({ shape: true, run: false }))
		},
	},
	{
		name: 'the array overload agrees with the string one',
		fn: async () => {
			// The two used to differ for a different reason and now must not differ at all: the run
			// list is where a RUN shadow is stated, and stating none there means none.
			const { zip } = await build((p) => {
				p.addSlide().addText([{ text: 'array' }], { x: 1, y: 1, w: 4, h: 1, shadow: SHADOW })
			})
			assertEqual(JSON.stringify(await shadowSites(zip)), JSON.stringify({ shape: true, run: false }))
		},
	},
	{
		name: "a shadow on a run's bag shadows the glyphs and not the shape",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText([{ text: 'run bag', options: { shadow: SHADOW } }], { x: 1, y: 1, w: 4, h: 1 })
			})
			assertEqual(JSON.stringify(await shadowSites(zip)), JSON.stringify({ shape: false, run: true }))
		},
	},
	{
		name: 'stating both, in the two places, is still expressible',
		fn: async () => {
			// The doubled shadow is not forbidden -- PowerPoint can be driven to it with two user
			// actions. What changed is that it takes two statements here too, as it does there.
			const { zip } = await build((p) => {
				p.addSlide().addText([{ text: 'both', options: { shadow: SHADOW } }], {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					shadow: SHADOW,
				})
			})
			assertEqual(JSON.stringify(await shadowSites(zip)), JSON.stringify({ shape: true, run: true }))
		},
	},
	{
		// The other run-level options still inherit from the shape: the change is about `shadow`
		// alone, not about a run losing what a shape says.
		name: 'the run-level options a run does inherit still reach it',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('inherited', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					bold: true,
					fontSize: 22,
					color: 'FF0000',
					align: 'center',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:rPr[^>]*\bsz="2200"/.test(xml), 'fontSize reaches the run; got: ' + xml)
			assert(/<a:rPr[^>]*\bb="1"/.test(xml), 'bold reaches the run')
			assert(/FF0000/.test(xml), 'color reaches the run')
			assert(/algn="ctr"/.test(xml), 'and align reaches the paragraph')
		},
	},
	{
		// `rtlMode` and `tabStops` are read off a RUN's bag by the paragraph builder and were
		// inherited nowhere: they reached a run only because the bare string form handed it the
		// shape's own object. Once that stopped, they needed an inherit list of their own.
		name: 'the paragraph-level options a run does inherit still reach it',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('rtl', { x: 1, y: 1, w: 4, h: 1, rtlMode: true })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:pPr[^>]*rtl="1"/.test(xml), 'rtlMode reaches the paragraph; got: ' + xml)
		},
	},
	{
		// The oracle, read rather than cited: a comment claiming what PowerPoint writes goes stale
		// silently, and the fixture is only evidence for as long as it still says what it is said
		// to say. Each box was given exactly one gesture (or, for the last, both), so where the
		// `<a:effectLst>` lands IS the answer.
		name: 'the PowerPoint-authored oracle keeps the two effects apart',
		fn: async () => {
			const zip = await JSZip.loadAsync(await readFixture('shadow-shape-vs-text'))
			const xml = await zip.file('ppt/slides/slide1.xml').async('string')
			const sites = {}
			for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []) {
				const name = /name="([^"]+)"/.exec(sp)?.[1]
				const spPr = /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(sp)
				const rPr = /<a:rPr[\s\S]*?<\/a:rPr>|<a:rPr[^>]*\/>/.exec(sp)
				if (name) {
					sites[name] = {
						shape: !!spPr && /<a:effectLst>/.test(spPr[0]),
						run: !!rPr && /<a:effectLst>/.test(rPr[0]),
					}
				}
			}
			assertEqual(
				JSON.stringify(sites),
				JSON.stringify({
					ShapeShadowOnly: { shape: true, run: false },
					TextShadowOnly: { shape: false, run: true },
					BothShadows: { shape: true, run: true },
				}),
				'PowerPoint writes a shape shadow to spPr and a text shadow to rPr, and only both gestures give both'
			)
		},
	},
])
