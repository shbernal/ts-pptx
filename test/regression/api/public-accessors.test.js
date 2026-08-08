import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { assert, assertEqual, build, defineRegressionSuite, readEntry } from '../../helpers.js'

// Acceptance: the public read-back accessors on `Presentation` and `Slide` return what was
// put in.
//
// These are API surface with no other test. Everything else in this directory asserts
// through the emitted package, which is the right oracle for anything that emits — but a
// getter wired to the wrong private field emits nothing, so that oracle cannot see it. A
// consumer reading `pres.slides` or `slide.hidden` back would, and today nothing would have
// noticed. Each case below therefore asserts the contract the accessor advertises, and
// where the value also reaches the package (`hidden` -> `show="0"`) it asserts both, so a
// getter and its emitter cannot drift apart.
//
// This is also what closes the `functions` axis back to the point of slack the coverage
// doctrine requires, after the browser lane's coverage was merged in — see
// scripts/coverage-gate.mjs. Coverage is why these were *found*, not why they are asserted:
// the assertions are the ones the accessors' own documentation implies.

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..')

defineRegressionSuite('Public accessors', [
	{
		name: 'version reports the released package version',
		fn: async () => {
			// A hand-maintained constant (`VERSION` in src/presentation.ts) against the
			// manifest it is supposed to track. A release that bumps package.json and forgets
			// the constant ships a library that misreports itself, and this is the only place
			// that would say so.
			const { pres } = await build(() => {})
			const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
			assertEqual(pres.version, manifest.version, 'pres.version should track package.json')
		},
	},
	{
		name: 'rtlMode defaults to false and round-trips',
		fn: async () => {
			const { pres } = await build((p) => {
				assertEqual(p.rtlMode, false, 'rtlMode should default to false')
				p.rtlMode = true
			})
			assertEqual(pres.rtlMode, true, 'rtlMode should read back what was set')
		},
	},
	{
		name: 'slides, masterSlide and slideLayouts expose what the presentation holds',
		fn: async () => {
			let firstSlide
			const { pres } = await build((p) => {
				firstSlide = p.addSlide()
				p.addSlide()
			})

			assertEqual(pres.slides.length, 2, 'slides should list every added slide')
			assert(pres.slides[0] === firstSlide, 'slides[0] should be the object addSlide() returned')

			assert(pres.masterSlide, 'masterSlide should be present')
			// The master is the presentation's own, not one of the content slides.
			assert(!pres.slides.includes(pres.masterSlide), 'masterSlide should not appear in slides')

			// One default layout exists before any `defineSlideMaster` call.
			assert(pres.slideLayouts.length >= 1, 'slideLayouts should include the default layout')
		},
	},
	{
		name: 'slide color and hidden round-trip, and hidden reaches the slide XML',
		fn: async () => {
			const { pres, zip } = await build((p) => {
				const added = p.addSlide()
				assertEqual(added.hidden, false, 'hidden should default to false')
				assertEqual(added.color, undefined, 'color should be unset by default')
				added.color = 'FF0000'
				added.hidden = true
			})

			// Read back through `pres.slides` rather than the local reference: that also
			// asserts the collection hands out the same object that was mutated, not a copy.
			const slide = pres.slides[0]
			assertEqual(slide.color, 'FF0000', 'color should read back what was set')
			assertEqual(slide.hidden, true, 'hidden should read back what was set')

			// The emitted side of the same fact: a hidden slide is `show="0"` on `<p:sld>`.
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('show="0"'), `expected show="0" on a hidden slide; got: ${xml.slice(0, 200)}`)
		},
	},
	{
		name: 'slideNumber round-trips through its setter',
		fn: async () => {
			const { pres } = await build((p) => {
				const added = p.addSlide()
				assertEqual(added.slideNumber, undefined, 'slideNumber should be unset by default')
				added.slideNumber = { x: 1.0, y: 6.5 }
			})

			const slide = pres.slides[0]
			assert(slide.slideNumber, 'slideNumber should read back after being set')
			assertEqual(slide.slideNumber.x, 1.0, 'slideNumber.x')
			assertEqual(slide.slideNumber.y, 6.5, 'slideNumber.y')
		},
	},
	{
		name: 'newAutoPagedSlides is empty until a table pages, then names the slides it made',
		fn: async () => {
			const rows = Array.from({ length: 40 }, (_, i) => [`r${i}c0`, `r${i}c1`])

			const { pres } = await build((p) => {
				p.addSlide().addTable([['a', 'b']], { x: 0.5, y: 0.5, w: 9, colW: [4.5, 4.5] })

				p.addSlide().addTable(rows, {
					x: 0.5,
					y: 0.5,
					w: 9,
					h: 4,
					colW: [4.5, 4.5],
					margin: 0,
					slideMargin: 0,
					autoPage: true,
					fontSize: 12,
				})
			})

			// Slides 1 and 2 are the two the callback added; any continuation slides the
			// autoPage table produced were appended after them.
			const plain = pres.slides[0]
			const paged = pres.slides[1]

			assertEqual(plain.newAutoPagedSlides.length, 0, 'a table that fits should page nothing')
			assert(paged.newAutoPagedSlides.length > 0, 'an overflowing autoPage table should report its new slides')
			// The continuation slides are real slides of this presentation, not copies.
			for (const made of paged.newAutoPagedSlides) {
				assert(pres.slides.includes(made), 'every auto-paged slide should be in pres.slides')
			}
		},
	},
])
