import {
	defineRegressionSuite,
	TsPptx,
	captureDiagnostics,
	assert,
	assertEqual,
	assertIncludes,
} from '../../helpers.js'
import { InvalidOptionError } from '../../../dist/node.js'

// The deck-level argument guards: `addSection`, `addSlide`, `defineLayout`, `defineSlideMaster`
// and `embedFont` all accept an object from an untyped caller, and each one has a branch for the
// shapes TypeScript would have refused. Those branches are the whole of what a JavaScript
// consumer meets when it gets a call wrong, and none of them was exercised.
//
// Two different contracts live here, and the split is deliberate:
//
//   - A guard that WARNS is one the deck can carry on past. What is pinned is the `code` (API,
//     branchable) and the fact that the deck still builds afterwards.
//   - A guard that THROWS is one where carrying on would produce a deck that misrepresents what
//     the caller asked for. Those are pinned by class and `code`, as in `error-taxonomy`.
//
// Messages are free to change and nothing below matches on their wording.

/** A deck with one section already in place, for the section-routing cases. */
function withSection(title = 'Intro') {
	const pres = new TsPptx()
	pres.addSection({ title })
	return pres
}

defineRegressionSuite('Deck argument guards', [
	{
		name: 'addSection without an argument warns and leaves the deck usable',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(async () => {
				const pres = new TsPptx()
				pres.addSection(/** @type {never} */ (undefined))
				pres.addSlide()
				return pres.toBytes()
			})

			assertIncludes(codes, 'section/missing-argument')
			assert(result.byteLength > 0, 'the deck still builds; a bad addSection is ignored, not fatal')
		},
	},
	{
		name: 'addSection `order` inserts at that position rather than appending',
		fn: async () => {
			// The documented use for `order`, and the only path that splices instead of pushing.
			// `order` counts from 1, as the option says: the index used to be spliced raw, so
			// `order: 1` landed second and `order: 0` — falsy — appended with nothing said.
			const pres = withSection('A')
			pres.addSection({ title: 'B' })
			pres.addSection({ title: 'Inserted', order: 1 })

			assertEqual(pres.sections.map((s) => s.title).join(','), 'Inserted,A,B', '`order: 1` is the first position')

			const { codes } = await captureDiagnostics(async () => {
				const zero = withSection('A')
				zero.addSection({ title: 'Zero', order: 0 })
				assertEqual(zero.sections.map((s) => s.title).join(','), 'A,Zero', '`order: 0` names no position')
			})
			assertIncludes(codes, 'section/invalid-order')
		},
	},
	{
		name: 'addSlide naming a section that does not exist warns, and still adds the slide',
		fn: async () => {
			// Dropping the slide would be the worse failure: the caller asked for a slide and a
			// typo in an unrelated option is not a reason not to get one.
			const { result, codes } = await captureDiagnostics(async () => {
				const pres = withSection()
				const slide = pres.addSlide({ sectionTitle: 'No Such Section' })
				return { slide, pres, bytes: await pres.toBytes() }
			})

			assertIncludes(codes, 'slide/section-not-found')
			assert(result.slide, 'the slide is returned regardless')
			assertEqual(result.pres.slides.length, 1, 'and it is on the deck')
			assert(result.bytes.byteLength > 0)
		},
	},
	{
		name: 'once sections are in use, a slide added without one lands in a generated default',
		fn: async () => {
			// PowerPoint has no concept of a loose slide alongside sections, so a slide added
			// without a `sectionTitle` has to be given somewhere to live.
			const pres = withSection()
			pres.addSlide()

			assertEqual(pres.sections.length, 2, 'a default section was appended')
			assertEqual(pres.sections[1].title, 'Default-1', 'named by how many defaults exist')
		},
	},
	{
		name: 'a second loose slide joins the existing default rather than making another',
		fn: async () => {
			// The branch the case above cannot reach: the last section is already a default,
			// so it is reused. One default per run of loose slides, not one per slide.
			const pres = withSection()
			pres.addSlide()
			pres.addSlide()

			assertEqual(pres.slides.length, 2, 'both slides are on the deck')
			assertEqual(pres.sections.length, 2, 'still exactly one generated default')
			assertEqual(
				pres.sections[1].title,
				'Default-1',
				'the second slide joined the existing default rather than making a Default-2'
			)
		},
	},
	{
		name: 'defineLayout warns about a field it cannot use, then coerces what it can',
		fn: async () => {
			// A dimension given as a numeric string is still usable — `Number()` recovers it —
			// so the warning is advice and the layout is defined anyway. A missing `name` is
			// the same: the definition lands, under the key `undefined`.
			const usable = [
				{ width: 10, height: 7 }, // no name
				{ name: 'StringHeight', width: 10, height: '7' },
				{ name: 'StringWidth', width: '10', height: 7 },
			]

			for (const layout of usable) {
				const { result, codes } = await captureDiagnostics(async () => {
					const pres = new TsPptx()
					pres.defineLayout(/** @type {never} */ (layout))
					return pres
				})
				assertEqual(codes.length, 1, `one diagnostic for ${JSON.stringify(layout)}`)
				assertEqual(codes[0], 'layout/invalid-definition')
				assert(result.LAYOUTS[String(layout.name)], `${JSON.stringify(layout)} is still defined`)
			}
		},
	},
	{
		name: 'defineLayout refuses a missing dimension instead of registering a NaN-sized layout',
		fn: async () => {
			// `width`/`height` have nothing to coerce from, so the warning is followed by the
			// unit conversion rejecting the value. What matters is the outcome: no layout is
			// registered that would later size a deck to NaN.
			for (const layout of [
				{ name: 'NoWidth', height: 7 },
				{ name: 'NoHeight', width: 10 },
			]) {
				let err = null
				const { codes } = await captureDiagnostics(async () => {
					try {
						new TsPptx().defineLayout(/** @type {never} */ (layout))
					} catch (ex) {
						err = ex
					}
				})

				assertEqual(codes[0], 'layout/invalid-definition', `warned for ${JSON.stringify(layout)}`)
				assert(err instanceof InvalidOptionError, `expected an InvalidOptionError; got: ${String(err)}`)
				assertEqual(err.code, 'coord/non-finite')
			}
		},
	},
	{
		name: 'a well-formed defineLayout warns about nothing and becomes selectable',
		fn: async () => {
			// The sensitivity check for the case above: these assertions can only mean
			// something if a good definition passes silently.
			const { result, codes } = await captureDiagnostics(async () => {
				const pres = new TsPptx()
				pres.defineLayout({ name: 'A3', width: 16.5, height: 11.7 })
				pres.layout = 'A3'
				return pres
			})

			assertEqual(codes.length, 0, 'a valid layout is silent')
			assertEqual(result.presLayout.name, 'A3')
		},
	},
	{
		name: 'defineSlideMaster without a title throws rather than defining an unaddressable master',
		fn: async () => {
			// `title` is the key `addSlide({ masterTitle })` matches on, so a master without one
			// could never be selected. Defining it anyway would fail later and further away.
			let err = null
			try {
				new TsPptx().defineSlideMaster(/** @type {never} */ ({ background: { color: 'FF0000' } }))
			} catch (ex) {
				err = ex
			}

			assert(err instanceof InvalidOptionError, 'expected an InvalidOptionError; got: ' + String(err))
			assertEqual(err.code, 'master/missing-title')
		},
	},
	{
		name: 'embedFont rejects a `data` string that is not base64',
		fn: async () => {
			// The string form is the one shape whose bytes cannot be trusted on arrival; a
			// Uint8Array either is or is not one. Undecodable input is refused here rather
			// than being embedded as a font PowerPoint will not open.
			let err = null
			try {
				await new TsPptx().embedFont({ data: 'not base64 at all!!', typeface: 'Silkscreen' })
			} catch (ex) {
				err = ex
			}

			assert(err instanceof InvalidOptionError, 'expected an InvalidOptionError; got: ' + String(err))
			assertEqual(err.code, 'font/invalid-base64')
		},
	},
])
