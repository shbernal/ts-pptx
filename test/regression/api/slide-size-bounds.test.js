import {
	defineRegressionSuite,
	TsPptx,
	captureDiagnostics,
	build,
	readEntry,
	assert,
	assertEqual,
} from '../../helpers.js'
import { InvalidOptionError } from '../../../dist/node.js'

// `p:sldSz/@cx` and `@cy` are `ST_SlideSizeCoordinate`, bounded 914400 to 51206400 EMU — one to
// fifty-six inches. `defineLayout` checked neither bound, and its six-armed guard chain was every
// arm a `warn`: a definition with two finite, truthy, absurd dimensions produced no diagnostic at
// all and emitted a `sldSz` PowerPoint offers to repair. A negative one reached the file as
// `cx="-4572000"`.
//
// No showcase deck defines a layout, so the byte-identity corpus cannot see any of this. These
// are the evidence.
//
// The policy being applied is `docs/diagnostics.md`'s: a finite out-of-range number clamps and
// warns; a value that is not a number at all throws. `deck-argument-guards.test.js` pins the
// other half of the contract — a numeric string is advice, not an error — so the bound had to be
// a clamp rather than a rejection.

/** The `<p:sldSz .../>` element of a deck built under `layoutName`. */
async function slideSize(defineFn, layoutName) {
	const { zip } = await build((pres) => {
		defineFn(pres)
		pres.layout = layoutName
		pres.addSlide()
	})
	const xml = await readEntry(zip, 'ppt/presentation.xml')
	const match = /<p:sldSz[^>]*\/>/.exec(xml)
	assert(match, 'presentation.xml has a p:sldSz')
	return match[0]
}

defineRegressionSuite('Slide size bounds', [
	{
		name: 'defineLayout throws a TsPptxError rather than a raw TypeError on a non-object',
		fn: async () => {
			// It used to warn `layout/invalid-definition` and then die on the next line's
			// `layout.name`, so the caller got a `TypeError` describing a property access — after
			// a warning describing the very input that could not survive. `docs/errors.md` says
			// every failure this library raises is a `TsPptxError`.
			for (const bad of [undefined, null, 'LAYOUT_WIDE', 42]) {
				let err = null
				try {
					new TsPptx().defineLayout(/** @type {never} */ (bad))
				} catch (ex) {
					err = ex
				}
				assert(
					err instanceof InvalidOptionError,
					`expected an InvalidOptionError for ${String(bad)}; got: ${String(err)}`
				)
				assertEqual(err.code, 'layout/invalid-definition')
			}
		},
	},
	{
		name: 'a slide smaller than the schema minimum clamps to one inch and says so',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(async () =>
				slideSize((pres) => pres.defineLayout({ name: 'Badge', width: 0.5, height: 0.5 }), 'Badge')
			)

			// One per dimension: both are out of range and each is reported where it is used.
			assertEqual(codes.filter((code) => code === 'layout/size-out-of-range').length, 2, 'both dimensions reported')
			assertEqual(result, '<p:sldSz cx="914400" cy="914400"/>')
		},
	},
	{
		name: 'a negative dimension clamps rather than reaching the file',
		fn: async () => {
			// The worst of the old cases: `-5` is truthy and finite, so not one of the six arms
			// fired, and `cx="-4572000"` went into the package.
			const { result, codes } = await captureDiagnostics(async () =>
				slideSize((pres) => pres.defineLayout({ name: 'Inverted', width: -5, height: 7.5 }), 'Inverted')
			)

			assert(codes.includes('layout/size-out-of-range'), 'the negative width is reported')
			assertEqual(result, '<p:sldSz cx="914400" cy="6858000"/>')
		},
	},
	{
		name: 'deriving a layout from the current one is caught by the bound, not by the type',
		fn: async () => {
			// `pptx.presLayout` returns EMU and `defineLayout` reads inches, so spreading one into
			// the other states a width of nine million inches. Every value is finite and nothing
			// used to warn. The fields now document the two directions; this is what stops the
			// deck from being emitted at that size.
			const { result, codes } = await captureDiagnostics(async () =>
				slideSize((pres) => pres.defineLayout({ ...pres.presLayout, name: 'Copy' }), 'Copy')
			)

			assertEqual(codes.filter((code) => code === 'layout/size-out-of-range').length, 2, 'both dimensions reported')
			assertEqual(result, '<p:sldSz cx="51206400" cy="51206400"/>')
		},
	},
	{
		name: 'a layout inside the bounds warns about nothing',
		fn: async () => {
			// The sensitivity check: the assertions above only mean something if a legal
			// definition passes silently and unchanged.
			const { result, codes } = await captureDiagnostics(async () =>
				slideSize((pres) => pres.defineLayout({ name: 'A3', width: 16.5, height: 11.7 }), 'A3')
			)

			assertEqual(codes.length, 0, 'no diagnostics for a legal layout: ' + codes.join(', '))
			assertEqual(result, '<p:sldSz cx="15087600" cy="10698480"/>')
		},
	},
])
