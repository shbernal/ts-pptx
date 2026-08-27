// The ECMA-376 preset colour table (`a:prstClr`), and the two other colour models the read
// path's `resolveColor` reaches: `a:hslClr` resolves, `a:scrgbClr` deliberately does not.
//
// The completeness case below is the one that matters. `presetColorHex` is a hand-checked
// constant table addressed through three spelling rules (`dk`/`lt`/`med` abbreviations, both
// spellings of grey, and case), and the failure mode of any such table is a name that silently
// falls through to `null` -- a colour the reader reports as absent rather than as wrong, which
// no rendering comparison would catch. So the full 190-value `ST_PresetColorVal` enumeration is
// pinned here, in schema order, and every one of them has to resolve.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { presetColorHex, resolveColorElement } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/**
 * Every value of `a:ST_PresetColorVal` (ECMA-376 20.1.10.47), in schema order.
 *
 * Copied from the schema rather than derived from the table under test, on purpose: a list
 * generated from the implementation would agree with it by construction and prove nothing.
 */
const PRESET_COLOR_VAL = [
	'aliceBlue',
	'antiqueWhite',
	'aqua',
	'aquamarine',
	'azure',
	'beige',
	'bisque',
	'black',
	'blanchedAlmond',
	'blue',
	'blueViolet',
	'brown',
	'burlyWood',
	'cadetBlue',
	'chartreuse',
	'chocolate',
	'coral',
	'cornflowerBlue',
	'cornsilk',
	'crimson',
	'cyan',
	'darkBlue',
	'darkCyan',
	'darkGoldenrod',
	'darkGray',
	'darkGrey',
	'darkGreen',
	'darkKhaki',
	'darkMagenta',
	'darkOliveGreen',
	'darkOrange',
	'darkOrchid',
	'darkRed',
	'darkSalmon',
	'darkSeaGreen',
	'darkSlateBlue',
	'darkSlateGray',
	'darkSlateGrey',
	'darkTurquoise',
	'darkViolet',
	'dkBlue',
	'dkCyan',
	'dkGoldenrod',
	'dkGray',
	'dkGrey',
	'dkGreen',
	'dkKhaki',
	'dkMagenta',
	'dkOliveGreen',
	'dkOrange',
	'dkOrchid',
	'dkRed',
	'dkSalmon',
	'dkSeaGreen',
	'dkSlateBlue',
	'dkSlateGray',
	'dkSlateGrey',
	'dkTurquoise',
	'dkViolet',
	'deepPink',
	'deepSkyBlue',
	'dimGray',
	'dimGrey',
	'dodgerBlue',
	'firebrick',
	'floralWhite',
	'forestGreen',
	'fuchsia',
	'gainsboro',
	'ghostWhite',
	'gold',
	'goldenrod',
	'gray',
	'grey',
	'green',
	'greenYellow',
	'honeydew',
	'hotPink',
	'indianRed',
	'indigo',
	'ivory',
	'khaki',
	'lavender',
	'lavenderBlush',
	'lawnGreen',
	'lemonChiffon',
	'lightBlue',
	'lightCoral',
	'lightCyan',
	'lightGoldenrodYellow',
	'lightGray',
	'lightGrey',
	'lightGreen',
	'lightPink',
	'lightSalmon',
	'lightSeaGreen',
	'lightSkyBlue',
	'lightSlateGray',
	'lightSlateGrey',
	'lightSteelBlue',
	'lightYellow',
	'ltBlue',
	'ltCoral',
	'ltCyan',
	'ltGoldenrodYellow',
	'ltGray',
	'ltGrey',
	'ltGreen',
	'ltPink',
	'ltSalmon',
	'ltSeaGreen',
	'ltSkyBlue',
	'ltSlateGray',
	'ltSlateGrey',
	'ltSteelBlue',
	'ltYellow',
	'lime',
	'limeGreen',
	'linen',
	'magenta',
	'maroon',
	'medAquamarine',
	'medBlue',
	'medOrchid',
	'medPurple',
	'medSeaGreen',
	'medSlateBlue',
	'medSpringGreen',
	'medTurquoise',
	'medVioletRed',
	'mediumAquamarine',
	'mediumBlue',
	'mediumOrchid',
	'mediumPurple',
	'mediumSeaGreen',
	'mediumSlateBlue',
	'mediumSpringGreen',
	'mediumTurquoise',
	'mediumVioletRed',
	'midnightBlue',
	'mintCream',
	'mistyRose',
	'moccasin',
	'navajoWhite',
	'navy',
	'oldLace',
	'olive',
	'oliveDrab',
	'orange',
	'orangeRed',
	'orchid',
	'paleGoldenrod',
	'paleGreen',
	'paleTurquoise',
	'paleVioletRed',
	'papayaWhip',
	'peachPuff',
	'peru',
	'pink',
	'plum',
	'powderBlue',
	'purple',
	'red',
	'rosyBrown',
	'royalBlue',
	'saddleBrown',
	'salmon',
	'sandyBrown',
	'seaGreen',
	'seaShell',
	'sienna',
	'silver',
	'skyBlue',
	'slateBlue',
	'slateGray',
	'slateGrey',
	'snow',
	'springGreen',
	'steelBlue',
	'tan',
	'teal',
	'thistle',
	'tomato',
	'turquoise',
	'violet',
	'wheat',
	'white',
	'whiteSmoke',
	'yellow',
	'yellowGreen',
]

/** Empty colour maps: nothing here needs the theme, but `resolveColorElement` takes a context. */
const CTX = { clrMap: new Map(), clrScheme: new Map(), fmtScheme: null, fontScheme: null }

/** Parse one DrawingML colour element from source and resolve it against an empty theme. */
function resolve(xml) {
	const doc = new DOMParser().parseFromString(`<a:wrap xmlns:a="${A_NS}">${xml}</a:wrap>`, 'text/xml')
	const [el] = doc.documentElement.getElementsByTagNameNS(A_NS, '*')
	return resolveColorElement(el ?? null, CTX)
}

describe('presetColorHex -- the ST_PresetColorVal table', () => {
	test('resolves every one of the 190 enumerated names', () => {
		assertEqual(PRESET_COLOR_VAL.length, 190, 'the pinned enumeration is the whole of ST_PresetColorVal')
		const unresolved = PRESET_COLOR_VAL.filter((name) => presetColorHex(name) === null)
		assertEqual(unresolved.join(', '), '', 'every preset name resolves to a hex')
		const malformed = PRESET_COLOR_VAL.filter((name) => !/^[0-9A-F]{6}$/.test(presetColorHex(name)))
		assertEqual(malformed.join(', '), '', 'every resolved value is 6 upper-case hex digits')
	})

	test('an abbreviated name is the long one, and grey is gray', () => {
		for (const [short, long] of [
			['dkSlateBlue', 'darkSlateBlue'],
			['ltGoldenrodYellow', 'lightGoldenrodYellow'],
			['medAquamarine', 'mediumAquamarine'],
		]) {
			assertEqual(presetColorHex(short), presetColorHex(long), `${short} is ${long}`)
		}
		for (const [grey, gray] of [
			['grey', 'gray'],
			['dimGrey', 'dimGray'],
			['ltSlateGrey', 'lightSlateGray'],
		]) {
			assertEqual(presetColorHex(grey), presetColorHex(gray), `${grey} is ${gray}`)
		}
	})

	// The `med` rule needs a guard the other two do not: `medium*` already starts with `med`,
	// so expanding it a second time would look for `mediumium*` and find nothing.
	test('the med rule does not re-expand a name that is already long', () => {
		assertEqual(presetColorHex('mediumSpringGreen'), '00FA9A', 'mediumSpringGreen survives normalization')
	})

	test('spot values, including the ones the CSS table is famous for', () => {
		// `green` is HTML green, not X11 green -- the value that trips a table built from the
		// wrong list, since X11 would make it 00FF00, which is what `lime` is here.
		assertEqual(presetColorHex('green'), '008000', 'green')
		assertEqual(presetColorHex('lime'), '00FF00', 'lime')
		// And `darkGray` is LIGHTER than `gray`, the other inherited oddity.
		assertEqual(presetColorHex('darkGray'), 'A9A9A9', 'darkGray')
		assertEqual(presetColorHex('gray'), '808080', 'gray')
		assertEqual(presetColorHex('black'), '000000', 'black')
		assertEqual(presetColorHex('white'), 'FFFFFF', 'white')
	})

	test('a name outside the enumeration is null, not a guess', () => {
		assertEqual(presetColorHex('rebeccaPurple'), null, 'a CSS4 name ECMA-376 never had')
		assertEqual(presetColorHex('mediumiumBlue'), null, 'the shape a careless med-expansion would produce')
		assertEqual(presetColorHex(''), null, 'empty')
		assertEqual(presetColorHex(null), null, 'absent attribute')
	})

	test('case is not load-bearing -- the input is a file somebody else wrote', () => {
		assertEqual(presetColorHex('BLACK'), '000000', 'upper case')
		assertEqual(presetColorHex('DkSlateBlue'), '483D8B', 'mixed case abbreviation')
	})
})

describe('resolveColorElement -- the colour models beyond srgb/scheme/sys', () => {
	test('a:prstClr resolves through the table, transforms and all', () => {
		const plain = resolve(`<a:prstClr val="ltGray"/>`)
		assert(plain, 'a preset colour resolves')
		assertEqual(plain.hex, 'D3D3D3', 'base hex comes from the preset table')
		assertEqual(plain.effectiveHex, 'D3D3D3', 'no transforms, so effective is the base')

		// The transform children ride along exactly as they do for a scheme colour, which is what
		// lets the theme-preserving flatten path re-emit them.
		const shaded = resolve(`<a:prstClr val="black"><a:alpha val="40000"/></a:prstClr>`)
		assert(shaded, 'a preset colour with a transform resolves')
		assert(Math.abs(shaded.alpha - 0.4) < 1e-9, `alpha 40000 -> 0.4, got ${shaded.alpha}`)
	})

	test('a:hslClr resolves through the same sRGB-HSL conversion the transforms use', () => {
		// Hue 0, saturation 100%, luminance 50% is pure red -- the check that the three units are
		// read as 60000ths of a degree and two thousandths-of-a-percent, not as anything else.
		assertEqual(resolve(`<a:hslClr hue="0" sat="100000" lum="50000"/>`).hex, 'FF0000', 'H0 S100 L50 is red')
		assertEqual(
			resolve(`<a:hslClr hue="7200000" sat="100000" lum="50000"/>`).hex,
			'00FF00',
			'hue 7200000/60000 = 120 degrees is green'
		)
		assertEqual(resolve(`<a:hslClr hue="0" sat="0" lum="0"/>`).hex, '000000', 'no luminance is black')
		assertEqual(resolve(`<a:hslClr hue="0" sat="0" lum="100000"/>`).hex, 'FFFFFF', 'full luminance is white')
	})

	// `a:ST_Percentage` is a union in Transitional: the integer form Office writes and a
	// `%`-suffixed decimal string. Reading only the first dropped a schema-legal value silently.
	test('the percent-suffixed spelling of ST_Percentage is read too', () => {
		assertEqual(resolve(`<a:hslClr hue="0" sat="100%" lum="50%"/>`).hex, 'FF0000', 'sat/lum as 100% / 50%')
	})

	// Deliberate, and pinned so it reads as a decision rather than an oversight: the schema
	// does not say whether an scrgb percentage is linear-light or sRGB-encoded, and the two
	// answers differ by a gamma curve. Reporting no colour is honest; reporting a guessed one
	// is not. `import-slide-preserve.test.js` leans on this too -- it is how that suite builds
	// a theme slot with nothing literal behind it.
	test('a:scrgbClr stays unresolved rather than being guessed at', () => {
		assertEqual(resolve(`<a:scrgbClr r="50000" g="50000" b="50000"/>`), null, 'scrgbClr reports no colour')
	})
})
