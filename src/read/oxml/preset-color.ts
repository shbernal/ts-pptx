/**
 * The ECMA-376 preset colour table (`a:prstClr/@val` → literal RGB).
 *
 * `ST_PresetColorVal` (§20.1.10.47) enumerates 190 names, and the standard gives each one an
 * RGB triple. They are the HTML/CSS named colours: every value in the enumeration has the same
 * RGB here as the CSS colour of the same name, and the 190 names collapse to the 140 distinct
 * CSS names below through three mechanical spelling rules the enumeration adds on top —
 *
 *  - abbreviated prefixes: `dk*`/`lt*`/`med*` alongside `dark*`/`light*`/`medium*`
 *    (`dkSlateBlue` is `darkSlateBlue`), and only as a *prefix* — `medium…` is not `mediumium…`;
 *  - both spellings of grey: `dimGrey` is `dimGray`;
 *  - case: the enumeration is camelCase, the table below is folded to lower case.
 *
 * So the table holds one row per colour rather than one per spelling, and {@link presetColorHex}
 * normalizes a name into it. That is the difference between 140 hand-checked constants and 190,
 * and it makes the abbreviation rule something a reader can see rather than something spread
 * across fifty near-duplicate rows.
 *
 * This is spec data, not observed behaviour: the values come from the enumeration's own
 * documentation, and the table is asserted complete against the schema's 190-value list in
 * `test/read/preset-color.test.js` — a name the reader cannot resolve is a hole in the table,
 * and that test is what finds it.
 */

/**
 * Canonical (lower-case, unabbreviated, `gray`-spelled) preset colour name → 6-hex RGB.
 * Reach it through {@link presetColorHex}, which does the normalizing.
 */
const PRESET_COLOR_HEX: Readonly<Record<string, string>> = {
	aliceblue: 'F0F8FF',
	antiquewhite: 'FAEBD7',
	aqua: '00FFFF',
	aquamarine: '7FFFD4',
	azure: 'F0FFFF',
	beige: 'F5F5DC',
	bisque: 'FFE4C4',
	black: '000000',
	blanchedalmond: 'FFEBCD',
	blue: '0000FF',
	blueviolet: '8A2BE2',
	brown: 'A52A2A',
	burlywood: 'DEB887',
	cadetblue: '5F9EA0',
	chartreuse: '7FFF00',
	chocolate: 'D2691E',
	coral: 'FF7F50',
	cornflowerblue: '6495ED',
	cornsilk: 'FFF8DC',
	crimson: 'DC143C',
	cyan: '00FFFF',
	darkblue: '00008B',
	darkcyan: '008B8B',
	darkgoldenrod: 'B8860B',
	darkgray: 'A9A9A9',
	darkgreen: '006400',
	darkkhaki: 'BDB76B',
	darkmagenta: '8B008B',
	darkolivegreen: '556B2F',
	darkorange: 'FF8C00',
	darkorchid: '9932CC',
	darkred: '8B0000',
	darksalmon: 'E9967A',
	darkseagreen: '8FBC8F',
	darkslateblue: '483D8B',
	darkslategray: '2F4F4F',
	darkturquoise: '00CED1',
	darkviolet: '9400D3',
	deeppink: 'FF1493',
	deepskyblue: '00BFFF',
	dimgray: '696969',
	dodgerblue: '1E90FF',
	firebrick: 'B22222',
	floralwhite: 'FFFAF0',
	forestgreen: '228B22',
	fuchsia: 'FF00FF',
	gainsboro: 'DCDCDC',
	ghostwhite: 'F8F8FF',
	gold: 'FFD700',
	goldenrod: 'DAA520',
	gray: '808080',
	green: '008000',
	greenyellow: 'ADFF2F',
	honeydew: 'F0FFF0',
	hotpink: 'FF69B4',
	indianred: 'CD5C5C',
	indigo: '4B0082',
	ivory: 'FFFFF0',
	khaki: 'F0E68C',
	lavender: 'E6E6FA',
	lavenderblush: 'FFF0F5',
	lawngreen: '7CFC00',
	lemonchiffon: 'FFFACD',
	lightblue: 'ADD8E6',
	lightcoral: 'F08080',
	lightcyan: 'E0FFFF',
	lightgoldenrodyellow: 'FAFAD2',
	lightgray: 'D3D3D3',
	lightgreen: '90EE90',
	lightpink: 'FFB6C1',
	lightsalmon: 'FFA07A',
	lightseagreen: '20B2AA',
	lightskyblue: '87CEFA',
	lightslategray: '778899',
	lightsteelblue: 'B0C4DE',
	lightyellow: 'FFFFE0',
	lime: '00FF00',
	limegreen: '32CD32',
	linen: 'FAF0E6',
	magenta: 'FF00FF',
	maroon: '800000',
	mediumaquamarine: '66CDAA',
	mediumblue: '0000CD',
	mediumorchid: 'BA55D3',
	mediumpurple: '9370DB',
	mediumseagreen: '3CB371',
	mediumslateblue: '7B68EE',
	mediumspringgreen: '00FA9A',
	mediumturquoise: '48D1CC',
	mediumvioletred: 'C71585',
	midnightblue: '191970',
	mintcream: 'F5FFFA',
	mistyrose: 'FFE4E1',
	moccasin: 'FFE4B5',
	navajowhite: 'FFDEAD',
	navy: '000080',
	oldlace: 'FDF5E6',
	olive: '808000',
	olivedrab: '6B8E23',
	orange: 'FFA500',
	orangered: 'FF4500',
	orchid: 'DA70D6',
	palegoldenrod: 'EEE8AA',
	palegreen: '98FB98',
	paleturquoise: 'AFEEEE',
	palevioletred: 'DB7093',
	papayawhip: 'FFEFD5',
	peachpuff: 'FFDAB9',
	peru: 'CD853F',
	pink: 'FFC0CB',
	plum: 'DDA0DD',
	powderblue: 'B0E0E6',
	purple: '800080',
	red: 'FF0000',
	rosybrown: 'BC8F8F',
	royalblue: '4169E1',
	saddlebrown: '8B4513',
	salmon: 'FA8072',
	sandybrown: 'F4A460',
	seagreen: '2E8B57',
	seashell: 'FFF5EE',
	sienna: 'A0522D',
	silver: 'C0C0C0',
	skyblue: '87CEEB',
	slateblue: '6A5ACD',
	slategray: '708090',
	snow: 'FFFAFA',
	springgreen: '00FF7F',
	steelblue: '4682B4',
	tan: 'D2B48C',
	teal: '008080',
	thistle: 'D8BFD8',
	tomato: 'FF6347',
	turquoise: '40E0D0',
	violet: 'EE82EE',
	wheat: 'F5DEB3',
	white: 'FFFFFF',
	whitesmoke: 'F5F5F5',
	yellow: 'FFFF00',
	yellowgreen: '9ACD32',
}

/**
 * Resolve an `a:prstClr/@val` preset colour name to its 6-hex RGB, or `null` when the name is
 * not one of the 190 `ST_PresetColorVal` values.
 *
 * Lenient about case for the same reason the rest of the read path is: the input is a file
 * someone else wrote, and a producer that spelled `Black` has still said which colour it meant.
 * A name that is not in the enumeration at all resolves to `null` rather than to a guess, so the
 * caller reports "no colour" instead of painting an invented one.
 * @param name - the raw attribute value, or `null` when the attribute is absent
 */
export function presetColorHex(name: string | null): string | null {
	if (!name) return null
	// Case first, so the abbreviation rules below need one spelling each rather than two. The
	// `med` rule needs the guard and the other two do not: no canonical name starts with `dk` or
	// `lt`, but `medium*` does start with `med`, and expanding it again yields `mediumium*`.
	const canonical = name
		.toLowerCase()
		.replace(/grey/g, 'gray')
		.replace(/^dk/, 'dark')
		.replace(/^lt/, 'light')
		.replace(/^med(?!ium)/, 'medium')
	return PRESET_COLOR_HEX[canonical] ?? null
}
