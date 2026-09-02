/**
 * ts-pptx: `ppt/theme/theme1.xml`
 *
 * Emit the Office theme: the `<a:clrScheme>` (with any caller color overrides
 * over the default Office scheme) plus the fixed font/format schemes.
 */

import { XML_DECL } from '../../constants-internal.js'
import { isHexColor, stripHash } from '../../hex-color.js'
import type { ThemeColorScheme } from '../../types/index.js'
import type { PresentationPropsInternal } from '../../types/internal.js'
import { warn } from '../../diagnostics.js'
import { el, raw, voidEl, type RawXml } from '../oxml/el.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { FMT_SCHEME_XML } from '../oxml/fmt-scheme.js'

/**
 * Theme `<a:clrScheme>` slots in OOXML document order, with their default Office color child.
 * `dk1`/`lt1` default to `sysClr` (windowText/window); the rest are `srgbClr`. A user override
 * for any slot is emitted as `<a:srgbClr>` (see `buildThemeClrScheme`).
 */
const THEME_CLR_SCHEME_DEFAULTS: ReadonlyArray<[keyof ThemeColorScheme, RawXml]> = [
	['dk1', raw(voidEl('a:sysClr', { val: 'windowText', lastClr: '000000' }))],
	['lt1', raw(voidEl('a:sysClr', { val: 'window', lastClr: 'FFFFFF' }))],
	['dk2', raw(voidEl('a:srgbClr', { val: '44546A' }))],
	['lt2', raw(voidEl('a:srgbClr', { val: 'E7E6E6' }))],
	['accent1', raw(voidEl('a:srgbClr', { val: '4472C4' }))],
	['accent2', raw(voidEl('a:srgbClr', { val: 'ED7D31' }))],
	['accent3', raw(voidEl('a:srgbClr', { val: 'A5A5A5' }))],
	['accent4', raw(voidEl('a:srgbClr', { val: 'FFC000' }))],
	['accent5', raw(voidEl('a:srgbClr', { val: '5B9BD5' }))],
	['accent6', raw(voidEl('a:srgbClr', { val: '70AD47' }))],
	['hlink', raw(voidEl('a:srgbClr', { val: '0563C1' }))],
	['folHlink', raw(voidEl('a:srgbClr', { val: '954F72' }))],
]

/**
 * Build the theme `<a:clrScheme>` block, applying any caller-supplied color overrides over the
 * default Office scheme. Invalid (non 6-digit-hex) overrides warn and keep the default rather
 * than emitting a degenerate color.
 * @param {ThemeColorScheme} [scheme] - per-slot hex overrides
 * @return {string} the `<a:clrScheme>...</a:clrScheme>` XML
 */
function buildThemeClrScheme(scheme?: ThemeColorScheme): string {
	return el(
		'a:clrScheme',
		{ name: 'Office' },
		THEME_CLR_SCHEME_DEFAULTS.map(([slot, defaultChild]) => {
			const override = scheme?.[slot]
			let child = defaultChild
			if (typeof override === 'string' && override.length > 0) {
				const hex = stripHash(override)
				if (isHexColor(hex)) child = raw(voidEl('a:srgbClr', { val: hex.toUpperCase() }))
				else
					warn(
						'theme/invalid-color-override',
						`makeXmlTheme: colorScheme.${slot} "${override}" is not a 6-digit hex color; keeping the Office default.`
					)
			}
			return raw(el(`a:${slot}`, null, child))
		})
	)
}

/**
 * The per-script theme font fallback table (`a:fontScheme`'s CT_TextFont lists), in OOXML
 * document order. Static Office defaults — no caller data, no escaping risk — unlike the
 * `<a:latin>`/`<a:ea>`/`<a:cs>` slots above.
 *
 * Forty-one of the forty-seven scripts take the same face in the major and minor lists, so the
 * shared table below is the whole thing and the two override maps hold the six-odd rows that
 * actually differ. Written out twice, ninety-four lines said less than these do, and one row
 * mistyped in one copy produces a theme that opens fine and paints a script in the wrong face —
 * a difference nobody reviewing a diff of two 47-row tables would see.
 */
const THEME_FONT_LIST: ReadonlyArray<readonly [string, string]> = [
	['Jpan', '游ゴシック'],
	['Hang', '맑은 고딕'],
	['Hans', '等线'],
	['Hant', '新細明體'],
	['Arab', 'Arial'],
	['Hebr', 'Arial'],
	['Thai', 'Cordia New'],
	['Ethi', 'Nyala'],
	['Beng', 'Vrinda'],
	['Gujr', 'Shruti'],
	['Khmr', 'DaunPenh'],
	['Knda', 'Tunga'],
	['Guru', 'Raavi'],
	['Cans', 'Euphemia'],
	['Cher', 'Plantagenet Cherokee'],
	['Yiii', 'Microsoft Yi Baiti'],
	['Tibt', 'Microsoft Himalaya'],
	['Thaa', 'MV Boli'],
	['Deva', 'Mangal'],
	['Telu', 'Gautami'],
	['Taml', 'Latha'],
	['Syrc', 'Estrangelo Edessa'],
	['Orya', 'Kalinga'],
	['Mlym', 'Kartika'],
	['Laoo', 'DokChampa'],
	['Sinh', 'Iskoola Pota'],
	['Mong', 'Mongolian Baiti'],
	['Viet', 'Arial'],
	['Uigh', 'Microsoft Uighur'],
	['Geor', 'Sylfaen'],
	['Armn', 'Arial'],
	['Bugi', 'Leelawadee UI'],
	['Bopo', 'Microsoft JhengHei'],
	['Java', 'Javanese Text'],
	['Lisu', 'Segoe UI'],
	['Mymr', 'Myanmar Text'],
	['Nkoo', 'Ebrima'],
	['Olck', 'Nirmala UI'],
	['Osma', 'Ebrima'],
	['Phag', 'Phagspa'],
	['Syrn', 'Estrangelo Edessa'],
	['Syrj', 'Estrangelo Edessa'],
	['Syre', 'Estrangelo Edessa'],
	['Sora', 'Nirmala UI'],
	['Tale', 'Microsoft Tai Le'],
	['Talu', 'Microsoft New Tai Lue'],
	['Tfng', 'Ebrima'],
]

/**
 * The seven scripts whose *heading* face differs from the body one. Latin and CJK headings take
 * the Light weight; the scripts with no Light cut take a serif (or, for Khmer, a different face
 * entirely) where the body takes a sans. Every other script in {@link THEME_FONT_LIST} uses one
 * face for both.
 */
const MAJOR_FONT_OVERRIDES: Readonly<Record<string, string>> = {
	Jpan: '游ゴシック Light',
	Hans: '等线 Light',
	Arab: 'Times New Roman',
	Hebr: 'Times New Roman',
	Thai: 'Angsana New',
	Khmr: 'MoolBoran',
	Viet: 'Times New Roman',
}

/** The major (heading) and minor (body) lists, both derived from the one table. */
const MAJOR_FONT_LIST: ReadonlyArray<readonly [string, string]> = THEME_FONT_LIST.map(
	([script, face]) => [script, MAJOR_FONT_OVERRIDES[script] ?? face] as const
)
const MINOR_FONT_LIST = THEME_FONT_LIST

function fontListChildren(scriptFaces: ReadonlyArray<readonly [string, string]>): RawXml[] {
	return scriptFaces.map(([script, typeface]) => raw(voidEl('a:font', { script, typeface })))
}

/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
export function makeXmlTheme(pres: PresentationPropsInternal): string {
	// Theme font faces are caller-supplied strings (`ThemeProps`). el()'s attrs escape by
	// construction: an unescaped `"` or `&` in a font name would close the attribute early and
	// emit non-parseable `theme1.xml`, which PowerPoint reports as a file needing repair.
	const majorLatin = pres.theme?.headFontFace
		? raw(voidEl('a:latin', { typeface: pres.theme.headFontFace }))
		: raw(voidEl('a:latin', { typeface: 'Calibri Light', panose: '020F0302020204030204' }))
	const minorLatin = pres.theme?.bodyFontFace
		? raw(voidEl('a:latin', { typeface: pres.theme.bodyFontFace }))
		: raw(voidEl('a:latin', { typeface: 'Calibri', panose: '020F0502020204030204' }))
	// East Asian (`<a:ea>`) and complex-script (`<a:cs>`) theme font slots. PowerPoint emits these
	// empty by default and resolves per-script via the `<a:font>` list that follows; setting them
	// lets CJK / complex-script runs fall back to a caller-chosen theme font.
	const majorEa = raw(voidEl('a:ea', { typeface: pres.theme?.headFontFaceEA ?? '' }))
	const minorEa = raw(voidEl('a:ea', { typeface: pres.theme?.bodyFontFaceEA ?? '' }))
	const majorCs = raw(voidEl('a:cs', { typeface: pres.theme?.headFontFaceCS ?? '' }))
	const minorCs = raw(voidEl('a:cs', { typeface: pres.theme?.bodyFontFaceCS ?? '' }))

	const fontScheme = el('a:fontScheme', { name: 'Office' }, [
		raw(el('a:majorFont', null, [majorLatin, majorEa, majorCs, ...fontListChildren(MAJOR_FONT_LIST)])),
		raw(el('a:minorFont', null, [minorLatin, minorEa, minorCs, ...fontListChildren(MINOR_FONT_LIST)])),
	])

	const themeElements = el('a:themeElements', null, [
		raw(buildThemeClrScheme(pres.theme?.colorScheme)),
		raw(fontScheme),
		raw(FMT_SCHEME_XML),
	])

	const extLst = el(
		'a:extLst',
		null,
		raw(
			el(
				'a:ext',
				{ uri: '{05A4C25C-085E-4340-85A3-A5531E510DB2}' },
				raw(
					voidEl('thm15:themeFamily', {
						'xmlns:thm15': 'http://schemas.microsoft.com/office/thememl/2012/main',
						name: 'Office Theme',
						id: '{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}',
						vid: '{4A3C46E8-61CC-4603-A589-7422A47A8E4A}',
					})
				)
			)
		)
	)

	return (
		XML_DECL +
		el('a:theme', { 'xmlns:a': OOXML_NS.a, name: 'Office Theme' }, [
			raw(themeElements),
			raw(voidEl('a:objectDefaults')),
			raw(voidEl('a:extraClrSchemeLst')),
			raw(extLst),
		])
	)
}
