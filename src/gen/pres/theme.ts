/**
 * ts-pptx: `ppt/theme/theme1.xml`
 *
 * Emit the Office theme: the `<a:clrScheme>` (with any caller color overrides
 * over the default Office scheme) plus the fixed font/format schemes.
 */

import { REGEX_HEX_COLOR, XML_DECL } from '../../core-enums-internal.js'
import type { ThemeColorScheme } from '../../core-interfaces.js'
import type { PresentationPropsInternal } from '../../types/internal.js'
import { warn } from '../../log.js'
import { el, raw, voidEl, type RawXml } from '../oxml/el.js'

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
				const hex = override.replace('#', '')
				if (REGEX_HEX_COLOR.test(hex)) child = raw(voidEl('a:srgbClr', { val: hex.toUpperCase() }))
				else
					warn(
						`makeXmlTheme: colorScheme.${slot} "${override}" is not a 6-digit hex color; keeping the Office default.`
					)
			}
			return raw(el(`a:${slot}`, null, child))
		})
	)
}

// Per-script theme font fallback tables (CT_TextFont lists), in OOXML document order. Static Office
// defaults — no caller data, no escaping risk — unlike the `<a:latin>`/`<a:ea>`/`<a:cs>` slots below.
const MAJOR_FONT_LIST: ReadonlyArray<readonly [string, string]> = [
	['Jpan', '游ゴシック Light'],
	['Hang', '맑은 고딕'],
	['Hans', '等线 Light'],
	['Hant', '新細明體'],
	['Arab', 'Times New Roman'],
	['Hebr', 'Times New Roman'],
	['Thai', 'Angsana New'],
	['Ethi', 'Nyala'],
	['Beng', 'Vrinda'],
	['Gujr', 'Shruti'],
	['Khmr', 'MoolBoran'],
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
	['Viet', 'Times New Roman'],
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
const MINOR_FONT_LIST: ReadonlyArray<readonly [string, string]> = [
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

function fontListChildren(scriptFaces: ReadonlyArray<readonly [string, string]>): RawXml[] {
	return scriptFaces.map(([script, typeface]) => raw(voidEl('a:font', { script, typeface })))
}

// `<a:fmtScheme>` — fill/line/effect/background style matrices referenced by shape style indices
// (`<p:style>`). Fixed Office boilerplate: no caller data reaches this block, so — unlike the
// `<a:clrScheme>`/`<a:fontScheme>` blocks above — there is no escaping risk to centralize by
// rebuilding it node-by-node. Kept as a raw literal (see app.ts for the same reasoning) rather than
// hand-transcribed into ~90 nested el()/voidEl() calls purely for uniformity.
const FMT_SCHEME_XML =
	'<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>'

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
		el('a:theme', { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main', name: 'Office Theme' }, [
			raw(themeElements),
			raw(voidEl('a:objectDefaults')),
			raw(voidEl('a:extraClrSchemeLst')),
			raw(extLst),
		])
	)
}
