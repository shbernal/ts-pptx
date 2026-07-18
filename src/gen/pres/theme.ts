/**
 * PptxGenJS: `ppt/theme/theme1.xml`
 *
 * Emit the Office theme: the `<a:clrScheme>` (with any caller color overrides
 * over the default Office scheme) plus the fixed font/format schemes.
 */

import { REGEX_HEX_COLOR, XML_DECL } from '../../core-enums.js'
import type { PresentationPropsInternal, ThemeColorScheme } from '../../core-interfaces.js'
import { warn } from '../../log.js'

/**
 * Theme `<a:clrScheme>` slots in OOXML document order, with their default Office color child.
 * `dk1`/`lt1` default to `sysClr` (windowText/window); the rest are `srgbClr`. A user override
 * for any slot is emitted as `<a:srgbClr>` (see `buildThemeClrScheme`).
 */
const THEME_CLR_SCHEME_DEFAULTS: ReadonlyArray<[keyof ThemeColorScheme, string]> = [
	['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'],
	['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'],
	['dk2', '<a:srgbClr val="44546A"/>'],
	['lt2', '<a:srgbClr val="E7E6E6"/>'],
	['accent1', '<a:srgbClr val="4472C4"/>'],
	['accent2', '<a:srgbClr val="ED7D31"/>'],
	['accent3', '<a:srgbClr val="A5A5A5"/>'],
	['accent4', '<a:srgbClr val="FFC000"/>'],
	['accent5', '<a:srgbClr val="5B9BD5"/>'],
	['accent6', '<a:srgbClr val="70AD47"/>'],
	['hlink', '<a:srgbClr val="0563C1"/>'],
	['folHlink', '<a:srgbClr val="954F72"/>'],
]

/**
 * Build the theme `<a:clrScheme>` block, applying any caller-supplied color overrides over the
 * default Office scheme. Invalid (non 6-digit-hex) overrides warn and keep the default rather
 * than emitting a degenerate color.
 * @param {ThemeColorScheme} [scheme] - per-slot hex overrides
 * @return {string} the `<a:clrScheme>...</a:clrScheme>` XML
 */
function buildThemeClrScheme(scheme?: ThemeColorScheme): string {
	const slots = THEME_CLR_SCHEME_DEFAULTS.map(([slot, defaultChild]) => {
		const override = scheme?.[slot]
		let child = defaultChild
		if (typeof override === 'string' && override.length > 0) {
			const hex = override.replace('#', '')
			if (REGEX_HEX_COLOR.test(hex)) child = `<a:srgbClr val="${hex.toUpperCase()}"/>`
			else
				warn(`makeXmlTheme: colorScheme.${slot} "${override}" is not a 6-digit hex color; keeping the Office default.`)
		}
		return `<a:${slot}>${child}</a:${slot}>`
	}).join('')
	return `<a:clrScheme name="Office">${slots}</a:clrScheme>`
}

/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
export function makeXmlTheme(pres: PresentationPropsInternal): string {
	const majorFont = pres.theme?.headFontFace
		? `<a:latin typeface="${pres.theme?.headFontFace}"/>`
		: '<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
	const minorFont = pres.theme?.bodyFontFace
		? `<a:latin typeface="${pres.theme?.bodyFontFace}"/>`
		: '<a:latin typeface="Calibri" panose="020F0502020204030204"/>'
	// East Asian (`<a:ea>`) and complex-script (`<a:cs>`) theme font slots. PowerPoint emits these
	// empty by default and resolves per-script via the `<a:font>` list that follows; setting them
	// lets CJK / complex-script runs fall back to a caller-chosen theme font.
	const majorEa = `<a:ea typeface="${pres.theme?.headFontFaceEA ?? ''}"/>`
	const minorEa = `<a:ea typeface="${pres.theme?.bodyFontFaceEA ?? ''}"/>`
	const majorCs = `<a:cs typeface="${pres.theme?.headFontFaceCS ?? ''}"/>`
	const minorCs = `<a:cs typeface="${pres.theme?.bodyFontFaceCS ?? ''}"/>`
	return `${XML_DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>${buildThemeClrScheme(pres.theme?.colorScheme)}<a:fontScheme name="Office"><a:majorFont>${majorFont}${majorEa}${majorCs}<a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont>${minorFont}${minorEa}${minorCs}<a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>`
}
