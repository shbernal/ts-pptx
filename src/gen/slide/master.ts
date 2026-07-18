/**
 * PptxGenJS: slide-master parts
 *
 * The slide master (`slideMaster1.xml`) and its `.rels`, plus the master text
 * styles (`<p:txStyles>`): per-level default values mirroring the built-in
 * Office master, with any `MasterTextStyleProps` overrides layered on top.
 */

import { CRLF, LAYOUT_IDX_SERIES_BASE, XML_DECL } from '../../core-enums.js'
import type {
	MasterBulletProps,
	MasterTextStyleLevel,
	MasterTextStyleProps,
	PresSlideInternal,
	SlideLayoutInternal,
} from '../../core-interfaces.js'
import { createColorElement, encodeXmlEntities, inch2Emu } from '../../gen-utils.js'
import { HUNDREDTHS_PER_POINT, ptToHundredths } from '../../units.js'
import { warn } from '../../log.js'
import { slideObjectRelationsToXml, slideObjectToXml } from './object.js'

// Default per-level values mirroring the built-in Office master (used as the base that
// `MasterTextStyleProps` overrides are layered onto). `bu` describes the level's default
// bullet: 'none' -> <a:buNone/>, undefined -> no bullet element (otherStyle), or a glyph.
interface MasterLevelDefault {
	marL: number // EMU
	indent?: number // EMU (omitted when undefined)
	algn: string // OOXML algn value
	spcBefPct?: number // <a:spcBef> percent (×1000); omitted when undefined
	bu?: 'none' | { char: string; font: string } // default bullet; undefined => emit no bullet element
	sz: number // <a:defRPr@sz>
	font: 'mj' | 'mn' // major (heading) vs minor (body) theme font family
}
const MASTER_TITLE_DEFAULT: MasterLevelDefault = {
	marL: 0,
	algn: 'ctr',
	spcBefPct: 0,
	bu: 'none',
	sz: 4400,
	font: 'mj',
}
const MASTER_BODY_DEFAULTS: MasterLevelDefault[] = [
	{
		marL: 342900,
		indent: -342900,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 3200,
		font: 'mn',
	},
	{
		marL: 742950,
		indent: -285750,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '–', font: 'Arial' },
		sz: 2800,
		font: 'mn',
	},
	{
		marL: 1143000,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2400,
		font: 'mn',
	},
	{
		marL: 1600200,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '–', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2057400,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '»', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2514600,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2971800,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 3429000,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 3886200,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
]
const MASTER_OTHER_DEFAULTS: MasterLevelDefault[] = [
	0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600,
].map((marL) => ({ marL, algn: 'l', sz: 1800, font: 'mn' as const }))

function masterAlignAttr(align: MasterTextStyleLevel['align']): string {
	switch (align) {
		case 'left':
			return 'l'
		case 'right':
			return 'r'
		case 'center':
			return 'ctr'
		case 'justify':
			return 'just'
		default:
			return ''
	}
}

/** Build the bullet element for a master level: caller override wins over the level default. */
function masterBulletXml(
	bulletOverride: boolean | MasterBulletProps | undefined,
	base: MasterLevelDefault['bu']
): string {
	// Explicit override
	if (bulletOverride === false) return '<a:buNone/>'
	if (bulletOverride && typeof bulletOverride === 'object') {
		const font = bulletOverride.fontFace ? `<a:buFont typeface="${encodeXmlEntities(bulletOverride.fontFace)}"/>` : ''
		if (bulletOverride.type === 'number') {
			const type = bulletOverride.numberType || 'arabicPeriod'
			const startAt =
				typeof bulletOverride.numberStartAt === 'number' ? ` startAt="${Math.round(bulletOverride.numberStartAt)}"` : ''
			return `${font}<a:buAutoNum type="${type}"${startAt}/>`
		}
		// character bullet (default)
		const char = bulletOverride.characterCode ? `&#x${bulletOverride.characterCode};` : '•'
		const buFont = bulletOverride.fontFace ? font : '<a:buFont typeface="Arial" pitchFamily="34" charset="0"/>'
		return `${buFont}<a:buChar char="${char}"/>`
	}
	// No override (undefined / true): keep the level's default bullet
	if (base === 'none') return '<a:buNone/>'
	if (base && typeof base === 'object')
		return `<a:buFont typeface="${base.font}" pitchFamily="34" charset="0"/><a:buChar char="${base.char}"/>`
	return '' // otherStyle: no bullet element by default
}

/** Serialize one `<a:lvlNpPr>` from its default, layering an optional caller override. */
function masterLevelXml(levelNum: number, base: MasterLevelDefault, levelOverride: MasterTextStyleLevel = {}): string {
	const marL =
		typeof levelOverride.marginLeft === 'number' && !isNaN(levelOverride.marginLeft)
			? inch2Emu(levelOverride.marginLeft)
			: base.marL
	const indentEmu =
		typeof levelOverride.indent === 'number' && !isNaN(levelOverride.indent)
			? inch2Emu(levelOverride.indent)
			: base.indent
	const algn = (levelOverride.align && masterAlignAttr(levelOverride.align)) || base.algn
	const indentAttr = typeof indentEmu === 'number' ? ` indent="${indentEmu}"` : ''

	let xml = `<a:lvl${levelNum}pPr marL="${marL}"${indentAttr} algn="${algn}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
	if (typeof base.spcBefPct === 'number') xml += `<a:spcBef><a:spcPct val="${base.spcBefPct}"/></a:spcBef>`
	xml += masterBulletXml(levelOverride.bullet, base.bu)

	// defRPr
	let sz = base.sz
	if (typeof levelOverride.fontSize === 'number') {
		if (isNaN(levelOverride.fontSize) || levelOverride.fontSize <= 0)
			warn(
				`master textStyles fontSize "${levelOverride.fontSize}" is invalid; keeping default ${base.sz / HUNDREDTHS_PER_POINT}pt.`
			)
		else sz = ptToHundredths(levelOverride.fontSize)
	}
	const boldAttr = levelOverride.bold ? ' b="1"' : ''
	const italicAttr = levelOverride.italic ? ' i="1"' : ''
	const colorXml = levelOverride.color ? createColorElement(levelOverride.color) : '<a:schemeClr val="tx1"/>'
	const latinXml = levelOverride.fontFace
		? `<a:latin typeface="${encodeXmlEntities(levelOverride.fontFace)}"/>`
		: `<a:latin typeface="+${base.font}-lt"/>`
	xml += `<a:defRPr sz="${sz}"${boldAttr}${italicAttr} kern="1200"><a:solidFill>${colorXml}</a:solidFill>${latinXml}<a:ea typeface="+${base.font}-ea"/><a:cs typeface="+${base.font}-cs"/></a:defRPr>`
	xml += `</a:lvl${levelNum}pPr>`
	return xml
}

/** Clamp a caller-provided per-level override array to the 9 valid list levels, warning on overflow. */
function masterLevelOverrides(levels: MasterTextStyleLevel[] | undefined, group: string): MasterTextStyleLevel[] {
	if (!Array.isArray(levels)) return []
	if (levels.length > 9) warn(`master textStyles.${group} has ${levels.length} levels; only the first 9 are used.`)
	return levels.slice(0, 9)
}

/**
 * Build the `<p:txStyles>` block from caller overrides layered onto the Office master defaults.
 * Only invoked when `defineSlideMaster({ textStyles })` was set; the unconfigured deck keeps the
 * verbatim default literal in `makeXmlMaster` for byte-identical output.
 */
function makeXmlMasterTxStyles(textStyles: MasterTextStyleProps): string {
	const title = masterLevelXml(1, MASTER_TITLE_DEFAULT, textStyles.title)
	const bodyOverrides = masterLevelOverrides(textStyles.body, 'body')
	const body = MASTER_BODY_DEFAULTS.map((base, i) => masterLevelXml(i + 1, base, bodyOverrides[i])).join('')
	const otherOverrides = masterLevelOverrides(textStyles.other, 'other')
	const other = MASTER_OTHER_DEFAULTS.map((base, i) => masterLevelXml(i + 1, base, otherOverrides[i])).join('')
	return (
		'<p:txStyles>' +
		`<p:titleStyle>${title}</p:titleStyle>` +
		`<p:bodyStyle>${body}</p:bodyStyle>` +
		`<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${other}</p:otherStyle>` +
		'</p:txStyles>'
	)
}

/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {PresSlideInternal} slide - slide object that represents master slide layout
 * @param {SlideLayoutInternal[]} layouts - slide layouts
 * @return {string} XML
 */
export function makeXmlMaster(slide: PresSlideInternal, layouts: SlideLayoutInternal[]): string {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	const layoutDefs = layouts.map(
		(_layoutDef, idx) =>
			`<p:sldLayoutId id="${LAYOUT_IDX_SERIES_BASE + idx}" r:id="rId${slide._rels.length + idx + 1}"/>`
	)

	let strXml = XML_DECL + CRLF
	strXml +=
		'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
	strXml += slideObjectToXml(slide)
	strXml +=
		'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
	strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>'
	// CT_HeaderFooter/@sldNum defaults to true (ECMA-376). When a slide-number placeholder is
	// defined on the master we must NOT disable it here, otherwise slides that PowerPoint inserts
	// from this master inherit sldNum="0" and the master slide number disappears.
	strXml += `<p:hf${slide._slideNumberProps ? '' : ' sldNum="0"'} hdr="0" ftr="0" dt="0"/>`
	strXml += slide._txStyles
		? makeXmlMasterTxStyles(slide._txStyles)
		: '<p:txStyles>' +
			' <p:titleStyle>' +
			'  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
			' </p:titleStyle>' +
			' <p:bodyStyle>' +
			'  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
			'  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
			'  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
			'  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
			'  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
			'  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
			'  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
			'  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
			'  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
			' </p:bodyStyle>' +
			' <p:otherStyle>' +
			'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
			'  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
			'  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
			'  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
			'  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
			'  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
			'  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
			'  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
			'  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
			'  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
			' </p:otherStyle>' +
			'</p:txStyles>'
	strXml += '</p:sldMaster>'

	return strXml
}

/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {PresSlideInternal} masterSlide - Slide object
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel(masterSlide: PresSlideInternal, slideLayouts: SlideLayoutInternal[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
		target: `../slideLayouts/slideLayout${idx + 1}.xml`,
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
	}))
	defaultRels.push({
		target: '../theme/theme1.xml',
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
	})

	return slideObjectRelationsToXml(masterSlide, defaultRels)
}
