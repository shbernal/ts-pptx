/**
 * PptxGenJS: XML Generation
 *
 * The bulk of the OOXML emitter: turns the in-memory slide/presentation model into
 * the many XML parts of a `.pptx` package (slides, masters, layouts, notes, theme,
 * `[Content_Types].xml`, `.rels`, app/core/custom props, table styles, view props).
 * Every export is a pure string builder — no I/O; the `pptxgen.ts` export flow calls
 * these and hands the strings to the zip writer. Charts live in `gen-charts.ts`.
 *
 * Contents — jump by grepping the `// ===== <region> =====` banners:
 *   - Value clamps & shape/geometry helpers   font/spacing clamps, crop, preset & custom geometry, cell borders
 *   - Slide serialization                     slideObjectToXml (the per-shape spTree builder) + its rels
 *   - Text body generation                    paragraph/run props, runs, math, genXmlTextBody, placeholders
 *   - Package-level parts                     [Content_Types].xml, root rels, app/core/custom props, presentation rels
 *   - Transitions & animations                slide timing, transitions, the p:anim* sequence builders
 *   - Slides, notes & layouts                 makeXmlSlide / notes parts / makeXmlLayout
 *   - Masters & text styles                   master txStyles defaults + makeXmlMaster
 *   - Slide/master rels & comments            per-part .rels and the comment author/thread parts
 *   - Theme, presentation & root files        theme, makeXmlPresentation, presProps, table styles, viewProps
 */

import { CRLF, REGEX_HEX_COLOR, TableStyle, XML_DECL } from './core-enums.js'
import type {
	BorderProps,
	PresentationPropsInternal,
	PresSlideInternal,
	TableStyleInternal,
	TableStyleRegionProps,
	ThemeColorScheme,
} from './core-interfaces.js'
import {
	createColorElement,
	encodeXmlEntities,
	genXmlColorSelection,
	getUuid,
	lineWidthToEmu,
	resolveBorderWidth,
} from './gen-utils.js'
// Notes parts live in gen/slide/notes.ts; re-exported so `genXml.*` keeps resolving in pptxgen.ts.
export {
	buildNotesSlideRels,
	getNotesFromSlide,
	makeXmlNotesMaster,
	makeXmlNotesMasterRel,
	makeXmlNotesSlide,
	makeXmlNotesSlideRel,
} from './gen/slide/notes.js'
// Comment parts live in gen/slide/comments.ts; re-exported for pptxgen.ts's `genXml.*` access.
export { makeXmlCommentAuthors, makeXmlComments, resolveCommentAuthors } from './gen/slide/comments.js'
export type { ResolvedComments } from './gen/slide/comments.js'
// Master parts live in gen/slide/master.ts; re-exported for pptxgen.ts's `genXml.*` access.
export { makeXmlMaster, makeXmlMasterRel } from './gen/slide/master.js'
// Slide + layout parts live in gen/slide/{slide,layout}.ts; re-exported for pptxgen.ts's `genXml.*` access.
export { makeXmlSlide, makeXmlSlideLayoutRel, makeXmlSlideRel } from './gen/slide/slide.js'
export { makeXmlLayout } from './gen/slide/layout.js'
// OPC package parts live in gen/opc/*; re-exported for pptxgen.ts's `genXml.*` access.
export { makeXmlContTypes } from './gen/opc/content-types.js'
export { makeXmlRootRels } from './gen/opc/root-rels.js'
export { makeXmlApp } from './gen/opc/app.js'
export { makeXmlCore } from './gen/opc/core.js'
export { makeXmlCustomProperties } from './gen/opc/custom-props.js'
import { warn } from './log.js'
import { type EmbeddedFont, FONT_REL_TYPE, flattenEmbeddedFaces, serializeEmbeddedFontLst } from './embedded-fonts.js'

// ===== Package-level parts =====

/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {PresSlideInternal[]} slides - Presenation Slides
 * @returns XML
 */
/**
 * The first relationship id free for embedded-font rels in `presentation.xml.rels`,
 * i.e. one past the last fixed rel {@link makeXmlPresentationRels} emits. Shared by
 * the rels writer and {@link makeXmlPresentation} so the `embeddedFontLst` face
 * `r:id`s match the relationships that back them.
 *
 * Layout: rId1 = slideMaster, rId2..(N+1) = N slides, then notesMaster/presProps/
 * viewProps/theme1/tableStyles (5), then commentAuthors (1, only with comments).
 */
function presentationFontRelStart(slides: PresSlideInternal[]): number {
	const hasComments = (slides || []).some((slide) => (slide._comments || []).length > 0)
	return slides.length + 7 + (hasComments ? 1 : 0)
}

export function makeXmlPresentationRels(slides: PresSlideInternal[], embeddedFonts?: EmbeddedFont[]): string {
	let intRelNum = 1
	let strXml = XML_DECL + CRLF
	strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
	strXml +=
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
	for (let idx = 1; idx <= slides.length; idx++) {
		strXml += `<Relationship Id="rId${++intRelNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${idx}.xml"/>`
	}
	intRelNum++
	strXml +=
		`<Relationship Id="rId${intRelNum + 0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>` +
		`<Relationship Id="rId${intRelNum + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>` +
		`<Relationship Id="rId${intRelNum + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
		`<Relationship Id="rId${intRelNum + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
		`<Relationship Id="rId${intRelNum + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`
	// The presentation-level commentAuthors part is shared by every slide's comments, so it is
	// related once from the presentation (only when the deck has at least one comment).
	if ((slides || []).some((slide) => (slide._comments || []).length > 0)) {
		strXml += `<Relationship Id="rId${intRelNum + 5}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>`
	}
	// Embedded fonts: one `font` rel per face, ids continuing past the fixed rels above.
	for (const face of flattenEmbeddedFaces(embeddedFonts || [], presentationFontRelStart(slides))) {
		strXml += `<Relationship Id="rId${face.rId}" Type="${FONT_REL_TYPE}" Target="fonts/font${face.partIndex}.fntdata"/>`
	}
	strXml += '</Relationships>'

	return strXml
}

// ===== Theme, presentation & root files =====

// XML-GEN: Last 5 functions create root /ppt files

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

/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {PresentationPropsInternal} pres - presentation
 * @return {string} XML
 */
export function makeXmlPresentation(pres: PresentationPropsInternal): string {
	let strXml =
		`${XML_DECL}${CRLF}` +
		'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		// When fonts are embedded we carry WHOLE faces, so `embedTrueTypeFonts="1"` (so
		// PowerPoint honors the embed) and `saveSubsetFonts="0"` (we did not subset).
		// With no embedded fonts, keep the historical inert `saveSubsetFonts="1"`.
		`xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${pres.rtlMode ? 'rtl="1"' : ''} ${(pres.embeddedFonts || []).some((font) => font.faces.some((face) => face.bytes)) ? 'embedTrueTypeFonts="1" saveSubsetFonts="0"' : 'saveSubsetFonts="1"'} autoCompressPictures="0"${pres.firstSlideNum !== 1 ? ` firstSlideNum="${pres.firstSlideNum}"` : ''}>`

	// STEP 1: Add slide master (SPEC: tag 1 under <presentation>)
	strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'

	// STEP 2: Add Notes Master (SPEC: tag 2 under <presentation>)
	// CT_Presentation child sequence (ECMA-376 Part 1 §19.2.1.26) requires
	// notesMasterIdLst to appear BEFORE sldIdLst. Emitting it after sldIdLst
	// (or after sldSz/notesSz) violates the schema and is flagged by
	// OpenXmlValidator as Sch_UnexpectedElementContentExpectingComplex.
	// (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
	strXml += `<p:notesMasterIdLst><p:notesMasterId r:id="rId${pres.slides.length + 2}"/></p:notesMasterIdLst>`

	// STEP 3: Add all Slides (SPEC: tag 3 under <presentation>)
	strXml += '<p:sldIdLst>'
	pres.slides.forEach((slide) => (strXml += `<p:sldId id="${slide._slideId}" r:id="rId${slide._rId}"/>`))
	strXml += '</p:sldIdLst>'

	// STEP 4: Add sizes
	strXml += `<p:sldSz cx="${pres.presLayout.width}" cy="${pres.presLayout.height}"/>`
	strXml += `<p:notesSz cx="${pres.presLayout.height}" cy="${pres.presLayout.width}"/>`

	// STEP 4b: Embedded fonts (CT_Presentation index 7 — after notesSz, before defaultTextStyle).
	// rIds continue past the fixed presentation rels and must match makeXmlPresentationRels.
	{
		const fonts = pres.embeddedFonts || []
		const flat = flattenEmbeddedFaces(fonts, presentationFontRelStart(pres.slides))
		const rIdOf = new Map(flat.map((face) => [`${face.fontIndex}:${face.slot}`, face.rId]))
		strXml += serializeEmbeddedFontLst(fonts, (fontIndex, slot) => rIdOf.get(`${fontIndex}:${slot}`))
	}

	// STEP 5: Add text styles
	strXml += '<p:defaultTextStyle>'
	for (let idy = 1; idy < 10; idy++) {
		strXml +=
			`<a:lvl${idy}pPr marL="${(idy - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
			'<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' +
			`</a:defRPr></a:lvl${idy}pPr>`
	}
	strXml += '</p:defaultTextStyle>'

	// STEP 6: Add Sections (if any)
	if (pres.sections && pres.sections.length > 0) {
		strXml += '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
		strXml += '<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
		pres.sections.forEach((sect) => {
			strXml += `<p14:section name="${encodeXmlEntities(sect.title)}" id="{${getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')}}"><p14:sldIdLst>`
			sect._slides.forEach((slide) => (strXml += `<p14:sldId id="${slide._slideId}"/>`))
			strXml += '</p14:sldIdLst></p14:section>'
		})
		strXml += '</p14:sectionLst></p:ext>'
		strXml +=
			'<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext>'
		strXml += '</p:extLst>'
	}

	// Done
	strXml += '</p:presentation>'
	return strXml
}

/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export function makeXmlPresProps(): string {
	return `${XML_DECL}${CRLF}<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`
}

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles(tableStyles: TableStyleInternal[] = []): string {
	const open = `${XML_DECL}${CRLF}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"`
	if (!tableStyles || tableStyles.length === 0) return `${open}/>`

	let strXml = `${open}>`
	tableStyles.forEach(({ guid, def }) => {
		strXml += `<a:tblStyle styleId="${guid}" styleName="${encodeXmlEntities(def.name)}">`
		// NOTE: regions MUST be emitted in CT_TableStyle schema order or PowerPoint reports the file as corrupt
		;(
			[
				['wholeTbl', def.wholeTbl],
				['band1H', def.band1H],
				['band2H', def.band2H],
				['band1V', def.band1V],
				['band2V', def.band2V],
				['lastCol', def.lastCol],
				['firstCol', def.firstCol],
				['lastRow', def.lastRow],
				['firstRow', def.firstRow],
			] as const
		).forEach(([name, region]) => {
			if (region) strXml += genXmlTableStyleRegion(name, region)
		})
		strXml += '</a:tblStyle>'
	})
	strXml += '</a:tblStyleLst>'
	return strXml
}

/**
 * Build one `CT_TablePartStyle` region (e.g. `firstRow`, `band1H`) for a custom table style.
 * Emits `tcTxStyle` (text) before `tcStyle` (cell fill/borders) per the schema sequence.
 * @param {string} name - region element name
 * @param {TableStyleRegionProps} region - region styling
 * @return {string} XML
 */
function genXmlTableStyleRegion(name: string, region: TableStyleRegionProps): string {
	let xml = `<a:${name}>`

	// A: tcTxStyle — text style (only when text formatting is requested)
	if (region.bold !== undefined || region.italic !== undefined || region.color) {
		const b = region.bold ? ' b="on"' : ''
		const i = region.italic ? ' i="on"' : ''
		xml += `<a:tcTxStyle${b}${i}><a:fontRef idx="minor"/>`
		xml += region.color ? createColorElement(region.color) : ''
		xml += '</a:tcTxStyle>'
	}

	// B: tcStyle — cell style: tcBdr (borders) then fill, in schema order
	if (region.border !== undefined || region.fill !== undefined) {
		xml += '<a:tcStyle>'
		if (region.border !== undefined) xml += genXmlTableStyleBorders(region.border)
		if (region.fill !== undefined) xml += `<a:fill>${genXmlColorSelection(region.fill)}</a:fill>`
		xml += '</a:tcStyle>'
	}

	xml += `</a:${name}>`
	return xml
}

/**
 * Build the `tcBdr` border block for a custom table style region.
 * A single `BorderProps` styles all four sides plus the interior grid lines; a
 * TRBL array styles only the four outer sides. Sides are emitted in schema order.
 * @param {BorderProps | BorderProps[]} border - border definition
 * @return {string} XML
 */
function genXmlTableStyleBorders(border: BorderProps | BorderProps[]): string {
	// NOTE: order MUST be left,right,top,bottom,insideH,insideV (CT_TableCellBorderStyle sequence)
	let sides: Array<[string, BorderProps | undefined]>
	if (Array.isArray(border)) {
		const [top, right, bottom, left] = border // TRBL input order
		sides = [
			['left', left],
			['right', right],
			['top', top],
			['bottom', bottom],
		]
	} else {
		sides = [
			['left', border],
			['right', border],
			['top', border],
			['bottom', border],
			['insideH', border],
			['insideV', border],
		]
	}

	let xml = '<a:tcBdr>'
	sides.forEach(([side, b]) => {
		if (!b) return
		xml += `<a:${side}>`
		if (b.type === 'none') {
			xml += '<a:ln><a:noFill/></a:ln>'
		} else {
			xml += `<a:ln w="${lineWidthToEmu(resolveBorderWidth(b, 1))}" cap="flat" cmpd="sng" algn="ctr">`
			xml += genXmlColorSelection({ color: b.color ?? '666666', transparency: b.transparency })
			xml += `<a:prstDash val="${b.type === 'dash' ? 'sysDash' : 'solid'}"/>`
			xml += '</a:ln>'
		}
		xml += `</a:${side}>`
	})
	xml += '</a:tcBdr>'
	return xml
}

/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export function makeXmlViewProps(): string {
	return `${XML_DECL}${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
}
