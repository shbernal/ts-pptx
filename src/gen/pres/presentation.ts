/**
 * PptxGenJS: `ppt/presentation.xml` + presProps/viewProps
 *
 * Emit the presentation part (masters/notes/slide id lists, sizes, embedded-font
 * list, default text styles, sections) and the small presProps/viewProps parts.
 */

import { CRLF, XML_DECL } from '../../core-enums.js'
import type { PresentationPropsInternal } from '../../core-interfaces.js'
import { encodeXmlEntities, getUuid } from '../../gen-utils.js'
import { flattenEmbeddedFaces, serializeEmbeddedFontLst } from '../../embedded-fonts.js'
import { presentationFontRelStart } from './presentation-rels.js'

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
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export function makeXmlViewProps(): string {
	return `${XML_DECL}${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
}
