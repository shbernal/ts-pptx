/**
 * PptxGenJS: `ppt/_rels/presentation.xml.rels`
 *
 * Emit the presentation relationships (slideMaster, slides, notesMaster/
 * presProps/viewProps/theme/tableStyles, optional commentAuthors, embedded-font
 * faces). `presentationFontRelStart` is shared with the presentation-part writer
 * so the embeddedFontLst face `r:id`s match the rels that back them.
 */

import { CRLF, XML_DECL } from '../../core-enums.js'
import type { PresSlideInternal } from '../../core-interfaces.js'
import { type EmbeddedFont, FONT_REL_TYPE, flattenEmbeddedFaces } from '../../embedded-fonts.js'

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
export function presentationFontRelStart(slides: PresSlideInternal[]): number {
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
