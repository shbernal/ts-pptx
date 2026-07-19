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
import { el, raw, voidEl } from '../oxml/el.js'

const SCHEMA_BASE = 'http://schemas.openxmlformats.org/'
const OFFICE_REL = SCHEMA_BASE + 'officeDocument/2006/relationships/'

function relationship(rId: number, type: string, target: string): string {
	return voidEl('Relationship', { Id: `rId${rId}`, Type: type, Target: target })
}

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
	const rels: string[] = [relationship(1, OFFICE_REL + 'slideMaster', 'slideMasters/slideMaster1.xml')]
	for (let idx = 1; idx <= slides.length; idx++) {
		rels.push(relationship(++intRelNum, OFFICE_REL + 'slide', `slides/slide${idx}.xml`))
	}
	intRelNum++
	rels.push(
		relationship(intRelNum + 0, OFFICE_REL + 'notesMaster', 'notesMasters/notesMaster1.xml'),
		relationship(intRelNum + 1, OFFICE_REL + 'presProps', 'presProps.xml'),
		relationship(intRelNum + 2, OFFICE_REL + 'viewProps', 'viewProps.xml'),
		relationship(intRelNum + 3, OFFICE_REL + 'theme', 'theme/theme1.xml'),
		relationship(intRelNum + 4, OFFICE_REL + 'tableStyles', 'tableStyles.xml')
	)
	// The presentation-level commentAuthors part is shared by every slide's comments, so it is
	// related once from the presentation (only when the deck has at least one comment).
	if ((slides || []).some((slide) => (slide._comments || []).length > 0)) {
		rels.push(relationship(intRelNum + 5, OFFICE_REL + 'commentAuthors', 'commentAuthors.xml'))
	}
	// Embedded fonts: one `font` rel per face, ids continuing past the fixed rels above.
	for (const face of flattenEmbeddedFaces(embeddedFonts || [], presentationFontRelStart(slides))) {
		rels.push(relationship(face.rId, FONT_REL_TYPE, `fonts/font${face.partIndex}.fntdata`))
	}

	return XML_DECL + CRLF + el('Relationships', { xmlns: SCHEMA_BASE + 'package/2006/relationships' }, rels.map(raw))
}
