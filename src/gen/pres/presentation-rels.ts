/**
 * ts-pptx: `ppt/_rels/presentation.xml.rels`
 *
 * Emit the presentation relationships (slideMaster, slides, notesMaster/
 * presProps/viewProps/theme/tableStyles, optional commentAuthors, embedded-font
 * faces). `presentationFontRelStart` is shared with the presentation-part writer
 * so the embeddedFontLst face `r:id`s match the rels that back them.
 */

import type { PresSlideInternal } from '../../types/internal.js'
import { type EmbeddedFont, FONT_REL_TYPE, flattenEmbeddedFaces } from '../../embedded-fonts.js'
import { relationshipEl, relationshipsPart } from '../opc/rels.js'
import {
	NOTES_MASTER_REL,
	OFFICE_REL,
	SLIDE_MASTER_REL,
	SLIDE_REL,
	TABLE_STYLES_REL,
	THEME_REL,
} from '../../ooxml/rel-types.js'

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
	const rels: string[] = [relationshipEl(1, SLIDE_MASTER_REL, 'slideMasters/slideMaster1.xml')]
	for (let idx = 1; idx <= slides.length; idx++) {
		rels.push(relationshipEl(++intRelNum, SLIDE_REL, `slides/slide${idx}.xml`))
	}
	intRelNum++
	rels.push(
		relationshipEl(intRelNum + 0, NOTES_MASTER_REL, 'notesMasters/notesMaster1.xml'),
		relationshipEl(intRelNum + 1, OFFICE_REL + 'presProps', 'presProps.xml'),
		relationshipEl(intRelNum + 2, OFFICE_REL + 'viewProps', 'viewProps.xml'),
		relationshipEl(intRelNum + 3, THEME_REL, 'theme/theme1.xml'),
		relationshipEl(intRelNum + 4, TABLE_STYLES_REL, 'tableStyles.xml')
	)
	// The presentation-level commentAuthors part is shared by every slide's comments, so it is
	// related once from the presentation (only when the deck has at least one comment).
	if ((slides || []).some((slide) => (slide._comments || []).length > 0)) {
		rels.push(relationshipEl(intRelNum + 5, OFFICE_REL + 'commentAuthors', 'commentAuthors.xml'))
	}
	// Embedded fonts: one `font` rel per face, ids continuing past the fixed rels above.
	for (const face of flattenEmbeddedFaces(embeddedFonts || [], presentationFontRelStart(slides))) {
		rels.push(relationshipEl(face.rId, FONT_REL_TYPE, `fonts/font${face.partIndex}.fntdata`))
	}

	return relationshipsPart(rels)
}
