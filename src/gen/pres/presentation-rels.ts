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
import { fontPath, NOTES_MASTER_PATH, slidePath, SLIDE_MASTER_PATH, targetFromPresentation } from '../opc/part-paths.js'
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
 * The relationship ids `presentation.xml.rels` gives its fixed parts, and the first id free
 * after them.
 *
 * Layout: rId1 = slideMaster, rId2..(N+1) = N slides, then notesMaster/presProps/viewProps/
 * theme1/tableStyles, then commentAuthors when the deck has any comment, then the embedded
 * fonts.
 *
 * The rels writer and `makeXmlPresentation` both need these numbers — the presentation part
 * names the notesMaster and each embedded face by `r:id` — and each derived them separately,
 * one as `slides.length + 2` under a comment saying it copied the other's logic. Inserting a
 * relationship before `notesMaster` then broke the presentation's own reference with no
 * signal.
 * @param slides - the deck's slides, whose count and comments decide the layout
 */
export function presentationFixedRelIds(slides: PresSlideInternal[]): {
	notesMaster: number
	presProps: number
	viewProps: number
	theme: number
	tableStyles: number
	commentAuthors: number | null
	fontStart: number
} {
	const first = (slides?.length ?? 0) + 2
	const hasComments = (slides || []).some((slide) => (slide._comments || []).length > 0)
	return {
		notesMaster: first,
		presProps: first + 1,
		viewProps: first + 2,
		theme: first + 3,
		tableStyles: first + 4,
		commentAuthors: hasComments ? first + 5 : null,
		fontStart: first + (hasComments ? 6 : 5),
	}
}

/**
 * The first relationship id free for embedded-font rels — {@link presentationFixedRelIds}'
 * `fontStart`, kept as its own name because that is what the two font walks ask for.
 */
export function presentationFontRelStart(slides: PresSlideInternal[]): number {
	return presentationFixedRelIds(slides).fontStart
}

export function makeXmlPresentationRels(slides: PresSlideInternal[], embeddedFonts?: EmbeddedFont[]): string {
	const fixed = presentationFixedRelIds(slides)
	const rels: string[] = [relationshipEl(1, SLIDE_MASTER_REL, targetFromPresentation(SLIDE_MASTER_PATH))]
	for (let idx = 1; idx <= slides.length; idx++) {
		rels.push(relationshipEl(idx + 1, SLIDE_REL, targetFromPresentation(slidePath(idx))))
	}
	rels.push(
		relationshipEl(fixed.notesMaster, NOTES_MASTER_REL, targetFromPresentation(NOTES_MASTER_PATH)),
		relationshipEl(fixed.presProps, OFFICE_REL + 'presProps', 'presProps.xml'),
		relationshipEl(fixed.viewProps, OFFICE_REL + 'viewProps', 'viewProps.xml'),
		relationshipEl(fixed.theme, THEME_REL, 'theme/theme1.xml'),
		relationshipEl(fixed.tableStyles, TABLE_STYLES_REL, 'tableStyles.xml')
	)
	// The presentation-level commentAuthors part is shared by every slide's comments, so it is
	// related once from the presentation (only when the deck has at least one comment).
	if (fixed.commentAuthors !== null) {
		rels.push(relationshipEl(fixed.commentAuthors, OFFICE_REL + 'commentAuthors', 'commentAuthors.xml'))
	}
	// Embedded fonts: one `font` rel per face, ids continuing past the fixed rels above.
	for (const face of flattenEmbeddedFaces(embeddedFonts || [], presentationFontRelStart(slides))) {
		rels.push(relationshipEl(face.rId, FONT_REL_TYPE, targetFromPresentation(fontPath(face.partIndex))))
	}

	return relationshipsPart(rels)
}
