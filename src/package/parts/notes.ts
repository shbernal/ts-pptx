/**
 * ts-pptx: the speaker-notes contribution to a written package
 *
 * A notes slide per slide (empty ones included, so every slide's `.rels` numbering is the same
 * whether or not it has notes) plus the single notes master the whole deck shares.
 *
 * `ppt/theme/theme2.xml` is the one piece of this that is *not* here: `notesMaster1.xml.rels`
 * references it, but it is written on the core path because the master's own theme part is, and
 * splitting one of a pair across the seam buys nothing.
 */

import type { ZipWriter } from '../../zip.js'
import type { ContentTypeOverride } from '../../gen/opc/content-types.js'
import type { PresentationPropsInternal, PresSlideInternal } from '../../types/internal.js'
import {
	buildNotesSlideRels,
	makeXmlNotesMaster,
	makeXmlNotesMasterRel,
	makeXmlNotesSlide,
	makeXmlNotesSlideRel,
} from '../../gen/slide/notes.js'
import { NOTES_MASTER_PATH, notesSlidePath, overrideName, relsPath } from '../../gen/opc/part-paths.js'
import type { PartContributor } from './shared.js'
import { NOTES_MASTER_CONTENT_TYPE, NOTES_SLIDE_CONTENT_TYPE } from '../../ooxml/rel-types.js'

export const notesContributor: PartContributor = {
	order: 10,
	parts: {
		withEachSlide(slide: PresSlideInternal, slideNumber: number, zip: ZipWriter): void {
			// Notes of empty strings are created for slides which do not have notes specified, to
			// keep track of _rels. The body and its rels are written from one resolution of the notes,
			// so their hyperlink ids cannot disagree.
			const notes = buildNotesSlideRels(slide)
			zip.add(notesSlidePath(slideNumber), makeXmlNotesSlide(slide, notes))
			zip.add(relsPath(notesSlidePath(slideNumber)), makeXmlNotesSlideRel(notes, slideNumber))
		},
		afterMaster(_pres: PresentationPropsInternal, zip: ZipWriter): void {
			zip.add(NOTES_MASTER_PATH, makeXmlNotesMaster())
			zip.add(relsPath(NOTES_MASTER_PATH), makeXmlNotesMasterRel())
		},
	},
	contentTypes: {
		presentation(): ContentTypeOverride[] {
			return [{ partName: overrideName(NOTES_MASTER_PATH), contentType: NOTES_MASTER_CONTENT_TYPE }]
		},
		trailing(pres: PresentationPropsInternal): ContentTypeOverride[] {
			return pres.slides.map((_slide, idx) => ({
				partName: overrideName(notesSlidePath(idx + 1)),
				contentType: NOTES_SLIDE_CONTENT_TYPE,
			}))
		},
	},
}
