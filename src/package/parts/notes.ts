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
	makeXmlNotesMaster,
	makeXmlNotesMasterRel,
	makeXmlNotesSlide,
	makeXmlNotesSlideRel,
} from '../../gen/slide/notes.js'
import { NOTES_MASTER_PATH, notesSlidePath, overrideName, relsPath } from '../../gen/opc/part-paths.js'
import type { PartContributor } from './shared.js'

const CT_NOTES_MASTER = 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'
const CT_NOTES_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'

export const notesContributor: PartContributor = {
	order: 10,
	parts: {
		withEachSlide(slide: PresSlideInternal, slideNumber: number, zip: ZipWriter): void {
			// Notes of empty strings are created for slides which do not have notes specified, to
			// keep track of _rels.
			zip.add(notesSlidePath(slideNumber), makeXmlNotesSlide(slide))
			zip.add(relsPath(notesSlidePath(slideNumber)), makeXmlNotesSlideRel(slide, slideNumber))
		},
		afterMaster(_pres: PresentationPropsInternal, zip: ZipWriter): void {
			zip.add(NOTES_MASTER_PATH, makeXmlNotesMaster())
			zip.add(relsPath(NOTES_MASTER_PATH), makeXmlNotesMasterRel())
		},
	},
	contentTypes: {
		presentation(): ContentTypeOverride[] {
			return [{ partName: overrideName(NOTES_MASTER_PATH), contentType: CT_NOTES_MASTER }]
		},
		trailing(pres: PresentationPropsInternal): ContentTypeOverride[] {
			return pres.slides.map((_slide, idx) => ({
				partName: overrideName(notesSlidePath(idx + 1)),
				contentType: CT_NOTES_SLIDE,
			}))
		},
	},
}
