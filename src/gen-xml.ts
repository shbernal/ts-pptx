/**
 * PptxGenJS: XML Generation — re-export barrel
 *
 * The OOXML emitter has been split into a layered `gen/` tree mirroring
 * `src/read/` (opc / pres / slide / drawingml / anim). This module now only
 * re-exports the part builders so existing consumers — chiefly `pptxgen.ts`,
 * which does `import * as genXml from './gen-xml.js'` — keep resolving
 * `genXml.makeXml*` unchanged. Charts still live in `gen-charts.ts`.
 *
 *   - gen/opc/*       [Content_Types].xml, root rels, app/core/custom props
 *   - gen/pres/*      presentation(.rels), theme, table styles, pres/view props
 *   - gen/slide/*     slide/layout/master parts + rels, notes, comments, spTree
 *   - gen/drawingml/* reusable DrawingML fragment builders (text, geometry, …)
 *   - gen/anim/*      slide timing, transitions, animation sequence builders
 */

// OPC package parts (gen/opc/*)
export { makeXmlContTypes } from './gen/opc/content-types.js'
export { makeXmlRootRels } from './gen/opc/root-rels.js'
export { makeXmlApp } from './gen/opc/app.js'
export { makeXmlCore } from './gen/opc/core.js'
export { makeXmlCustomProperties } from './gen/opc/custom-props.js'

// Presentation-level parts (gen/pres/*)
export { makeXmlPresentationRels } from './gen/pres/presentation-rels.js'
export { makeXmlPresentation, makeXmlPresProps, makeXmlViewProps } from './gen/pres/presentation.js'
export { makeXmlTheme } from './gen/pres/theme.js'
export { makeXmlTableStyles } from './gen/pres/table-styles.js'

// Slide-level parts (gen/slide/*)
export { makeXmlSlide, makeXmlSlideLayoutRel, makeXmlSlideRel } from './gen/slide/slide.js'
export { makeXmlLayout } from './gen/slide/layout.js'
export { makeXmlMaster, makeXmlMasterRel } from './gen/slide/master.js'
export {
	buildNotesSlideRels,
	getNotesFromSlide,
	makeXmlNotesMaster,
	makeXmlNotesMasterRel,
	makeXmlNotesSlide,
	makeXmlNotesSlideRel,
} from './gen/slide/notes.js'
export { makeXmlCommentAuthors, makeXmlComments, resolveCommentAuthors } from './gen/slide/comments.js'
export type { ResolvedComments } from './gen/slide/comments.js'
