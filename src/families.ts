/**
 * ts-pptx: the construct families, as the `pptx-ts/families` subpath publishes them
 *
 * What a caller hands to `createPresentation({ use: [...] })`. Each value is one family: the
 * methods it adds to a slide, the child descriptors it recognises in a slide master or a group,
 * the renderer that emits its XML, and the parts it puts in the package.
 *
 * Naming one here is what puts its code in your bundle, which is the whole point: a program that
 * composes `charts` pays for charts, and a program that does not, does not.
 *
 * The core tier -- text, shapes, images, groups and speaker notes -- is composed for you. Those
 * five are exported anyway, so a caller can be explicit; listing one changes nothing.
 */

import { animationsFamily } from './families/animations.js'
import { chartFamily } from './families/chart.js'
import { commentsFamily } from './families/comments.js'
import { connectorFamily } from './families/connector.js'
import { groupFamily } from './families/group.js'
import { imageFamily } from './families/image.js'
import { measureFamily } from './families/measure.js'
import { mediaFamily } from './families/media.js'
import { model3dFamily } from './families/model3d.js'
import { notesFamily } from './families/notes.js'
import { oleFamily } from './families/ole.js'
import { shapeFamily } from './families/shape.js'
import { tableDomFamily } from './families/table-dom.js'
import { tableFamily } from './families/table.js'
import { textFamily } from './families/text.js'
import { zoomFamily } from './families/zoom.js'

export type { ConstructFamily } from './families/shared.js'

/** Text boxes. Core: composed whether you ask for it or not. */
export const text = textFamily
/** Preset shapes (`addShape`), and the `rect`/`line`/`roundRect` descriptors. Core. */
export const shapes = shapeFamily
/** Raster and SVG pictures. Core. */
export const images = imageFamily
/** Groups, from descriptors (`addGroup`) or from objects already on the slide (`groupObjects`). Core. */
export const groups = groupFamily
/** Speaker notes, and the notes slides they are written to. Core. */
export const notes = notesFamily

/** Charts of every type, their embedded workbooks, and the chartEx sidecars. The expensive one. */
export const charts = chartFamily
/** Tables, including auto-paging onto continuation slides. */
export const tables = tableFamily
/**
 * `tableToSlides`: reproducing a rendered HTML `<table>`. Needs a live DOM, so it belongs to a
 * program that has one; the same conversion is a free function on `pptx-ts/html`.
 */
export const domTables = tableDomFamily
/** Embedded and online audio/video. */
export const media = mediaFamily
/** Connectors: lines drawn between two points. */
export const connectors = connectorFamily
/** Embedded OLE objects, whose payload travels inside the `.pptx`. */
export const oleObjects = oleFamily
/** Embedded 3D models (`.glb`) with their preview picture. */
export const models3d = model3dFamily
/** Slide, Section and Summary Zoom tiles. */
export const zooms = zoomFamily
/** Review comments and the deck-wide author list. */
export const comments = commentsFamily
/** Preset build animations. */
export const animations = animationsFamily
/** `measureText`, `overflowsBox` and `tableLayout` on the presentation. */
export const measure = measureFamily
