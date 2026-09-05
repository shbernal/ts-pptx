/**
 * ts-pptx: the construct families the three main entries are composed with
 *
 * The full authoring surface, as a list. `index.ts`, `node.ts` and `browser.ts` each compose every
 * family, so `TsPptx` authors everything it ever has; the browser entry adds the table family's
 * live-DOM half on top, because that one needs a `document` to read.
 *
 * It lives beside `entry-surface.ts`, and for the same reason: the three entries publish one API
 * and differ only in the runtime adapter they hand down, so what they are composed with is one
 * list rather than three that can drift.
 *
 * Naming a family here is what puts its code in the graph -- a named import inside a reachable
 * function body is retained unconditionally. That is the whole point of the seam
 * (`families/shared.ts`): a program composed with fewer families links fewer of them.
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
import { tableFamily } from './families/table.js'
import { textFamily } from './families/text.js'
import { zoomFamily } from './families/zoom.js'
import type { ConstructFamily } from './families/shared.js'

/** Every construct family the library can author with. */
export const ALL_CONSTRUCT_FAMILIES: readonly ConstructFamily[] = Object.freeze([
	animationsFamily,
	chartFamily,
	commentsFamily,
	connectorFamily,
	groupFamily,
	imageFamily,
	measureFamily,
	mediaFamily,
	model3dFamily,
	notesFamily,
	oleFamily,
	shapeFamily,
	tableFamily,
	textFamily,
	zoomFamily,
])
