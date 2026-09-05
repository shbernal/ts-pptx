/**
 * ts-pptx: every shape family's renderer, in one table
 *
 * The whole mapping from a slide object's kind to the code that emits its XML. It lives apart
 * from the dispatch that reads it (`gen/slide/object.ts`) for one reason: a named import inside a
 * reachable function body is retained unconditionally, so a dispatch that names the ten renderers
 * itself links every family into every program that writes a slide — a text box would carry the
 * chart emitter. Naming them here instead makes "which families does this program support" a
 * question about which table is handed to the walk.
 *
 * Nothing else belongs here. Anything a renderer shares with another renderer goes in
 * `objects/shared.ts`, so this module stays a list of names.
 */

import { SlideObjectType } from '../../enums.js'
import { renderChartObject } from './objects/chart.js'
import { renderConnectorObject } from './objects/connector.js'
import { renderImageObject } from './objects/image.js'
import { renderMediaObject } from './objects/media.js'
import { renderModel3dObject } from './objects/model3d.js'
import { renderOleObject } from './objects/ole.js'
import { renderTableObject } from './objects/table.js'
import { renderTextObject } from './objects/text.js'
import { renderZoomObject } from './objects/zoom.js'
import type { RendererTable } from './objects/shared.js'

/**
 * Every shape family the library can emit — the table the full authoring surface writes with.
 *
 * `Required` rather than the `Partial` the type allows: this one is the complete set by
 * definition, so a family added to `RenderedObjectType` has to be given a renderer here
 * rather than going quietly missing from the deck that claims to support everything.
 *
 * A text box and a layout placeholder share `renderTextObject`. They are one emitter with two
 * entry kinds — the placeholder differences it does draw come off `ctx.placeholder`, not off the
 * object's `_type`.
 */
export const ALL_OBJECT_RENDERERS: Required<RendererTable> = Object.freeze({
	[SlideObjectType.chart]: renderChartObject,
	[SlideObjectType.connector]: renderConnectorObject,
	[SlideObjectType.image]: renderImageObject,
	[SlideObjectType.media]: renderMediaObject,
	[SlideObjectType.model3d]: renderModel3dObject,
	[SlideObjectType.oleObject]: renderOleObject,
	[SlideObjectType.placeholder]: renderTextObject,
	[SlideObjectType.table]: renderTableObject,
	[SlideObjectType.text]: renderTextObject,
	[SlideObjectType.zoom]: renderZoomObject,
})
