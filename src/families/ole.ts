/**
 * ts-pptx: the OLE construct family
 *
 * Embedded OLE objects (PowerPoint's Insert > Object > Create from File): the payload's bytes ship
 * inside the `.pptx`, so a double-click opens the source document in place.
 */

import { SlideObjectType } from '../enums.js'
import { renderOleObject } from '../gen/slide/objects/ole.js'
import { addOleObjectDefinition } from '../gen/define/ole.js'
import type { ConstructFamily } from './shared.js'

export const oleFamily: ConstructFamily = {
	name: 'ole',
	authors: {
		addOleObject(slide, options) {
			addOleObjectDefinition(slide, options)
		},
	},
	renderers: {
		[SlideObjectType.oleObject]: renderOleObject,
	},
}
