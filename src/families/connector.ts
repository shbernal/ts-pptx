/**
 * ts-pptx: the connector construct family
 *
 * Lines drawn between two points, emitted as `<p:cxnSp>`.
 */

import { SlideObjectType } from '../enums.js'
import { renderConnectorObject } from '../gen/slide/objects/connector.js'
import { addConnectorDefinition } from '../gen/define/connector.js'
import type { ConstructFamily } from './shared.js'

export const connectorFamily: ConstructFamily = {
	name: 'connector',
	authors: {
		addConnector(slide, options) {
			addConnectorDefinition(slide, options)
		},
	},
	renderers: {
		[SlideObjectType.connector]: renderConnectorObject,
	},
}
