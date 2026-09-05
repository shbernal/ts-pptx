/**
 * ts-pptx: the connector construct family
 *
 * Lines drawn between two points, emitted as `<p:cxnSp>`.
 */

import { addConnectorDefinition } from '../gen/define/connector.js'
import type { ConstructFamily } from './shared.js'

export const connectorFamily: ConstructFamily = {
	name: 'connector',
	authors: {
		addConnector(slide, options) {
			addConnectorDefinition(slide, options)
		},
	},
}
