/**
 * ts-pptx: the 3D-model construct family
 *
 * Embedded `.glb` models (PowerPoint's Insert > 3D Models), with the preview picture every other
 * consumer draws instead.
 */

import { addModel3dDefinition } from '../gen/define/model3d.js'
import type { ConstructFamily } from './shared.js'

export const model3dFamily: ConstructFamily = {
	name: 'model3d',
	authors: {
		addModel3d(slide, options) {
			addModel3dDefinition(slide, options)
		},
	},
}
