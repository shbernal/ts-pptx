/**
 * ts-pptx: the image construct family
 *
 * Raster and SVG pictures, and the `image` descriptor a slide master or a group is written with.
 */

import { addImageDefinition } from '../gen/define/image.js'
import type { ConstructFamily } from './shared.js'

export const imageFamily: ConstructFamily = {
	name: 'image',
	authors: {
		addImage(slide, options) {
			addImageDefinition(slide, options)
		},
	},
	children: {
		image(target, child) {
			addImageDefinition(target, child)
		},
	},
}
