/**
 * ts-pptx: the media construct family
 *
 * Embedded and online audio/video. The default video poster already loads with the media rather
 * than with the library; this is the rest of it.
 */

import { addMediaDefinition } from '../gen/define/media.js'
import type { ConstructFamily } from './shared.js'

export const mediaFamily: ConstructFamily = {
	name: 'media',
	authors: {
		addMedia(slide, options) {
			addMediaDefinition(slide, options)
		},
	},
}
