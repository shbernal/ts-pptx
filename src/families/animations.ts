/**
 * ts-pptx: the animations construct family
 *
 * Preset build animations (`addAnimation`). Authoring one is only a push onto the slide's list --
 * what the family is really worth is the timing tree that list turns into, and that is still
 * emitted from the slide writer rather than from here.
 */

import type { ConstructFamily } from './shared.js'

export const animationsFamily: ConstructFamily = {
	name: 'animations',
	authors: {
		addAnimation(slide, options) {
			slide._animations.push(options)
		},
	},
}
