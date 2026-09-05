/**
 * ts-pptx: the comments construct family
 *
 * Legacy PowerPoint review comments: the marker on the slide, the per-slide comment part, and the
 * deck-wide author list.
 */

import { addCommentDefinition } from '../gen/define/comment.js'
import type { ConstructFamily } from './shared.js'

export const commentsFamily: ConstructFamily = {
	name: 'comments',
	authors: {
		addComment(slide, options) {
			addCommentDefinition(slide, options)
		},
	},
}
