/**
 * ts-pptx: the comments construct family
 *
 * Legacy PowerPoint review comments: the marker on the slide, the per-slide comment part, and the
 * deck-wide author list.
 */

import { commentsContributor } from '../package/parts/comments.js'
import { addCommentDefinition } from '../gen/define/comment.js'
import type { ConstructFamily } from './shared.js'

export const commentsFamily = {
	name: 'comments',
	authors: {
		addComment(slide, options) {
			addCommentDefinition(slide, options)
		},
	},
	parts: commentsContributor,
} satisfies ConstructFamily
