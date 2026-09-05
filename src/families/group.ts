/**
 * ts-pptx: the group construct family
 *
 * Both ways to make a `<p:grpSp>`: from child descriptors (`addGroup`) and from objects already on
 * the slide (`groupObjects`). The group *renderer* is not here -- it recurses back into the shape
 * walk and owns the frame and bounds logic that belongs to the dispatch, so it stays there.
 */

import { addGroupDefinition, groupObjectsDefinition } from '../gen/define/group.js'
import type { ConstructFamily } from './shared.js'

export const groupFamily = {
	name: 'group',
	authors: {
		addGroup(slide, children, options) {
			// The child descriptors resolve through the same table `createSlideMaster` uses, so a group
			// can hold exactly the kinds this presentation was composed to author -- no more.
			addGroupDefinition(slide, children, options || {}, slide._childAuthors)
		},
		groupObjects(slide, objectNames, options) {
			groupObjectsDefinition(slide, objectNames, options || {})
		},
	},
} satisfies ConstructFamily
