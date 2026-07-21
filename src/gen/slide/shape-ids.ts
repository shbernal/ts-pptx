/**
 * PptxGenJS: slide shape-id allocation
 *
 * Compute the `<p:cNvPr>` id every slide shape (including group children) is
 * rendered with, and resolve an `objectName` back to that id. References that
 * must name a shape before the render walk reaches it — a connector's
 * `<a:stCxn>`, an animation's `<p:spTgt spid>` — go through these helpers, so
 * they are shared by both the slide-object renderer and the animation builder.
 */

import { SlideObjectType } from '../../core-enums.js'
import type { SlideObject } from '../../types/internal.js'

/**
 * Every object a slide renders, paired with the `<p:cNvPr>` id it is rendered with: top-level
 * objects first (`index + 2`, in add order), then group children, seeded past the last top-level id
 * and allocated pre-order (a nested group takes an id before its own children do).
 *
 * This **mirrors** the allocation in `slideObjectToXml`, which hands out ids as it walks the tree.
 * A reference that must name an id *before* the walk reaches it — a connector's `<a:stCxn>`, an
 * animation's `<p:spTgt spid>` — cannot wait for that, so it resolves through this map instead.
 * The two must stay in step: `test/regression/group-shapes.test.js` parses the emitted `cNvPr` ids
 * back out and asserts each reference points at the shape it names, so drift fails there.
 * @param slideObjects - the slide's top-level objects
 * @returns each object's `<p:cNvPr>` id, keyed by object identity, in id order
 */
export function collectSlideShapeIds(slideObjects: SlideObject[]): Map<SlideObject, number> {
	const shapeIds = new Map<SlideObject, number>()
	slideObjects.forEach((obj, idx) => shapeIds.set(obj, idx + 2))

	let childIdxAlloc = slideObjects.length
	const allocGroupChildren = (children: SlideObject[]): void => {
		children.forEach((child) => {
			shapeIds.set(child, childIdxAlloc++ + 2)
			if (child._type === SlideObjectType.group) allocGroupChildren(child._groupObjects || [])
		})
	}
	slideObjects.forEach((obj) => {
		if (obj._type === SlideObjectType.group) allocGroupChildren(obj._groupObjects || [])
	})

	return shapeIds
}

/**
 * The `<p:cNvPr>` id of the object named `objectName`, or `null` when the slide has no such object.
 *
 * Group children are searched too: `buildGroupObject` splices them out of `_slideObjects` and into
 * their group's `_groupObjects`, but they are still `<p:cNvPr>`-named on this same slide and so are
 * legitimate targets for a connector binding or an animation. Searching only `_slideObjects` is what
 * silently dropped both.
 *
 * A top-level object wins over a group child of the same name (`slideObjectToXml` warns separately
 * about the duplicate), which leaves resolution unchanged for every deck without groups.
 * @param shapeIds - the slide's shape ids, from `collectSlideShapeIds` (iterated in id order)
 * @param objectName - the `objectName` to resolve
 * @returns the object's `<p:cNvPr>` id, or `null` when unresolved
 */
export function resolveObjectNameToId(shapeIds: Map<SlideObject, number>, objectName: string): number | null {
	for (const [obj, id] of shapeIds) {
		if (obj.options?.objectName === objectName) return id
	}
	return null
}
