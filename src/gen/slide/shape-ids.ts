/**
 * ts-pptx: slide shape-id allocation
 *
 * Compute the `<p:cNvPr>` id every slide shape (including group children) is
 * rendered with, and resolve an `objectName` back to that id. References that
 * must name a shape before the render walk reaches it — a connector's
 * `<a:stCxn>`, an animation's `<p:spTgt spid>` — go through these helpers, so
 * they are shared by both the slide-object renderer and the animation builder.
 */

import { SlideObjectType } from '../../enums.js'
import type { SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue } from '../utils.js'

/**
 * The four `SlideObjectType` members that live in `_slideObjects` without drawing anything on the
 * slide: speaker notes belong to the notes part, a table cell is rendered by its owning table, a
 * hyperlink is a relationship carried by whatever shape owns it, and `online` is unreachable.
 * `slideObjectToXml` spells the same four out as the arms that emit nothing.
 */
const NON_RENDERING_TYPES: ReadonlySet<SlideObjectType> = new Set([
	SlideObjectType.notes,
	SlideObjectType.tablecell,
	SlideObjectType.hyperlink,
	SlideObjectType.online,
])

/**
 * The slide's top-level objects that actually reach the shape tree, in add order.
 *
 * This is the sequence a `shapeIndex` addresses and the sequence ids are handed out along. Both
 * used to run over `_slideObjects` itself, which counts the four non-rendering members: a slide
 * whose first entry was `addNotes` gave every later shape an id one higher than it emitted, so
 * `shapeIndex: 0` produced a `<p:spTgt spid>` naming nothing on the slide.
 * @param slideObjects - the slide's top-level objects
 * @returns the subset that renders a shape
 */
export function renderedSlideObjects(slideObjects: SlideObject[]): SlideObject[] {
	return slideObjects.filter((obj) => !NON_RENDERING_TYPES.has(obj._type))
}

/**
 * Every object a slide renders, paired with the `<p:cNvPr>` id it is rendered with: top-level
 * objects first (`index + 2`, in add order), then group children, seeded past the last top-level id
 * and allocated pre-order (a nested group takes an id before its own children do).
 *
 * Only the objects that render take an id, so the emitted ids run without gaps and every id in the
 * map names a shape that exists.
 *
 * This **mirrors** the allocation in `slideObjectToXml`, which hands out ids as it walks the tree.
 * A reference that must name an id *before* the walk reaches it — a connector's `<a:stCxn>`, an
 * animation's `<p:spTgt spid>` — cannot wait for that, so it resolves through this map instead.
 * The two must stay in step: `test/regression/shape/group-shapes.test.js` parses the emitted `cNvPr` ids
 * back out and asserts each reference points at the shape it names, so drift fails there.
 * @param slideObjects - the slide's top-level objects
 * @returns each object's `<p:cNvPr>` id, keyed by object identity, in id order
 */
export function collectSlideShapeIds(slideObjects: SlideObject[]): Map<SlideObject, number> {
	const shapeIds = new Map<SlideObject, number>()
	const rendered = renderedSlideObjects(slideObjects)
	rendered.forEach((obj, idx) => shapeIds.set(obj, idx + 2))

	let childIdxAlloc = rendered.length
	const allocGroupChildren = (children: SlideObject[]): void => {
		children.forEach((child) => {
			shapeIds.set(child, childIdxAlloc++ + 2)
			if (child._type === SlideObjectType.group) allocGroupChildren(child._groupObjects || [])
		})
	}
	rendered.forEach((obj) => {
		if (obj._type === SlideObjectType.group) allocGroupChildren(obj._groupObjects || [])
	})

	return shapeIds
}

/**
 * The `<p:cNvPr>` id of the object named `objectName`, or `null` when the slide has no such object.
 *
 * `objectName` is the **raw**, caller-supplied name — escaping is this helper's job. Every
 * `add*Definition` stores the name attribute-escaped (`encodeXmlAttrValue`) so `cNvPrOpen` can emit
 * it verbatim (see its comment in `gen/slide/objects/shared.ts`), which means the stored text is not the
 * caller's string: a shape added as `'Q&A'` is held as `Q&amp;A`. Comparing here rather than at each
 * call site keeps that one rule in one place — re-deriving it per caller is exactly how the
 * animation lookup came to disagree with the connector one and drop every effect naming a shape
 * whose name contained `&`, `<`, `>`, `"`, `'`, a tab or a newline.
 *
 * Group children are searched too: `buildGroupObject` splices them out of `_slideObjects` and into
 * their group's `_groupObjects`, but they are still `<p:cNvPr>`-named on this same slide and so are
 * legitimate targets for a connector binding or an animation. Searching only `_slideObjects` is what
 * silently dropped both.
 *
 * A top-level object wins over a group child of the same name (`slideObjectToXml` warns separately
 * about the duplicate), which leaves resolution unchanged for every deck without groups.
 * @param shapeIds - the slide's shape ids, from `collectSlideShapeIds` (iterated in id order)
 * @param objectName - the raw `objectName` to resolve, as the caller spelled it
 * @returns the object's `<p:cNvPr>` id, or `null` when unresolved
 */
export function resolveObjectNameToId(shapeIds: Map<SlideObject, number>, objectName: string): number | null {
	const key = encodeXmlAttrValue(objectName)
	for (const [obj, id] of shapeIds) {
		if (obj.options?.objectName === key) return id
	}
	return null
}
