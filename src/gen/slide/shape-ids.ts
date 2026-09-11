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
 * How many `<p:cNvPr>` ids an object's render uses, starting at the one the allocator gives it.
 *
 * One for everything but a Summary Zoom. Its `mc:Fallback` is a group holding one picture per tile,
 * numbered from the zoom's own id plus one (`gen/slide/objects/zoom.ts`). The pictures are the
 * fallback's rendering of the zoom rather than shapes of their own, but a consumer reading the
 * fallback sees them beside every other shape on the slide, so their ids are held back from the
 * shapes that follow.
 * @param obj - a slide object that renders
 * @returns the number of ids it uses
 */
export function shapeIdCount(obj: SlideObject): number {
	return obj._type === SlideObjectType.zoom && obj.zoom?.variant === 'summary' ? 1 + obj.zoom.tiles.length : 1
}

/**
 * Every object a slide renders, paired with the `<p:cNvPr>` id it is rendered with: top-level
 * objects first (from 2, in add order), then group children, seeded past the last top-level id
 * and allocated pre-order (a nested group takes an id before its own children do). Each object
 * holds {@link shapeIdCount} ids, so the next object starts past all of them.
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
	let nextId = 2
	const allocate = (obj: SlideObject): void => {
		shapeIds.set(obj, nextId)
		nextId += shapeIdCount(obj)
	}
	rendered.forEach(allocate)

	const allocGroupChildren = (children: SlideObject[]): void => {
		children.forEach((child) => {
			allocate(child)
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
 * `objectName` is the **raw**, caller-supplied name, and it is compared as it is: every
 * `add*Definition` stores the name as the caller wrote it, and only `cNvPrOpen` escapes it. Connector
 * bindings and animation targets both resolve through here, so they cannot disagree about a name
 * containing `&`, `<`, `>`, `"`, `'`, a tab or a newline.
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
export function resolveObjectNameToId(shapeIds: ReadonlyMap<SlideObject, number>, objectName: string): number | null {
	for (const [obj, id] of shapeIds) {
		if (obj.options?.objectName === objectName) return id
	}
	return null
}
