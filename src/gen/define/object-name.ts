/**
 * ts-pptx: Slide-Object Name Indexing
 *
 * The per-slide default-name counter shared by every `add*Definition`. Kept in one place so a
 * grouped child and a later top-level object never collide on the same Selection Pane name
 * (see the note on {@link nextObjectNameIdx}).
 */
import { SlideObjectType } from '../../core-enums.js'
import type { PresSlideInternal } from '../../types/internal.js'

/**
 * Take the next slide-wide index for `type`'s default Selection Pane name (`Shape 0`, `Image 1`,
 * `Group 1`, …).
 *
 * Default names used to be derived by counting the matching objects already in
 * `target._slideObjects`. `buildGroupObject` splices group children back out of that array, so the
 * count never advanced past them and a later top-level object reused a grouped child's name. This
 * counter is immune to the splice: every object of a kind consumes an index when it is added,
 * whether it stays top-level or moves into a group — at any nesting depth, since `target` stays the
 * slide all the way down.
 *
 * Shapes and text boxes deliberately share the `text` bucket (both are `_type === text`), which is
 * what keeps `Shape 0` and `Text 0` from colliding on one slide. Callers take an index
 * unconditionally, including when the caller supplied an explicit `objectName`, so an object's
 * index is its ordinal among its kind rather than a count of the defaulted ones.
 * @param target - slide (or master) the object is being added to
 * @param type - the object's `_type`
 * @returns the index this object takes
 */
export function nextObjectNameIdx(target: PresSlideInternal, type: SlideObjectType): number {
	const counts = (target._objectNameCounts ??= {})
	const idx = counts[type] ?? 0
	counts[type] = idx + 1
	return idx
}
