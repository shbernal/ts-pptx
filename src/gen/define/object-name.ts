/**
 * ts-pptx: Slide-Object Name Indexing
 *
 * The per-slide default-name counter shared by every `add*Definition`. Kept in one place so a
 * grouped child and a later top-level object never collide on the same Selection Pane name
 * (see the note on {@link nextObjectNameIdx}).
 */
import { SlideObjectType } from '../../enums.js'
import type { PresSlideInternal } from '../../types/internal.js'
import { encodeXmlAttrValue, validateObjectName } from '../utils.js'

/**
 * Take the next slide-wide index for `type`'s default Selection Pane name — 0-based here, and
 * offset to the 1-based name by {@link resolveObjectName}.
 *
 * Default names used to be derived by counting the matching objects already in
 * `target._slideObjects`. `buildGroupObject` splices group children back out of that array, so the
 * count never advanced past them and a later top-level object reused a grouped child's name. This
 * counter is immune to the splice: every object of a kind consumes an index when it is added,
 * whether it stays top-level or moves into a group — at any nesting depth, since `target` stays the
 * slide all the way down.
 *
 * Shapes and text boxes deliberately share the `text` bucket (both are `_type === text`), which is
 * what keeps `Shape 1` and `Text 1` from colliding on one slide. Callers take an index
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

/**
 * The Selection Pane name an object takes: the caller's, validated and attribute-encoded, or the
 * next default for its kind — `Shape 1`, `Image 1`, `Group 1`, `Slide Zoom 1`, …
 *
 * Eleven definers wrote this out, and the index base they appended used to disagree: six counted
 * from 0 (`Shape 0`, `Text 0`, `Image 0`, `Connector 0`, `Media 0`, `Table 0`), four from 1
 * (`Group 1`, `Object 1`, `3D Model 1`, and the zoom tiles), and `addChart` neither — it derived
 * `Chart 0` by counting the chart objects already on the slide, which is the numbering this
 * counter exists to replace ({@link nextObjectNameIdx}).
 *
 * They are all 1-based now, which is the base PowerPoint itself uses: it names an inserted
 * rectangle `Rectangle 1`, and nothing it authors is ever suffixed `0`. There is no base
 * parameter left to pass, so a new definer cannot pick the other convention by accident.
 *
 * The index is taken unconditionally — including when the caller supplied a name — so an object's
 * index is its ordinal among its kind. See {@link nextObjectNameIdx}.
 *
 * @param target - slide (or master) the object is being added to
 * @param type - the object's `_type`, which selects the counter bucket
 * @param spec - `label` opens the default name, `kind` names the API in a validation warning, and
 *   `supplied` is the caller's own name when they gave one. `fallback` replaces the `label N`
 *   default for a kind that has a better one to offer — a placeholder is named after the
 *   placeholder it fills — and is used as given, since it is the library's own string rather than
 *   the caller's and has nothing to validate.
 */
export function resolveObjectName(
	target: PresSlideInternal,
	type: SlideObjectType,
	spec: { label: string; kind: string; supplied: string | undefined; fallback?: string }
): string {
	const idx = nextObjectNameIdx(target, type)
	if (spec.supplied) return encodeXmlAttrValue(validateObjectName(spec.supplied, spec.kind))
	return spec.fallback ?? `${spec.label} ${idx + 1}`
}
