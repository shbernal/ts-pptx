/**
 * ts-pptx: Group Definition
 *
 * The group (`<p:grpSp>`) definition layer: `addGroupDefinition` builds a group from child
 * descriptors (recursing for nested groups), `groupObjectsDefinition` lifts already-authored
 * top-level objects into a group, and `addChildDefinition` is the shared child-descriptor
 * dispatch (also used by the slide-master definition).
 */
import { SlideObjectType } from '../../enums.js'
import { warn } from '../../diagnostics.js'
import type { GroupChildProps, GroupProps, SlideMasterObject } from '../../types/index.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { resolveObjectName } from './object-name.js'
import { CHILD_DESCRIPTOR_FAMILIES, type ChildAuthors, type ChildDescriptorKey } from '../../families/shared.js'
import { InvalidOptionError, UnsupportedFeatureError } from '../../errors.js'
import { pickDefined } from '../../options-internal.js'

/**
 * Dispatch a key-tagged child-object descriptor (`{ text }`, `{ image }`, `{ shape }`, …) to the
 * family that recognizes it. Shared by `createSlideMaster` (slide master `objects`) and
 * `addGroupDefinition` (group children) so the descriptor mapping lives in one place.
 *
 * The mapping is handed in rather than spelled here. It used to be an `if` chain naming four
 * `add*Definition` functions, which put the chart definer — and every plot module under it — in the
 * graph of any program that can define a slide master, which is all of them.
 *
 * `placeholder` is intentionally not handled here — it is master-specific and needs the object's
 * index for `_placeholderIdx`, so `createSlideMaster` handles that case itself.
 * @param target - slide (or master) the object is appended to
 * @param object - the child descriptor
 * @param children - the child-descriptor authors this presentation was composed with
 * @returns `true` if the descriptor was recognized and added, else `false`
 */
export function addChildDefinition(
	target: PresSlideInternal,
	object: SlideMasterObject | GroupChildProps,
	children: ChildAuthors
): boolean {
	// A descriptor carries exactly one key, so this reads it rather than testing the table's keys in
	// whatever order the family list produced them: which family was composed first is not allowed to
	// decide anything.
	for (const key of Object.keys(object) as ChildDescriptorKey[]) {
		const author = children[key] as ((t: PresSlideInternal, child: unknown) => void) | undefined
		if (!author) continue
		author(target, (object as Record<string, unknown>)[key])
		return true
	}
	return false
}

/**
 * Warn about a child descriptor {@link addChildDefinition} did not author, from `call`.
 *
 * Two conditions arrive here and they have different fixes, so they get different codes. A key
 * {@link CHILD_DESCRIPTOR_FAMILIES} claims is valid and typed and was simply left uncomposed —
 * the family that would have drawn it is not in this presentation's list, so the message names it
 * the way `familyMethodUnavailable` names the family behind a missing method. Any other key is a
 * typo, which is all this walk used to be able to say, because before the split a descriptor that
 * matched nothing could not type-check in the first place.
 * @param object - the descriptor nothing authored
 * @param call - the public call the descriptor was written for, e.g. `addGroup()`
 */
export function warnUnauthoredChild(object: SlideMasterObject | GroupChildProps, call: string): void {
	const keys = Object.keys(object)
	const uncomposed = (keys as ChildDescriptorKey[]).find((key) => CHILD_DESCRIPTOR_FAMILIES[key])
	if (uncomposed) {
		const family = CHILD_DESCRIPTOR_FAMILIES[uncomposed]
		warn(
			'family/child-not-composed',
			`${call}: the '${uncomposed}' child descriptor needs the "${family}" construct family, and this presentation ` +
				`was composed without it; skipping. Add the "${family}" family to the list the presentation is composed with.`
		)
	} else {
		warn(
			'group/unrecognized-child',
			`${call} received an unrecognized child descriptor (${keys.join(', ')}); skipping.`
		)
	}
}

/**
 * Build a group (`<p:grpSp>`) render-object from its child descriptors, without appending the
 * group itself to the slide. Nested `group` children recurse, so a group can contain other groups.
 *
 * An identity child coordinate space is kept at every depth (emitted in `gen/slide/object.ts` as
 * `chOff/chExt == off/ext`), so children — including descendants of nested groups — keep their
 * slide-absolute coordinates and grouping is visually a no-op while making the objects one
 * selectable PowerPoint group. Charts, media, tables, and placeholders are not supported as group
 * children yet; each is skipped with a warning. When `opts.x/y/w/h` are omitted the group's bounds
 * are auto-computed (in `gen/slide/object.ts`) as the bounding box of its children.
 *
 * `target` stays the slide at every depth so leaf descendants register their image/chart rels and
 * unique ids slide-level, even when nested inside child groups.
 * @param target - slide the group's leaf children register rels against
 * @param children - the child-object descriptors
 * @param opts - group position/size/name options
 * @param childAuthors - the child-descriptor authors this presentation was composed with
 */
function buildGroupObject(
	target: PresSlideInternal,
	children: GroupChildProps[],
	opts: GroupProps,
	childAuthors: ChildAuthors
): SlideObject {
	const groupObjects: SlideObject[] = []

	;(children || []).forEach((child) => {
		// Nested group: recurse and embed the child group object directly (no slide splice — its own
		// leaf descendants still register against `target` inside the recursive call).
		if ('group' in child) {
			groupObjects.push(buildGroupObject(target, child.group.children, child.group.options || {}, childAuthors))
			return
		}
		// Reject object types grouping does not support yet (rels/ID/transform work pending).
		if ('chart' in child || 'placeholder' in child || 'table' in child || 'media' in child) {
			warn('group/unsupported-child', `addGroup() does not support '${Object.keys(child)[0]}' children yet; skipping.`)
			return
		}
		// Reuse the existing add*Definition logic (which registers any image/chart rels on the slide,
		// correctly — grouped children still reference slide-level relationships), then move the
		// just-appended object(s) off the slide's top-level list into this group's child list.
		const before = target._slideObjects.length
		if (!addChildDefinition(target, child, childAuthors)) {
			warnUnauthoredChild(child, 'addGroup()')
			return
		}
		groupObjects.push(...target._slideObjects.splice(before))
	})

	// A group with no renderable children resolves to a zero-size <p:grpSp> (auto-bounds over an empty
	// bbox is 0×0). That is the degenerate result AGENTS.md says to warn on rather than emit silently —
	// the same class as the partial-frame fallback. Warn once here, at the group that is actually empty:
	// a nested empty group has already warned in its own recursion, and its non-empty parent has not.
	if (groupObjects.length === 0) {
		warn(
			'group/no-children',
			`addGroup(): group "${opts.objectName ?? ''}" has no renderable children; emitting an empty, zero-size group. ` +
				'Pass at least one supported child (text/shape, image, or a nested group).'
		)
	}

	// Called after the children above so nested groups number inside-out (see `makeGroupObject`).
	return makeGroupObject(target, groupObjects, opts)
}

/**
 * Wrap already-built child render-objects in a group (`<p:grpSp>`) render-object, without appending
 * the group anywhere. Shared by `buildGroupObject` (children built from descriptors) and
 * `groupObjectsDefinition` (children lifted off the slide) so both entry points name and frame a
 * group identically — including the all-or-nothing frame, which is left to `gen/slide/object.ts` to resolve:
 * passing `x/y/w/h` through unset is what makes the bounds auto-compute to the children's bounding
 * box, so neither caller needs bounds math of its own.
 *
 * Call this *after* the children exist, so nested groups number inside-out: the name index is taken
 * here, and a child group that was built first has already taken a lower one.
 * @param target - slide the group belongs to (owns the name counter)
 * @param groupObjects - the group's child render-objects, in z-order
 * @param opts - group position/size/name options
 */
function makeGroupObject(target: PresSlideInternal, groupObjects: SlideObject[], opts: GroupProps): SlideObject {
	// Per slide rather than per process: a module-global counter made two identical presentations
	// built in one process disagree on their group names.
	const objectName = resolveObjectName(target, SlideObjectType.group, {
		label: 'Group',
		kind: 'group',
		supplied: opts.objectName,
	})

	return {
		_type: SlideObjectType.group,
		_groupObjects: groupObjects,
		// `pickDefined`, not a key-by-key literal: a group that states no `rotate` carries no
		// `rotate` key, where the literal wrote one holding `undefined` — and the group emitter
		// decides on key presence for the geometry it derives from its children.
		options: {
			...pickDefined(opts, ['x', 'y', 'w', 'h', 'rotate', 'flipH', 'flipV', 'altText', 'objectLock']),
			objectName,
		},
	}
}

/**
 * Add a group (`<p:grpSp>`) of child objects to a slide. Children may include nested groups.
 * @param target - slide the group is added to
 * @param children - the child-object descriptors
 * @param opts - group position/size/name options
 * @param childAuthors - the child-descriptor authors this presentation was composed with
 */
export function addGroupDefinition(
	target: PresSlideInternal,
	children: GroupChildProps[],
	opts: GroupProps,
	childAuthors: ChildAuthors
): void {
	target._slideObjects.push(buildGroupObject(target, children, opts, childAuthors))
}

/**
 * Object kinds `groupObjects()` accepts. Text (which shapes and text boxes both are), images,
 * connectors, and groups — the last so consumers can build nested logical groups out of groups they
 * already made. Charts, media, tables, and placeholders are excluded for the same reason
 * `addGroup()` excludes them (rels/id/transform work pending); a placeholder is excluded on top of
 * that because grouping it would sever the layout inheritance that makes it a placeholder.
 */
const GROUPABLE_TYPES: readonly SlideObjectType[] = [
	SlideObjectType.text,
	SlideObjectType.image,
	SlideObjectType.connector,
	SlideObjectType.group,
]

/**
 * Whether `groupObjects()` would accept this object, given only what it is. The remaining reasons a
 * call can still fail are about the *selection* rather than the object — an unresolved, duplicated
 * or ambiguous name — so a consumer cannot decide those from one object either way.
 *
 * Exported because `Slide.objects` reports it as `canGroup`, and a consumer that had to re-state the
 * rule would be pinning a copy of the list above: it would keep refusing the day grouping learns a
 * new kind, and go on offering the day one is withdrawn, with nothing failing in either direction.
 * @param {SlideObject} obj - a top-level authored object
 * @returns {boolean} true when its kind is groupable and it is not a placeholder
 */
export function isGroupableObject(obj: SlideObject): boolean {
	return GROUPABLE_TYPES.includes(obj._type) && !obj.options?.placeholder
}

/**
 * Depth-first search for an object named `name` among group children. Used only to explain a failed
 * lookup (see `groupObjectsDefinition`).
 */
function findNameInGroups(objects: SlideObject[], name: string): boolean {
	return objects.some((obj) => {
		if (obj._type !== SlideObjectType.group) return false
		return (obj._groupObjects || []).some(
			(child) => child.options?.objectName === name || findNameInGroups([child], name)
		)
	})
}

/**
 * Move already-authored, top-level slide objects into one group (`<p:grpSp>`), addressed by
 * `objectName`. The counterpart to `addGroup()` for consumers that compose a slide from independent
 * renderers: the objects already exist, so replaying their descriptors just to group them would mean
 * every renderer exposing its internals.
 *
 * Grouping is visually a no-op — children keep their slide-absolute geometry, their ids, their rels,
 * and their relative z-order, exactly as with `addGroup()`. Two ordering rules make that true:
 * children are ordered by their existing slide z-order (**not** by the order they are named, which
 * would silently restack the slide), and the group itself takes the topmost member's former slot, so
 * it lands where the selection's top edge already was.
 *
 * Every failure here throws rather than warns. An unmatched name is not a degenerate deck the author
 * might still want — it means the intended object is silently still loose on the slide, which is the
 * footgun the group was meant to remove.
 * @param target - slide holding the objects
 * @param objectNames - `objectName`s of the top-level objects to group
 * @param opts - group position/size/name options (same semantics as `addGroup()`)
 */
export function groupObjectsDefinition(target: PresSlideInternal, objectNames: string[], opts: GroupProps): void {
	if (!Array.isArray(objectNames) || objectNames.length === 0) {
		throw new InvalidOptionError(
			'group/missing-object-names',
			"groupObjects() requires a non-empty array of objectNames. Ex: `slide.groupObjects(['Title', 'Logo'])`"
		)
	}

	const requested = new Set<string>()
	objectNames.forEach((name) => {
		if (typeof name !== 'string' || name.trim().length === 0) {
			throw new InvalidOptionError(
				'group/invalid-object-name',
				`groupObjects(): every objectName must be a non-empty string (got ${JSON.stringify(name)}).`
			)
		}
		if (requested.has(name))
			throw new InvalidOptionError(
				'group/duplicate-object-name',
				`groupObjects(): objectName "${name}" was named more than once.`
			)
		requested.add(name)
	})

	// Resolve every name before moving anything, so a bad name leaves the slide untouched instead of
	// half-grouped.
	const members: SlideObject[] = []
	objectNames.forEach((name) => {
		// Stored names are the caller's own spelling (only `cNvPrOpen` escapes them), so the lookup
		// and every message below compare and quote the name as given.
		const matches = target._slideObjects.filter((obj) => obj.options?.objectName === name)
		const [obj, ambiguous] = matches
		if (!obj) {
			// Distinguish "no such object" from "already grouped": both leave the caller's name
			// unresolved, but only one of them is a typo.
			const hint = findNameInGroups(target._slideObjects, name)
				? 'it is already inside a group (an object can only belong to one group)'
				: 'no top-level object on this slide has that objectName'
			throw new InvalidOptionError('group/unresolved-object-name', `groupObjects(): cannot group "${name}" — ${hint}.`)
		}
		if (ambiguous) {
			throw new InvalidOptionError(
				'group/ambiguous-object-name',
				`groupObjects(): objectName "${name}" is ambiguous — ${matches.length} objects on this slide share it. Give them unique objectNames.`
			)
		}
		if (!isGroupableObject(obj)) {
			const kind = obj.options?.placeholder ? 'placeholder' : obj._type
			throw new UnsupportedFeatureError(
				'group/kind-not-groupable',
				`groupObjects(): cannot group "${name}" — grouping a ${String(kind)} is not supported yet.`
			)
		}
		members.push(obj)
	})

	// Order children by existing z-order, not by the order named: grouping must not restack the slide.
	const memberSet = new Set(members)
	const children = target._slideObjects.filter((obj) => memberSet.has(obj))

	// The wrapper takes the topmost member's former slot — i.e. it sits above everything the selection
	// sat above, and below everything it sat below. Counting the *survivors* ahead of that slot gives
	// the post-removal index directly, so the removal below cannot shift it.
	const topmostIdx = Math.max(...members.map((obj) => target._slideObjects.indexOf(obj)))
	const insertAt = target._slideObjects.slice(0, topmostIdx).filter((obj) => !memberSet.has(obj)).length

	const regrouped = target._slideObjects.filter((obj) => !memberSet.has(obj))
	regrouped.splice(insertAt, 0, makeGroupObject(target, children, opts))
	// Rewrite in place (rather than reassigning `_slideObjects`) so any held reference stays live.
	target._slideObjects.splice(0, target._slideObjects.length, ...regrouped)
}
