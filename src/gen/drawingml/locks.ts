/**
 * ts-pptx: DrawingML object-lock serialization
 *
 * Emit the `a:spLocks` / `a:picLocks` / `a:graphicFrameLocks` / group-shape lock
 * elements. Each locking element type supports a different set of flags; the
 * `*_LOCK_ATTRS` tables list the valid attribute names in ECMA-376 emit order.
 */

import type { ObjectLockProps } from '../../types/index.js'
import { voidEl } from '../oxml/el.js'
import { warn } from '../../diagnostics.js'

// Object lock attributes valid for each DrawingML locking element, in emit order (ECMA-376 §20.1.2.2.x / §20.1.2.2.34).
// Every name is a key of `ObjectLockProps`, so serialization is a filtered lookup; the assertion
// below the four tables is what keeps that true.
export const SHAPE_LOCK_ATTRS = [
	'noGrp',
	'noSelect',
	'noRot',
	'noChangeAspect',
	'noMove',
	'noResize',
	'noEditPoints',
	'noAdjustHandles',
	'noChangeArrowheads',
	'noChangeShapeType',
	'noTextEdit',
] as const
export const PICTURE_LOCK_ATTRS = [
	'noGrp',
	'noSelect',
	'noRot',
	'noChangeAspect',
	'noMove',
	'noResize',
	'noEditPoints',
	'noAdjustHandles',
	'noChangeArrowheads',
	'noChangeShapeType',
	'noCrop',
] as const
export const GRAPHIC_FRAME_LOCK_ATTRS = [
	'noGrp',
	'noDrilldown',
	'noSelect',
	'noChangeAspect',
	'noMove',
	'noResize',
] as const
export const GROUP_SHAPE_LOCK_ATTRS = ['noGrp', 'noSelect', 'noRot', 'noChangeAspect', 'noMove', 'noResize'] as const

/** Every attribute name the four tables between them list. */
type LockAttrName =
	| (typeof SHAPE_LOCK_ATTRS)[number]
	| (typeof PICTURE_LOCK_ATTRS)[number]
	| (typeof GRAPHIC_FRAME_LOCK_ATTRS)[number]
	| (typeof GROUP_SHAPE_LOCK_ATTRS)[number]

/**
 * Both halves of "the tables mirror `ObjectLockProps`", checked rather than asserted in a comment.
 *
 * No single table can be exhaustive — the four locking elements accept different attribute
 * subsets, and each table's order is byte-significant — so the claim is about their union. A key
 * added to the interface and to no table makes the second entry `never`, and a name in a table
 * that is not a key of the interface makes the first one `never`; either way `true` stops being
 * assignable and this fails to compile.
 *
 * Nothing imports it and nothing ever will: the `export` is what keeps it from being reported as
 * an unused binding, which is the one thing that would get an assertion like this deleted. Left
 * exported deliberately rather than by oversight — this note is the difference between the two.
 */
export const LOCK_ATTRS_MATCH_OBJECT_LOCK_PROPS: [
	LockAttrName extends keyof ObjectLockProps ? true : never,
	keyof Required<ObjectLockProps> extends LockAttrName ? true : never,
] = [true, true]

/**
 * Serialize an object-lock element (`a:spLocks` / `a:picLocks` / `a:graphicFrameLocks`).
 * Only flags set to `true` AND valid for this element type are emitted; a flag set on an
 * unsupported element type is dropped with a warning (silent coercion is a footgun).
 * @param tag - locking element tag, e.g. `'a:spLocks'`
 * @param allowed - attribute names this element type supports, in desired emit order
 * @param locks - merged lock flags (callers fold any hard-coded default in first)
 * @param objectName - for the warning message
 * @returns the locking element string, or `''` when no applicable flag is set
 */
export function genXmlObjectLock(
	tag: string,
	allowed: readonly string[],
	locks: ObjectLockProps | undefined,
	objectName?: string
): string {
	if (!locks) return ''
	const lockMap = locks as Record<string, boolean | undefined>
	for (const key of Object.keys(lockMap)) {
		if (lockMap[key] && !allowed.includes(key)) {
			warn(
				'object-lock/unsupported-on-shape',
				`objectLock.${key} is not supported on <${tag}> (object "${objectName ?? ''}") and was ignored.`
			)
		}
	}
	const set = allowed.filter((name) => lockMap[name] === true)
	return set.length > 0 ? voidEl(tag, Object.fromEntries(set.map((name) => [name, '1']))) : ''
}
