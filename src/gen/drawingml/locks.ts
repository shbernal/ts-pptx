/**
 * PptxGenJS: DrawingML object-lock serialization
 *
 * Emit the `a:spLocks` / `a:picLocks` / `a:graphicFrameLocks` / group-shape lock
 * elements. Each locking element type supports a different set of flags; the
 * `*_LOCK_ATTRS` tables list the valid attribute names in ECMA-376 emit order.
 */

import type { ObjectLockProps } from '../../core-interfaces.js'
import { warn } from '../../log.js'

// Object lock attributes valid for each DrawingML locking element, in emit order (ECMA-376 §20.1.2.2.x / §20.1.2.2.34).
// Object keys in `ObjectLockProps` mirror these attribute names 1:1, so serialization is a filtered lookup.
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
			warn(`objectLock.${key} is not supported on <${tag}> (object "${objectName ?? ''}") and was ignored.`)
		}
	}
	const attrs = allowed.filter((name) => lockMap[name] === true).map((name) => `${name}="1"`)
	return attrs.length > 0 ? `<${tag} ${attrs.join(' ')}/>` : ''
}
