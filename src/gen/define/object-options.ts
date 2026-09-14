/**
 * ts-pptx: what an OLE object, a 3D model, a zoom and a media object record about themselves
 *
 * Each of these definers built the same options literal for its slide record, the frame, the
 * Selection Pane name, the alt text and the lock flags, with only the name's label and the frame's
 * defaults differing; media spelled it as a run of `if`s. The three that embed a payload also
 * refused a missing source with the same throw, word for word but for the method's name.
 */

import type { InvalidOptionErrorCode } from '../../codes.js'
import type { SlideObjectType } from '../../enums.js'
import { InvalidOptionError } from '../../errors.js'
import type { Coord } from '../../types/index.js'
import type { ObjectOptionsInternal, PresSlideInternal } from '../../types/internal.js'
import type { ObjectLockProps } from '../../types/object.js'
import { type AuthoredFrame, resolveAuthoredFrame } from './frame.js'
import { resolveObjectName } from './object-name.js'

/** The options every one of these objects takes: a frame, a name, alt text and locks. */
export type FramedObjectProps = AuthoredFrame & {
	objectName?: string | undefined
	altText?: string | undefined
	objectLock?: ObjectLockProps | undefined
}

/**
 * The options an object's slide record carries: its frame, its Selection Pane name, and its alt
 * text and lock flags when the caller stated them.
 *
 * The name takes its index before the frame is resolved, so a definer calls this after its own
 * checks have passed: a throw before it leaves the kind's name counter where it was.
 * @param target - the slide the object is added to
 * @param type - the object's `_type`, which selects the name counter
 * @param opt - the caller's options
 * @param spec - `label` opens the default name and `kind` names the object in a name warning (see
 *   `resolveObjectName`); `api` opens a frame warning and `defaults` fills each axis left unstated
 *   (see `resolveAuthoredFrame`)
 */
export function framedObjectOptions(
	target: PresSlideInternal,
	type: SlideObjectType,
	opt: FramedObjectProps,
	spec: { label: string; kind: string; api: string; defaults: Record<'x' | 'y' | 'w' | 'h', Coord> }
): ObjectOptionsInternal {
	const objectName = resolveObjectName(target, type, { label: spec.label, kind: spec.kind, supplied: opt.objectName })
	return {
		...resolveAuthoredFrame(opt, spec.defaults, spec.api),
		objectName,
		...(opt.altText ? { altText: opt.altText } : {}),
		...(opt.objectLock ? { objectLock: opt.objectLock } : {}),
	}
}

/**
 * Refuse an embedded object that names neither a `path` nor `data`: the payload is the object, and
 * there is no default to embed in its place.
 * @param opt - the caller's options
 * @param code - the error's code, in the object's own namespace
 * @param api - the method the caller called, opening the message
 */
export function requirePayloadSource(
	opt: { path?: string | undefined; data?: string | undefined },
	code: Extract<InvalidOptionErrorCode, `${string}/missing-source`>,
	api: string
): void {
	if (!opt.path && !opt.data) throw new InvalidOptionError(code, `${api}(): either \`data\` or \`path\` are required!`)
}
