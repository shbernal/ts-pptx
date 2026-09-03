/**
 * ts-pptx: slide build animations
 *
 * Resolve preset build animations to their target shape ids and assemble the
 * `mainSeq` / `bldLst` trees that PowerPoint reads for click-triggered entrance,
 * emphasis and exit effects. Consumed by the slide `<p:timing>` builder.
 */

import type { AnimationProps } from '../../types/index.js'
import type { SlideObject } from '../../types/internal.js'
import { warn } from '../../diagnostics.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { renderedSlideObjects, resolveObjectNameToId } from '../slide/shape-ids.js'

/**
 * Resolved animation target's shape id (`spid`), or `null` when it cannot be resolved — in which
 * case the effect is dropped, so every `null` warns rather than leaving the animation silently
 * missing from the deck.
 *
 * `objectName` resolves through `shapeIds`, which covers group children: they are `<p:cNvPr>`-named
 * on the slide and animate like any other shape, but are not in `_slideObjects`, so the old lookup
 * there dropped every animation targeting one. The name goes in raw — `resolveObjectNameToId` owns
 * matching it against the attribute-escaped form the slide object stores — so the warning below
 * quotes the same spelling the caller passed.
 *
 * `shapeIndex` is a 0-based index into the top-level objects that RENDER, which is the same
 * sequence `collectSlideShapeIds` allocates along; group children take ids past that range. It
 * counted `_slideObjects` itself, and four of that array's members draw nothing — so `addNotes`
 * before the first shape shifted every index by one and `shapeIndex: 0` addressed the notes.
 * An index outside `[0, renderedCount)` would emit a `<p:spTgt spid>` naming no shape on the
 * slide — a dangling spid PowerPoint reports as a repair (0x80070570) — so it warns and drops,
 * exactly like an unresolvable `objectName` does.
 * @param shapeIds - the slide's shape ids, from `collectSlideShapeIds`
 * @param slideObjects - the slide's top-level objects, which a `shapeIndex` addresses
 * @param anim - the animation to resolve
 * @returns the target's `<p:cNvPr>` id, or `null`
 */
export function resolveAnimationSpid(
	shapeIds: Map<SlideObject, number>,
	slideObjects: SlideObject[],
	anim: AnimationProps
): number | null {
	if (typeof anim.shapeIndex === 'number') {
		// Through the map, not `shapeIndex + 2`: the id basis belongs to one allocator.
		const targets = renderedSlideObjects(slideObjects)
		const target = targets[anim.shapeIndex]
		if (target) return shapeIds.get(target) ?? null
		warn(
			'animation/target-index-out-of-range',
			`addAnimation: shapeIndex ${anim.shapeIndex} is out of range (slide has ${targets.length} top-level shape(s)), so its "${anim.preset}" effect was dropped.`
		)
		return null
	}
	if (anim.objectName) {
		const id = resolveObjectNameToId(shapeIds, anim.objectName)
		if (id !== null) return id
		warn(
			'animation/target-not-found',
			`addAnimation: no object named "${anim.objectName}" on the slide, so its "${anim.preset}" effect was dropped.`
		)
		return null
	}
	warn(
		'animation/target-missing',
		`addAnimation: the "${anim.preset}" effect names no target (pass shapeIndex or objectName), so it was dropped.`
	)
	return null
}

interface AnimPresetMeta {
	presetID: number
	presetClass: 'entr' | 'emph' | 'exit'
	presetSubtype: number
	defaultDurationMs: number
	/** Emit the effect's behavior nodes (each `<p:cTn>` consumes one id via `next`). */
	behaviors: (spid: number, dur: number, next: () => number) => string
}

const ANIM_SET_VISIBLE = (spid: number, next: () => number): string =>
	el('p:set', null, [
		raw(
			el('p:cBhvr', null, [
				raw(
					el(
						'p:cTn',
						{ id: next(), dur: 1, fill: 'hold' },
						raw(el('p:stCondLst', null, raw(voidEl('p:cond', { delay: 0 }))))
					)
				),
				raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
				raw(el('p:attrNameLst', null, raw(el('p:attrName', null, 'style.visibility')))),
			])
		),
		raw(el('p:to', null, raw(voidEl('p:strVal', { val: 'visible' })))),
	])

const ANIM_SET_HIDDEN = (spid: number, dur: number, next: () => number): string =>
	el('p:set', null, [
		raw(
			el('p:cBhvr', null, [
				raw(
					el(
						'p:cTn',
						{ id: next(), dur: 1, fill: 'hold' },
						raw(el('p:stCondLst', null, raw(voidEl('p:cond', { delay: Math.max(0, dur - 1) }))))
					)
				),
				raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
				raw(el('p:attrNameLst', null, raw(el('p:attrName', null, 'style.visibility')))),
			])
		),
		raw(el('p:to', null, raw(voidEl('p:strVal', { val: 'hidden' })))),
	])

/** `p:animEffect` filter transition (fade/wipe(down)/…), shared by entrance and exit presets. */
const ANIM_EFFECT = (transition: 'in' | 'out', filter: string, spid: number, dur: number, next: () => number): string =>
	el(
		'p:animEffect',
		{ transition, filter },
		raw(
			el('p:cBhvr', null, [
				raw(voidEl('p:cTn', { id: next(), dur })),
				raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
			])
		)
	)

const ANIM_FADE = (transition: 'in' | 'out', spid: number, dur: number, next: () => number): string =>
	ANIM_EFFECT(transition, 'fade', spid, dur, next)

/**
 * One `ppt_x`/`ppt_y` motion `<p:anim>` for Fly In (entrance) / Fly Out (exit).
 * PowerPoint authors the two directions differently (captured in
 * `slide-animation-presets.oracle.json`): the entrance form references the
 * hashed run-time variables (`#ppt_x`) and holds (`fill="hold"`) starting
 * off-screen and ending in place; the exit form uses the bare variables
 * (`ppt_x`, no `#`), omits `fill="hold"`, and ends off-screen.
 */
const ANIM_FLY_AXIS = (
	direction: 'in' | 'out',
	axis: 'x' | 'y',
	spid: number,
	dur: number,
	next: () => number
): string => {
	const h = direction === 'in' ? '#' : ''
	let from: string
	let to: string
	if (axis === 'x') {
		from = `${h}ppt_x`
		to = `${h}ppt_x`
	} else if (direction === 'in') {
		from = `1+${h}ppt_h/2`
		to = `${h}ppt_y`
	} else {
		from = `${h}ppt_y`
		to = `1+${h}ppt_h/2`
	}
	const fill = direction === 'in' ? 'hold' : null
	return el('p:anim', { calcmode: 'lin', valueType: 'num' }, [
		raw(
			el('p:cBhvr', { additive: 'base' }, [
				raw(voidEl('p:cTn', { id: next(), dur, fill })),
				raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
				raw(el('p:attrNameLst', null, raw(el('p:attrName', null, `ppt_${axis}`)))),
			])
		),
		raw(
			el('p:tavLst', null, [
				raw(el('p:tav', { tm: 0 }, raw(el('p:val', null, raw(voidEl('p:strVal', { val: from })))))),
				raw(el('p:tav', { tm: 100000 }, raw(el('p:val', null, raw(voidEl('p:strVal', { val: to })))))),
			])
		),
	])
}

const ANIM_PRESETS: Record<string, AnimPresetMeta> = {
	fadeIn: {
		presetID: 10,
		presetClass: 'entr',
		presetSubtype: 0,
		defaultDurationMs: 500,
		behaviors: (spid, dur, next) => ANIM_SET_VISIBLE(spid, next) + ANIM_FADE('in', spid, dur, next),
	},
	flyIn: {
		presetID: 2,
		presetClass: 'entr',
		presetSubtype: 4,
		defaultDurationMs: 500,
		behaviors: (spid, dur, next) =>
			ANIM_SET_VISIBLE(spid, next) +
			ANIM_FLY_AXIS('in', 'x', spid, dur, next) +
			ANIM_FLY_AXIS('in', 'y', spid, dur, next),
	},
	appear: {
		presetID: 1,
		presetClass: 'entr',
		presetSubtype: 0,
		defaultDurationMs: 500,
		behaviors: (spid, _dur, next) => ANIM_SET_VISIBLE(spid, next),
	},
	wipe: {
		presetID: 22,
		presetClass: 'entr',
		presetSubtype: 4,
		defaultDurationMs: 500,
		behaviors: (spid, dur, next) => ANIM_SET_VISIBLE(spid, next) + ANIM_EFFECT('in', 'wipe(down)', spid, dur, next),
	},
	grow: {
		presetID: 6,
		presetClass: 'emph',
		presetSubtype: 0,
		defaultDurationMs: 2000,
		behaviors: (spid, dur, next) =>
			el('p:animScale', null, [
				raw(
					el('p:cBhvr', null, [
						raw(voidEl('p:cTn', { id: next(), dur, fill: 'hold' })),
						raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
					])
				),
				raw(voidEl('p:by', { x: 150000, y: 150000 })),
			]),
	},
	spin: {
		presetID: 8,
		presetClass: 'emph',
		presetSubtype: 0,
		defaultDurationMs: 2000,
		behaviors: (spid, dur, next) =>
			el('p:animRot', { by: 21600000 }, [
				raw(
					el('p:cBhvr', null, [
						raw(voidEl('p:cTn', { id: next(), dur, fill: 'hold' })),
						raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
						raw(el('p:attrNameLst', null, raw(el('p:attrName', null, 'r')))),
					])
				),
			]),
	},
	fadeOut: {
		presetID: 10,
		presetClass: 'exit',
		presetSubtype: 0,
		defaultDurationMs: 500,
		behaviors: (spid, dur, next) => ANIM_FADE('out', spid, dur, next) + ANIM_SET_HIDDEN(spid, dur, next),
	},
	flyOut: {
		presetID: 2,
		presetClass: 'exit',
		presetSubtype: 4,
		defaultDurationMs: 500,
		behaviors: (spid, dur, next) =>
			ANIM_FLY_AXIS('out', 'x', spid, dur, next) +
			ANIM_FLY_AXIS('out', 'y', spid, dur, next) +
			ANIM_SET_HIDDEN(spid, dur, next),
	},
}

const ANIM_NODE_TYPE: Record<string, string> = {
	onClick: 'clickEffect',
	withPrevious: 'withEffect',
	afterPrevious: 'afterEffect',
}

type ResolvedAnimation = { anim: AnimationProps; spid: number }

/**
 * Assemble the `mainSeq` from preset effects. Effects are grouped into click
 * steps: an `onClick` effect opens a new click group; `afterPrevious` opens a new
 * sub-step (delayed by the previous effect's duration) within it; `withPrevious`
 * joins the current sub-step. This reproduces PowerPoint's interactive build tree.
 */
export function buildAnimationSeq(animations: ResolvedAnimation[], next: () => number): string {
	interface SubGroup {
		delay: number
		effects: ResolvedAnimation[]
	}
	interface ClickGroup {
		subs: SubGroup[]
	}
	const groups: ClickGroup[] = []
	let prevDuration = 0
	for (const entry of animations) {
		const meta = ANIM_PRESETS[entry.anim.preset]
		if (!meta) continue
		const trigger = entry.anim.trigger ?? 'onClick'
		const duration = typeof entry.anim.durationMs === 'number' ? entry.anim.durationMs : meta.defaultDurationMs
		if (trigger === 'onClick' || groups.length === 0) {
			groups.push({ subs: [{ delay: 0, effects: [entry] }] })
		} else if (trigger === 'afterPrevious') {
			groups[groups.length - 1]?.subs.push({ delay: prevDuration, effects: [entry] })
		} else {
			// withPrevious — join the current sub-step
			const subs = groups[groups.length - 1]?.subs
			const lastSub = subs?.[subs.length - 1]
			if (lastSub) lastSub.effects.push(entry)
		}
		prevDuration = duration
	}

	const emitEffect = (entry: ResolvedAnimation): string => {
		const meta = ANIM_PRESETS[entry.anim.preset]
		if (!meta) return ''
		const nodeType = ANIM_NODE_TYPE[entry.anim.trigger ?? 'onClick']
		const duration = typeof entry.anim.durationMs === 'number' ? entry.anim.durationMs : meta.defaultDurationMs
		const effectId = next()
		const behaviors = meta.behaviors(entry.spid, duration, next)
		return el(
			'p:par',
			null,
			raw(
				el(
					'p:cTn',
					{
						id: effectId,
						presetID: meta.presetID,
						presetClass: meta.presetClass,
						presetSubtype: meta.presetSubtype,
						fill: 'hold',
						grpId: 0,
						nodeType,
					},
					[
						raw(el('p:stCondLst', null, raw(voidEl('p:cond', { delay: 0 })))),
						raw(el('p:childTnLst', null, raw(behaviors))),
					]
				)
			)
		)
	}

	const emitSub = (sub: SubGroup): string => {
		const subId = next()
		const effects = sub.effects.map(emitEffect).join('')
		return el(
			'p:par',
			null,
			raw(
				el('p:cTn', { id: subId, fill: 'hold' }, [
					raw(el('p:stCondLst', null, raw(voidEl('p:cond', { delay: sub.delay })))),
					raw(el('p:childTnLst', null, raw(effects))),
				])
			)
		)
	}

	const emitGroup = (group: ClickGroup): string => {
		const groupId = next()
		const subs = group.subs.map(emitSub).join('')
		return el(
			'p:par',
			null,
			raw(
				el('p:cTn', { id: groupId, fill: 'hold' }, [
					raw(el('p:stCondLst', null, raw(voidEl('p:cond', { delay: 'indefinite' })))),
					raw(el('p:childTnLst', null, raw(subs))),
				])
			)
		)
	}

	const clickGroups = groups.map(emitGroup).join('')
	return el('p:seq', { concurrent: 1, nextAc: 'seek' }, [
		raw(
			el('p:cTn', { id: 2, dur: 'indefinite', nodeType: 'mainSeq' }, raw(el('p:childTnLst', null, raw(clickGroups))))
		),
		raw(
			el(
				'p:prevCondLst',
				null,
				raw(el('p:cond', { evt: 'onPrev', delay: 0 }, raw(el('p:tgtEl', null, raw(voidEl('p:sldTgt'))))))
			)
		),
		raw(
			el(
				'p:nextCondLst',
				null,
				raw(el('p:cond', { evt: 'onNext', delay: 0 }, raw(el('p:tgtEl', null, raw(voidEl('p:sldTgt'))))))
			)
		),
	])
}

/** One `<p:bldP>` per animated shape, in order of first appearance. */
export function buildBldList(animations: ResolvedAnimation[]): string {
	const seen = new Set<number>()
	const bldPs: string[] = []
	for (const { spid } of animations) {
		if (seen.has(spid)) continue
		seen.add(spid)
		bldPs.push(voidEl('p:bldP', { spid, grpId: 0 }))
	}
	return el('p:bldLst', null, raw(bldPs.join('')))
}
