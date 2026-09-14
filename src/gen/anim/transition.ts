/**
 * ts-pptx: slide-show transitions
 *
 * Build the `<p:transition>` tree (positioned in CT_Slide between `p:clrMapOvr`
 * and `p:timing`): the transition-type element, its optional sound action, and
 * the `mc:AlternateContent` envelope that carries an exact `p14:dur`.
 */

import type { TransitionProps, TransitionType } from '../../types/index.js'
import type { PresSlideInternal } from '../../types/internal.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { xsdBoolIfTrue } from '../../ooxml/xsd-boolean.js'
import { checkEnumOrWarn } from '../../ooxml/check-enum.js'
import { TRANSITION_TYPES, TRANSITION_VARIANT_ATTRIBUTES } from '../../ooxml/st-enums.js'
import { warnOnce } from '../../diagnostics.js'

/** Map a `ST_TransitionSpeed`-less exact duration (ms) to PowerPoint's coarse `spd` bucket. */
function transitionSpeedForDuration(durationMs: number): 'slow' | 'med' | 'fast' {
	if (durationMs <= 500) return 'fast'
	if (durationMs <= 1000) return 'med'
	return 'slow'
}

/**
 * The attributes of a transition's type element, from the caller's `variant`.
 *
 * Each `variant` key is written as an attribute name, so a key the element does not declare would
 * be an attribute the schema does not allow there, and a key that is not a name at all would be
 * markup. Only the attributes the element declares are written, and the rest warn.
 */
function transitionVariantAttrs(type: TransitionType, variant: Record<string, string> | undefined): XmlAttrs | null {
	if (!variant) return null
	const allowed: readonly string[] = TRANSITION_VARIANT_ATTRIBUTES[type]
	const attrs: XmlAttrs = {}
	for (const [name, value] of Object.entries(variant)) {
		if (allowed.includes(name)) {
			attrs[name] = value
			continue
		}
		warnOnce(
			'transition/unknown-variant',
			`slide transition \`${type}\` variant \`${name}\` is not an attribute of <p:${type}> and is ignored — ${allowed.length > 0 ? `use ${allowed.join(', ')}` : 'it takes none'}.`,
			{ received: name, valid: allowed }
		)
	}
	return attrs
}

/**
 * Build the `p:sndAc` sound-action child of `p:transition` (positioned after the
 * transition-type element). A start sound is `<p:stSnd [loop="1"]><p:snd r:embed
 * name/></p:stSnd>` referencing the embedded audio by `soundRId`; the stop-previous form is
 * `<p:endSnd/>` (no rel). Returns `''` when the transition has no sound, or a start sound was
 * registered under no relationship on this slide.
 */
function transitionSoundToXml(transition: TransitionProps, soundRId: number | undefined): string {
	const sound = transition.sound
	if (!sound) return ''
	if (sound.stopPrevious) return el('p:sndAc', null, raw(voidEl('p:endSnd')))
	if (soundRId === undefined) return '' // no embedded part registered
	return el(
		'p:sndAc',
		null,
		raw(
			el(
				'p:stSnd',
				{ loop: xsdBoolIfTrue(sound.loop) },
				raw(voidEl('p:snd', { 'r:embed': `rId${soundRId}`, name: sound.name || null }))
			)
		)
	)
}

/**
 * Build the slide-show transition tree (`p:transition`), positioned in `CT_Slide`
 * between `p:clrMapOvr` and `p:timing`. Emits PowerPoint's `mc:AlternateContent`
 * form (a `p14` Choice carrying the exact `p14:dur`, plus a base `mc:Fallback`)
 * when `durationMs` is set, and the bare `<p:transition>` otherwise. See
 * `docs/animations-and-transitions.md`.
 * @param slide - the slide whose transition is emitted
 * @param soundRId - the relationship id this slide's start sound was registered under, if one was.
 *   It is the slide's, not the transition's: one transition object can sit on several slides, each
 *   declaring the sound under an id of its own.
 * @returns {string} the transition XML, or `''` when the slide has no transition the writer can emit
 */
export function slideTransitionToXml(slide: PresSlideInternal, soundRId?: number): string {
	const transition = slide.transition
	if (!transition?.type) return ''

	// The type is written as the element's name, so one that is not a transition this writer knows
	// warns and the slide is written with no transition at all.
	const type = checkEnumOrWarn(transition.type, TRANSITION_TYPES, 'transition/unknown-type', 'slide transition `type`')
	if (!type) return ''

	const typeEl = voidEl(`p:${type}`, transitionVariantAttrs(type, transition.variant))
	const sndAc = transitionSoundToXml(transition, soundRId)

	const hasDuration = typeof transition.durationMs === 'number' && Number.isFinite(transition.durationMs)
	const speed = transition.speed ?? (hasDuration ? transitionSpeedForDuration(transition.durationMs as number) : null)
	const baseAttrs: XmlAttrs = {
		spd: speed || null,
		advClick: transition.advanceOnClick === false ? '0' : null,
		advTm: typeof transition.advanceAfterMs === 'number' ? Math.round(transition.advanceAfterMs) : null,
	}

	if (!hasDuration) return el('p:transition', baseAttrs, [raw(typeEl), raw(sndAc)])

	const dur = Math.round(transition.durationMs as number)
	return el('mc:AlternateContent', { 'xmlns:mc': OOXML_NS.mc }, [
		raw(
			el(
				'mc:Choice',
				{ 'xmlns:p14': OOXML_NS.p14, Requires: 'p14' },
				raw(el('p:transition', { ...baseAttrs, 'p14:dur': dur }, [raw(typeEl), raw(sndAc)]))
			)
		),
		raw(el('mc:Fallback', null, raw(el('p:transition', baseAttrs, [raw(typeEl), raw(sndAc)])))),
	])
}
