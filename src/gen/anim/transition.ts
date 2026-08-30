/**
 * ts-pptx: slide-show transitions
 *
 * Build the `<p:transition>` tree (positioned in CT_Slide between `p:clrMapOvr`
 * and `p:timing`): the transition-type element, its optional sound action, and
 * the `mc:AlternateContent` envelope that carries an exact `p14:dur`.
 */

import type { TransitionProps } from '../../types/index.js'
import type { PresSlideInternal } from '../../types/internal.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'

/** Map a `ST_TransitionSpeed`-less exact duration (ms) to PowerPoint's coarse `spd` bucket. */
function transitionSpeedForDuration(durationMs: number): 'slow' | 'med' | 'fast' {
	if (durationMs <= 500) return 'fast'
	if (durationMs <= 1000) return 'med'
	return 'slow'
}

/**
 * Build the slide-show transition tree (`p:transition`), positioned in `CT_Slide`
 * between `p:clrMapOvr` and `p:timing`. Emits PowerPoint's `mc:AlternateContent`
 * form (a `p14` Choice carrying the exact `p14:dur`, plus a base `mc:Fallback`)
 * when `durationMs` is set, and the bare `<p:transition>` otherwise. See
 * `docs/animations-and-transitions.md`.
 * @returns {string} the transition XML, or `''` when the slide has no transition
 */
/**
 * Build the `p:sndAc` sound-action child of `p:transition` (positioned after the
 * transition-type element). A start sound is `<p:stSnd [loop="1"]><p:snd r:embed
 * name/></p:stSnd>` referencing the embedded WAV by the relationship id stamped on
 * `transition._sndRId`; the stop-previous form is `<p:endSnd/>` (no rel). Returns
 * `''` when the transition has no sound.
 */
function transitionSoundToXml(transition: TransitionProps): string {
	const sound = transition.sound
	if (!sound) return ''
	if (sound.stopPrevious) return el('p:sndAc', null, raw(voidEl('p:endSnd')))
	if (typeof transition._sndRId !== 'number') return '' // no embedded part registered
	return el(
		'p:sndAc',
		null,
		raw(
			el(
				'p:stSnd',
				{ loop: sound.loop ? '1' : null },
				raw(voidEl('p:snd', { 'r:embed': `rId${transition._sndRId}`, name: sound.name || null }))
			)
		)
	)
}

export function slideTransitionToXml(slide: PresSlideInternal): string {
	const transition = slide.transition
	if (!transition?.type) return ''

	const typeEl = voidEl(`p:${transition.type}`, transition.variant ?? null)
	const sndAc = transitionSoundToXml(transition)

	const hasDuration = typeof transition.durationMs === 'number' && Number.isFinite(transition.durationMs)
	const speed = transition.speed ?? (hasDuration ? transitionSpeedForDuration(transition.durationMs as number) : null)
	const baseAttrs: XmlAttrs = {
		spd: speed || null,
		advClick: transition.advanceOnClick === false ? '0' : null,
		advTm: typeof transition.advanceAfterMs === 'number' ? Math.round(transition.advanceAfterMs) : null,
	}

	if (!hasDuration) return el('p:transition', baseAttrs, [raw(typeEl), raw(sndAc)])

	const dur = Math.round(transition.durationMs as number)
	return el('mc:AlternateContent', { 'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006' }, [
		raw(
			el(
				'mc:Choice',
				{ 'xmlns:p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main', Requires: 'p14' },
				raw(el('p:transition', { ...baseAttrs, 'p14:dur': dur }, [raw(typeEl), raw(sndAc)]))
			)
		),
		raw(el('mc:Fallback', null, raw(el('p:transition', baseAttrs, [raw(typeEl), raw(sndAc)])))),
	])
}
