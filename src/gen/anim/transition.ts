/**
 * PptxGenJS: slide-show transitions
 *
 * Build the `<p:transition>` tree (positioned in CT_Slide between `p:clrMapOvr`
 * and `p:timing`): the transition-type element, its optional sound action, and
 * the `mc:AlternateContent` envelope that carries an exact `p14:dur`.
 */

import type { PresSlideInternal, TransitionProps } from '../../core-interfaces.js'
import { encodeXmlEntities } from '../../gen-utils.js'

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
	if (sound.stopPrevious) return '<p:sndAc><p:endSnd/></p:sndAc>'
	if (typeof transition._sndRId !== 'number') return '' // no embedded part registered
	const loopAttr = sound.loop ? ' loop="1"' : ''
	const nameAttr = sound.name ? ` name="${encodeXmlEntities(sound.name)}"` : ''
	return `<p:sndAc><p:stSnd${loopAttr}><p:snd r:embed="rId${transition._sndRId}"${nameAttr}/></p:stSnd></p:sndAc>`
}

export function slideTransitionToXml(slide: PresSlideInternal): string {
	const transition = slide.transition
	if (!transition?.type) return ''

	const variantAttrs = Object.entries(transition.variant ?? {})
		.map(([name, value]) => ` ${name}="${encodeXmlEntities(String(value))}"`)
		.join('')
	const typeEl = `<p:${transition.type}${variantAttrs}/>`
	const sndAc = transitionSoundToXml(transition)

	const hasDuration = typeof transition.durationMs === 'number' && isFinite(transition.durationMs)
	const speed = transition.speed ?? (hasDuration ? transitionSpeedForDuration(transition.durationMs as number) : null)
	const baseAttrs =
		`${speed ? ` spd="${speed}"` : ''}` +
		`${transition.advanceOnClick === false ? ' advClick="0"' : ''}` +
		`${typeof transition.advanceAfterMs === 'number' ? ` advTm="${Math.round(transition.advanceAfterMs)}"` : ''}`

	if (!hasDuration) return `<p:transition${baseAttrs}>${typeEl}${sndAc}</p:transition>`

	const dur = Math.round(transition.durationMs as number)
	return (
		'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
		'<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">' +
		`<p:transition${baseAttrs} p14:dur="${dur}">${typeEl}${sndAc}</p:transition>` +
		'</mc:Choice>' +
		'<mc:Fallback>' +
		`<p:transition${baseAttrs}>${typeEl}${sndAc}</p:transition>` +
		'</mc:Fallback>' +
		'</mc:AlternateContent>'
	)
}
