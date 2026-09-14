/**
 * Typed read/edit model for a slide's `p:transition` (slide-show transition).
 *
 * The element sits in `CT_Slide` between `p:clrMapOvr` and `p:timing` and comes
 * in two shapes PowerPoint emits:
 *
 * - **bare** `<p:transition spd=… advClick=… advTm=…><p:TYPE …/></p:transition>`
 *   when the duration matches a coarse `spd` speed bucket; and
 * - **wrapped** `<mc:AlternateContent>` carrying a `<mc:Choice Requires="p14">`
 *   whose `p:transition` adds the precise `p14:dur` (milliseconds), plus a base
 *   `<mc:Fallback>` without it — emitted for an off-bucket exact duration.
 *
 * The getter prefers the `p14` Choice (so `durationMs` is recovered); the setter
 * emits the `mc:AlternateContent` form when `durationMs` is set and the bare form
 * otherwise. See `docs/contributing/design/animations.md`.
 */
import { InvalidOptionError } from '../../errors.js'
import { transitionSpeedForDuration } from '../../ooxml/transition-speed.js'
import {
	OOXML_NS,
	attr,
	boolValue,
	childElements,
	createElement,
	firstChild,
	numberValue,
	setAttr,
	type Document,
	type Element,
} from '../oxml/dom.js'

/** `ST_TransitionSpeed` — the coarse speed bucket (`spd`), default `fast`. */
export type TransitionSpeed = 'slow' | 'med' | 'fast'

/** Decoded slide transition (`p:transition`), as returned by {@link Slide.transition}. */
export interface TransitionInfo {
	/**
	 * The transition-type element's local name (e.g. `fade`, `push`, `wipe`,
	 * `cut`, `dissolve`). Base ECMA-376 types use the `p` namespace; modern
	 * PowerPoint types live in `p14`/`p15`/`p159` (see {@link namespace}).
	 */
	type: string
	/** The transition element's namespace prefix (`p` for base, else `p14`/`p15`/`p159`). */
	namespace: string
	/** Coarse speed bucket; `fast` when `spd` is absent (its schema default). */
	speed: TransitionSpeed
	/** Exact duration in milliseconds (`p14:dur`), or `null` when only a speed bucket is set. */
	durationMs: number | null
	/** Whether the slide advances on click (`advClick`, default `true`). */
	advanceOnClick: boolean
	/** Auto-advance delay in milliseconds (`advTm`), or `null` when not timed. */
	advanceAfterMs: number | null
	/** Type-specific variant attributes (e.g. `{ dir: 'd' }` for push, `{ spokes: '2' }` for wheel). */
	variant: Record<string, string>
	/** Transition sound (`p:sndAc`), or `null` when the transition is silent. */
	sound: TransitionSoundInfo | null
}

/** Decoded transition sound (`p:sndAc`). A start sound (`p:stSnd`) or the stop-previous form (`p:endSnd`). */
export interface TransitionSoundInfo {
	/** `start` for an embedded start sound (`p:stSnd`), `stop` for the stop-previous form (`p:endSnd`). */
	form: 'start' | 'stop'
	/** Whether the start sound loops until the next sound (`p:stSnd @loop`). Always `false` for `stop`. */
	loop: boolean
	/** Relationship id of the embedded WAV (`p:snd @r:embed`), or `null` (stop form / missing). */
	embedRid: string | null
	/** Display name on `p:snd @name`, or `null`. */
	name: string | null
}

/**
 * Transition fields accepted by the {@link Slide.transition} setter. A {@link TransitionInfo} read
 * from the same slide is a valid input, so `slide.transition = { ...slide.transition, speed: 'slow' }`
 * changes the speed and keeps everything else, the sound included.
 */
export interface TransitionInput {
	type: string
	/** Namespace prefix for the type element; defaults to `p` (base ECMA-376). */
	namespace?: string
	/**
	 * Coarse speed bucket (`spd`). When omitted it is derived from `durationMs` if one is given, and
	 * otherwise not written. A stated `fast`, the schema default, is not written when the slide's
	 * transition had no `spd` either, so a transition read back and assigned again keeps its markup.
	 */
	speed?: TransitionSpeed
	/** Exact duration in milliseconds (`p14:dur`), a finite number from 0; `null` or omitted writes none. */
	durationMs?: number | null
	advanceOnClick?: boolean
	/** Auto-advance delay in milliseconds (`advTm`), a finite number from 0; `null` or omitted writes none. */
	advanceAfterMs?: number | null
	variant?: Record<string, string>
	/**
	 * The transition sound (`p:sndAc`), in three states.
	 *
	 * - **omitted** keeps the sound the slide's transition already has;
	 * - **`null`** removes it;
	 * - **a value** is accepted only when it is that same sound, as a spread of the getter passes it.
	 *
	 * A different sound throws `transition/sound-unsupported`: a start sound names an embedded audio
	 * part, which this setter does not add, so writing one would point the slide at a relationship it
	 * may not have. Add a sound when authoring the deck (`slide.transition = { sound }` on the write
	 * side) instead.
	 */
	sound?: TransitionSoundInfo | null
}

/** Child element names of `p:transition` that are not the transition-type choice. */
const NON_TYPE_CHILDREN = new Set(['sndAc', 'extLst'])

/** The single transition-type child element (the choice), or `null`. */
function typeElement(transition: Element): Element | null {
	for (const child of childElements(transition)) {
		if (!NON_TYPE_CHILDREN.has(child.localName ?? '')) return child
	}
	return null
}

/** Namespace prefix for a URI, defaulting to `p` for the base presentationml namespace. */
function prefixFor(uri: string | null): string {
	if (!uri) return 'p'
	for (const [prefix, ns] of Object.entries(OOXML_NS)) {
		if (ns === uri) return prefix
	}
	return 'p'
}

/** Collect a transition-type element's non-namespace attributes by local name. */
function variantAttrs(element: Element): Record<string, string> {
	const out: Record<string, string> = {}
	const attrs = element.attributes
	for (let i = 0; i < attrs.length; i++) {
		const a = attrs[i]
		if (!a || a.prefix === 'xmlns' || a.name === 'xmlns') continue
		out[a.localName ?? a.name] = a.value
	}
	return out
}

/**
 * Locate the slide's transition under the root `p:sld`, returning the
 * `p:transition` element to decode (the `p14` Choice when wrapped) and the
 * outermost node that represents it (the bare element or the `mc:AlternateContent`
 * wrapper), or `null` when the slide has no transition.
 */
export function findTransition(root: Element): { transition: Element; outer: Element } | null {
	const bare = firstChild(root, 'p:transition')
	if (bare) return { transition: bare, outer: bare }

	const altContent = firstChild(root, 'mc:AlternateContent')
	if (altContent) {
		const choice = firstChild(altContent, 'mc:Choice')
		const transition = choice ? firstChild(choice, 'p:transition') : null
		if (transition) return { transition, outer: altContent }
		// Degenerate AlternateContent with only a Fallback — still surface it.
		const fallback = firstChild(altContent, 'mc:Fallback')
		const fbTransition = fallback ? firstChild(fallback, 'p:transition') : null
		if (fbTransition) return { transition: fbTransition, outer: altContent }
	}
	return null
}

/** Decode a `p:transition` element into a {@link TransitionInfo}, or `null` when malformed. */
export function parseTransition(root: Element): TransitionInfo | null {
	const found = findTransition(root)
	if (!found) return null
	const { transition } = found
	const type = typeElement(transition)
	if (!type) return null

	const advClick = boolValue(attr(transition, 'advClick'))
	const advTm = numberValue(attr(transition, 'advTm'))
	const spd = attr(transition, 'spd')
	return {
		type: type.localName ?? '',
		namespace: prefixFor(type.namespaceURI),
		speed: spd === 'slow' || spd === 'med' || spd === 'fast' ? spd : 'fast',
		durationMs: numberValue(attr(transition, 'p14:dur')),
		advanceOnClick: advClick ?? true,
		advanceAfterMs: advTm,
		variant: variantAttrs(type),
		sound: parseSound(transition),
	}
}

/** Decode the `p:sndAc` sound-action child of a transition into a {@link TransitionSoundInfo}, or `null`. */
function parseSound(transition: Element): TransitionSoundInfo | null {
	const sndAc = firstChild(transition, 'p:sndAc')
	if (!sndAc) return null
	const stSnd = firstChild(sndAc, 'p:stSnd')
	if (stSnd) {
		const snd = firstChild(stSnd, 'p:snd')
		return {
			form: 'start',
			loop: boolValue(attr(stSnd, 'loop')) ?? false,
			embedRid: snd ? attr(snd, 'r:embed') : null,
			name: snd ? attr(snd, 'name') : null,
		}
	}
	if (firstChild(sndAc, 'p:endSnd')) return { form: 'stop', loop: false, embedRid: null, name: null }
	return null
}

/**
 * A transition time the caller stated, as whole milliseconds, or `null` when not stated. `NaN`,
 * `Infinity` and a negative number throw: each was written straight into `p14:dur` or `advTm`.
 */
function transitionTime(value: number | null | undefined, name: 'durationMs' | 'advanceAfterMs'): number | null {
	if (value === undefined || value === null) return null
	if (typeof value !== 'number' || !Number.isFinite(value) || value < 0)
		throw new InvalidOptionError(
			'transition/invalid-time',
			`slide.transition: \`${name}\` is a number of milliseconds from 0; got ${String(value)}.`
		)
	return Math.round(value)
}

/**
 * The `p:sndAc` the new transition carries: the current one when the caller left `sound` out or
 * passed that same sound back, none for `null`, and a refusal for any other sound. See
 * {@link TransitionInput.sound}.
 */
function soundFor(input: TransitionSoundInfo | null | undefined, current: Element | null): Element | null {
	const existing = current ? firstChild(current, 'p:sndAc') : null
	if (input === undefined) return existing
	if (input === null) return null
	const now = current ? parseSound(current) : null
	const same =
		!!existing &&
		!!now &&
		now.form === input.form &&
		now.loop === !!input.loop &&
		now.embedRid === (input.embedRid ?? null) &&
		now.name === (input.name ?? null)
	if (same) return existing
	throw new InvalidOptionError(
		'transition/sound-unsupported',
		"slide.transition: `sound` can keep the slide's own transition sound (leave it out, or pass the one `slide.transition` reads) or remove it (`null`); adding a different sound is not supported here."
	)
}

/** What {@link buildTransitionElement} writes, resolved and checked once for both forms. */
interface ResolvedTransition {
	speed: TransitionSpeed | null
	durationMs: number | null
	advanceAfterMs: number | null
	sound: Element | null
}

/** Build a `p:transition` element (without the `p14:dur` attribute unless `withDur`) from an input. */
function buildTransitionElement(
	doc: Document,
	input: TransitionInput,
	resolved: ResolvedTransition,
	withDur: boolean
): Element {
	const transition = createElement(doc, 'p:transition')
	if (resolved.speed) setAttr(transition, 'spd', resolved.speed)
	if (withDur && resolved.durationMs !== null) setAttr(transition, 'p14:dur', String(resolved.durationMs))
	if (input.advanceOnClick === false) setAttr(transition, 'advClick', '0')
	if (resolved.advanceAfterMs !== null) setAttr(transition, 'advTm', String(resolved.advanceAfterMs))

	const prefix = input.namespace ?? 'p'
	const type = createElement(doc, `${prefix}:${input.type}`)
	for (const [name, value] of Object.entries(input.variant ?? {})) setAttr(type, name, value)
	transition.appendChild(type)
	// `CT_SlideTransition` orders the sound after the type choice.
	if (resolved.sound) transition.appendChild(resolved.sound.cloneNode(true))
	return transition
}

/**
 * Build the DOM node for a transition: the bare `p:transition` when no exact
 * duration is requested, or an `mc:AlternateContent` wrapper (a `p14` Choice
 * carrying `p14:dur` plus a base `mc:Fallback`) when `durationMs` is set.
 *
 * Every value is checked before anything is built, so a refused input changes nothing.
 * @param doc - the slide part's document
 * @param input - the caller's transition
 * @param root - the slide's root `p:sld`, whose current transition supplies the sound and the
 *   `spd` a round trip keeps; omit it for a slide with no transition to carry over
 */
export function buildTransition(doc: Document, input: TransitionInput, root?: Element | null): Element {
	const current = root ? (findTransition(root)?.transition ?? null) : null
	const durationMs = transitionTime(input.durationMs, 'durationMs')
	const advanceAfterMs = transitionTime(input.advanceAfterMs, 'advanceAfterMs')
	const sound = soundFor(input.sound, current)
	const stated = input.speed ?? (durationMs !== null ? transitionSpeedForDuration(durationMs) : null)
	// `fast` is the schema default. Keeping it absent where it was absent is what lets a transition
	// read back (whose `speed` reports `fast` for a missing `spd`) be assigned again unchanged.
	const keepsDefaultAbsent = stated === 'fast' && !!current && attr(current, 'spd') === null
	const resolved: ResolvedTransition = { speed: keepsDefaultAbsent ? null : stated, durationMs, advanceAfterMs, sound }

	if (durationMs === null) return buildTransitionElement(doc, input, resolved, false)

	const alt = createElement(doc, 'mc:AlternateContent')
	const choice = createElement(doc, 'mc:Choice')
	// `Requires` names a prefix, so the prefix is declared where `Requires` is read, as PowerPoint
	// writes it. Declared only on the inner `p:transition` (which `p14:dur` does), it was out of
	// scope on a slide whose root does not declare it, and the schema refused the choice.
	choice.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:p14', OOXML_NS.p14)
	setAttr(choice, 'Requires', 'p14')
	choice.appendChild(buildTransitionElement(doc, input, resolved, true))
	alt.appendChild(choice)
	const fallback = createElement(doc, 'mc:Fallback')
	fallback.appendChild(buildTransitionElement(doc, input, resolved, false))
	alt.appendChild(fallback)
	return alt
}

/** Remove any existing transition node (bare or wrapped) from the slide root. */
export function removeTransition(root: Element): void {
	const found = findTransition(root)
	if (found) root.removeChild(found.outer)
}
