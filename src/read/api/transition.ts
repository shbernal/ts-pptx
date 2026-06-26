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
 * otherwise. See `docs/animations-and-transitions.md`.
 */
import {
	OOXML_NS,
	attr,
	boolValue,
	childElements,
	createElement,
	firstChild,
	intValue,
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

/** Transition fields accepted by the {@link Slide.transition} setter. `speed` defaults are derived from `durationMs`. */
export interface TransitionInput {
	type: string
	/** Namespace prefix for the type element; defaults to `p` (base ECMA-376). */
	namespace?: string
	speed?: TransitionSpeed
	durationMs?: number | null
	advanceOnClick?: boolean
	advanceAfterMs?: number | null
	variant?: Record<string, string>
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
	// Modern transition namespaces not in the canonical registry.
	if (uri === 'http://schemas.microsoft.com/office/powerpoint/2012/main') return 'p15'
	if (uri === 'http://schemas.microsoft.com/office/powerpoint/2015/9/main') return 'p159'
	return 'p'
}

/** Collect a transition-type element's non-namespace attributes by local name. */
function variantAttrs(element: Element): Record<string, string> {
	const out: Record<string, string> = {}
	const attrs = element.attributes
	for (let i = 0; i < attrs.length; i++) {
		const a = attrs[i]
		if (a.prefix === 'xmlns' || a.name === 'xmlns') continue
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
	const advTm = intValue(attr(transition, 'advTm'))
	const spd = attr(transition, 'spd')
	return {
		type: type.localName ?? '',
		namespace: prefixFor(type.namespaceURI),
		speed: spd === 'slow' || spd === 'med' || spd === 'fast' ? spd : 'fast',
		durationMs: intValue(attr(transition, 'p14:dur')),
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

/** Map an exact duration (ms) to PowerPoint's coarse speed bucket. */
function speedForDuration(durationMs: number): TransitionSpeed {
	if (durationMs <= 500) return 'fast'
	if (durationMs <= 1000) return 'med'
	return 'slow'
}

/** Build a `p:transition` element (without the `p14:dur` attribute) from an input. */
function buildTransitionElement(
	doc: Document,
	input: TransitionInput,
	speed: TransitionSpeed | null,
	withDur: boolean
): Element {
	const transition = createElement(doc, 'p:transition')
	if (speed) setAttr(transition, 'spd', speed)
	if (withDur && typeof input.durationMs === 'number')
		setAttr(transition, 'p14:dur', String(Math.round(input.durationMs)))
	if (input.advanceOnClick === false) setAttr(transition, 'advClick', '0')
	if (typeof input.advanceAfterMs === 'number') setAttr(transition, 'advTm', String(Math.round(input.advanceAfterMs)))

	const prefix = input.namespace ?? 'p'
	const type = createElement(doc, `${prefix}:${input.type}`)
	for (const [name, value] of Object.entries(input.variant ?? {})) setAttr(type, name, value)
	transition.appendChild(type)
	return transition
}

/**
 * Build the DOM node for a transition: the bare `p:transition` when no exact
 * duration is requested, or an `mc:AlternateContent` wrapper (a `p14` Choice
 * carrying `p14:dur` plus a base `mc:Fallback`) when `durationMs` is set.
 */
export function buildTransition(doc: Document, input: TransitionInput): Element {
	const hasDuration = typeof input.durationMs === 'number'
	const speed = input.speed ?? (hasDuration ? speedForDuration(input.durationMs as number) : null)

	if (!hasDuration) return buildTransitionElement(doc, input, speed, false)

	const alt = createElement(doc, 'mc:AlternateContent')
	const choice = createElement(doc, 'mc:Choice')
	setAttr(choice, 'Requires', 'p14')
	choice.appendChild(buildTransitionElement(doc, input, speed, true))
	alt.appendChild(choice)
	const fallback = createElement(doc, 'mc:Fallback')
	fallback.appendChild(buildTransitionElement(doc, input, speed, false))
	alt.appendChild(fallback)
	return alt
}

/** Remove any existing transition node (bare or wrapped) from the slide root. */
export function removeTransition(root: Element): void {
	const found = findTransition(root)
	if (found) root.removeChild(found.outer)
}
