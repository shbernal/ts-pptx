/**
 * Slide transition and animation types.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */

/**
 * Base ECMA-376 slide-transition type (`p:transition`'s single type child). Each
 * maps to a `<p:TYPE/>` element; type-specific variants (e.g. direction) go in
 * {@link TransitionProps.variant}. Modern PowerPoint-only transitions (Morph,
 * Vortex, …) live in the `p14`/`p15`/`p159` namespaces and are out of authoring
 * scope for now. See `docs/animations-and-transitions.md`.
 */
export type TransitionType =
	| 'blinds'
	| 'checker'
	| 'circle'
	| 'comb'
	| 'cover'
	| 'cut'
	| 'diamond'
	| 'dissolve'
	| 'fade'
	| 'newsflash'
	| 'plus'
	| 'pull'
	| 'push'
	| 'random'
	| 'randomBar'
	| 'split'
	| 'strips'
	| 'wedge'
	| 'wheel'
	| 'wipe'
	| 'zoom'

/**
 * Slide-show transition applied between slides (`p:transition`). Assign to
 * {@link Slide.transition}. Setting `durationMs` emits PowerPoint's
 * `mc:AlternateContent` form (a `p14` Choice carrying `p14:dur` plus a base
 * `mc:Fallback`); otherwise only the coarse `speed` bucket is written.
 * @example slide.transition = { type: 'push', durationMs: 1250, variant: { dir: 'd' } }
 */
export interface TransitionProps {
	/** Transition type (the `<p:TYPE/>` element), e.g. `fade`, `push`, `wipe`, `cut`, `dissolve`. */
	type: TransitionType
	/** Exact duration in milliseconds (`p14:dur`); emits the `mc:AlternateContent` form. */
	durationMs?: number
	/** Coarse speed bucket (`spd`); derived from `durationMs` when omitted, else defaults to `fast`. */
	speed?: 'slow' | 'med' | 'fast'
	/** Advance on mouse click (`advClick`). @default true */
	advanceOnClick?: boolean
	/** Auto-advance after this many milliseconds (`advTm`). */
	advanceAfterMs?: number
	/** Type-specific variant attributes, e.g. `{ dir: 'd' }` for push, `{ spokes: '2' }` for wheel. */
	variant?: Record<string, string>
	/** Sound played with the transition (`p:sndAc`): a start sound (embedded WAV) or the stop-previous form. */
	sound?: TransitionSoundProps
	/**
	 * Internal: the slide relationship id assigned to the embedded sound part, stamped
	 * by the export-time registration pass. Not part of the authoring surface.
	 */
	_sndRId?: number
}

/**
 * A transition sound (`p:sndAc`). Either a **start sound** — an embedded WAV played
 * when the transition runs (`p:stSnd` → `p:snd`), optionally looped — or the
 * **stop-previous** form (`p:endSnd`) that silences a still-playing transition
 * sound. Built-in PowerPoint sounds embed identically to a custom file, so there is
 * no separate built-in path: supply the WAV bytes via `data` or `path`.
 * See `docs/animations-and-transitions.md`.
 */
export interface TransitionSoundProps {
	/** Embedded sound bytes as a base64 data URI (e.g. `data:audio/wav;base64,…`) or raw base64. */
	data?: string
	/** Path to a sound file, read at export time (alternative to `data`). */
	path?: string
	/** Display name emitted on `<p:snd @name>` (e.g. `ding.wav`); defaults to the file name. */
	name?: string
	/** Loop the sound until the next sound starts (`<p:stSnd loop="1">`). @default false */
	loop?: boolean
	/** Emit the stop-previous form (`<p:endSnd/>`) instead of a start sound. Mutually exclusive with `data`/`path`. */
	stopPrevious?: boolean
}

/**
 * A preset build-animation effect. The supported set is fixed (each is a verbatim
 * template captured from PowerPoint); adding one means adding a fixture + template,
 * not a new code path. See `docs/animations-and-transitions.md`.
 */
export type PresetEffect = 'fadeIn' | 'flyIn' | 'appear' | 'wipe' | 'grow' | 'spin' | 'fadeOut' | 'flyOut'

/** When a build-animation effect starts relative to the preceding one. */
export type AnimationTrigger = 'onClick' | 'withPrevious' | 'afterPrevious'

/**
 * A preset build animation on a shape (entrance/emphasis/exit), added via
 * {@link Slide.addAnimation}. Target the shape by its 0-based add order
 * (`shapeIndex`, mapping to the generated `spid = shapeIndex + 2`) or by
 * `objectName`. Effects play in the order added, grouped into click steps by
 * `trigger`.
 */
export interface AnimationProps {
	/** The preset effect to play. */
	preset: PresetEffect
	/**
	 * 0-based add order of the target shape on the slide (`spid = shapeIndex + 2`).
	 * Counts top-level objects only; use `objectName` to target a shape inside a group.
	 */
	shapeIndex?: number
	/**
	 * Target shape by its `objectName` (alternative to `shapeIndex`). Any shape on the slide,
	 * including one inside a group; an unresolved name warns and drops the effect.
	 */
	objectName?: string
	/** Trigger relative to the previous effect. @default 'onClick' */
	trigger?: AnimationTrigger
	/** Effect duration in milliseconds; preset-specific default when omitted. */
	durationMs?: number
}
