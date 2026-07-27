/**
 * A slide's show transition (`p:transition`) → the write API's `TransitionProps`.
 *
 * **The whole difficulty is one type mismatch.** `TransitionInfo.type` is an *open string*,
 * because the read model decodes PowerPoint's modern effects (Morph, Vortex, Ripple, …)
 * alongside the base ECMA-376 ones and tells them apart by namespace rather than by name.
 * The write API's `TransitionType` is a closed union of 21 base names. Passing the read
 * model's string straight through would therefore produce a script that does not compile
 * for exactly the decks a converter is most likely to meet — a modern transition is what
 * PowerPoint's own UI offers first. So {@link WRITABLE_TYPES} is the filter, and a name that
 * does not survive it becomes a fidelity note rather than a compile error.
 *
 * **Why `speed` is always emitted.** `spd` is absent from most PowerPoint-authored
 * transitions (its schema default is `fast`) and the read model reports the default rather
 * than the absence, so there is nothing to distinguish "absent" from "explicitly fast". That
 * is harmless on its own — both read back as `fast`. What is *not* harmless is omitting it
 * when a duration is present: the write path then derives the bucket from `durationMs`
 * (`transitionSpeedForDuration`), and PowerPoint's own pairing does not always agree with
 * that formula — the sound fixture carries `spd="slow"` with `p14:dur="2000"`, which the
 * derivation happens to reproduce, while a 600ms `fast` transition would come back `med`.
 * Emitting the read value verbatim removes the derivation from the path entirely, which is
 * the same discipline the rest of the IR follows: state what was read, never recompute it.
 */
import type { Slide } from '../../read/api/slide.js'
import type { TransitionInfo } from '../../read/api/transition.js'
import type { TransitionType } from '../../types/animation.js'
import type { NoteScope } from '../fidelity.js'
import type { TransitionIr, TransitionSoundIr } from '../ir.js'
import type { AssetResolver } from './shape.js'

/**
 * The transitions the write API can author, as a lookup.
 *
 * Spelled as a `Record` keyed by `TransitionType` rather than as an array of strings so the
 * compiler polices it in *both* directions: a name the union does not have is rejected here,
 * and a name the union gains without being added here fails to satisfy the record. A plain
 * `Set<string>` would silently drift out of date, and the symptom would be a printed script
 * that does not compile — found by whoever runs it, not by this repository's test suite.
 */
const WRITABLE_TYPES: Record<TransitionType, true> = {
	blinds: true,
	checker: true,
	circle: true,
	comb: true,
	cover: true,
	cut: true,
	diamond: true,
	dissolve: true,
	fade: true,
	newsflash: true,
	plus: true,
	pull: true,
	push: true,
	random: true,
	randomBar: true,
	split: true,
	strips: true,
	wedge: true,
	wheel: true,
	wipe: true,
	zoom: true,
}

/**
 * `true` when the write API has a name for this transition.
 *
 * The namespace check does no work on PowerPoint's own output and is kept anyway. Every one
 * of the 21 modern effect names in the probed table (`conveyor`, `morph`, `vortex`, …) is
 * distinct from every base name, so today the name lookup alone would reach the same verdict
 * — measured, not assumed. But "is this a base ECMA-376 transition" is a question about the
 * *namespace*, and answering it by name would be relying on that disjointness holding for
 * every namespace Microsoft ever adds. The write path emits `p:${type}` unconditionally, so
 * the cost of being wrong is a modern effect silently rewritten as a base one with the same
 * name and no note — a conversion that reports itself as faithful. `script-ir.test.js`
 * authors a `p14:fade` to keep this branch honest, since no fixture can.
 */
function isWritable(info: TransitionInfo): boolean {
	return info.namespace === 'p' && Object.hasOwn(WRITABLE_TYPES, info.type)
}

/**
 * Map a slide's transition, recording a note for anything that does not survive.
 *
 * Returns `undefined` both when the slide has no transition and when it has one this
 * library cannot author — the difference between those two is what the note carries.
 */
export function transitionToIr(slide: Slide, notes: NoteScope, assets: AssetResolver): TransitionIr | undefined {
	const info = slide.transition
	if (!info) return undefined

	if (!isWritable(info)) {
		notes.note(
			'slide.transition',
			'dropped',
			'unwritable',
			`this slide's ${info.namespace === 'p' ? '' : `${info.namespace} `}${info.type} transition is not one of the 21 base ECMA-376 transitions the write API can author${
				info.namespace === 'p' ? '' : " — it is one of PowerPoint's modern effects, which have no write-API vocabulary"
			}, so the regenerated slide advances with no effect`
		)
		return undefined
	}

	// Each optional key is spread in only when the source carried it, so an absent attribute
	// stays absent rather than becoming an explicit default the source never wrote.
	const sound = soundToIr(slide, info, notes, assets)
	return {
		type: info.type,
		speed: info.speed,
		...(info.durationMs === null ? {} : { durationMs: info.durationMs }),
		// Both are omitted at their OOXML default, so the emitted script says only what the
		// source said: `advClick` defaults to true and `advTm` to unset.
		...(info.advanceOnClick ? {} : { advanceOnClick: false }),
		...(info.advanceAfterMs === null ? {} : { advanceAfterMs: info.advanceAfterMs }),
		...(Object.keys(info.variant).length > 0 ? { variant: { ...info.variant } } : {}),
		...(sound ? { sound } : {}),
	}
}

/**
 * The transition's sound (`p:sndAc`), in the two forms OOXML gives it.
 *
 * The stop-previous form is pure structure — no relationship, no media part — so it maps
 * unconditionally. A start sound is an embedded WAV reached through the slide's own
 * relationships, which is the only place in this converter that resolves an `r:embed` by
 * hand: every other media reference arrives already resolved to a partname by the read
 * model's shape accessors, and a transition sound has no shape to hang off.
 */
function soundToIr(
	slide: Slide,
	info: TransitionInfo,
	notes: NoteScope,
	assets: AssetResolver
): TransitionSoundIr | undefined {
	const sound = info.sound
	if (!sound) return undefined
	if (sound.form === 'stop') return { stopPrevious: true }

	const partName = soundPartName(slide, sound.embedRid)
	const asset = partName === null ? null : assets.assetFor(partName)
	if (!asset) {
		notes.note(
			'slide.transitionSound',
			'dropped',
			'unsupported',
			`this transition's start sound (${sound.embedRid ?? 'no r:embed'}) does not resolve to an audio part in the package, so the regenerated transition is silent`
		)
		return undefined
	}

	return {
		data: asset,
		...(sound.name === null ? {} : { name: sound.name }),
		// Omitted at its default, like `advanceOnClick`: `p:stSnd@loop` defaults to false.
		...(sound.loop ? { loop: true } : {}),
	}
}

/** Resolve `p:snd@r:embed` to the audio part it names, or `null` when it does not resolve. */
function soundPartName(slide: Slide, embedRid: string | null): string | null {
	if (!embedRid) return null
	const relationship = slide.relationships.get(embedRid)
	// An External audio rel has no partname to read bytes from; `resolveTarget` throws on
	// both cases, so they are screened here rather than caught.
	if (!relationship || relationship.targetMode === 'External') return null
	return slide.relationships.resolveTarget(embedRid)
}
