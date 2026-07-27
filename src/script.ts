/**
 * `ts-pptx/script` — turn an existing `.pptx` into a description of the write-API calls
 * that would rebuild it.
 *
 * **Why this is its own subsystem.** It depends on both halves of the library: it reads a
 * deck through `ts-pptx/read` and targets the write API's option types. That rules out
 * living inside either. `src/read/` in particular is documented as isomorphic — bytes in,
 * bytes out, no `node:fs` — and a converter whose output is *source text* rather than a
 * package would quietly break that guarantee for everyone importing `ts-pptx/read`.
 *
 * **What it is not.** This is not a lossless round-trip. Some of a deck cannot be
 * expressed through the public write API, and some of it cannot even be seen through the
 * public read API — measurement against the fixture corpus found the read side is the
 * tighter of the two. Rather than warn about that to a log, every loss is a
 * {@link FidelityNote} attached to the IR, which makes the losses testable: a round-trip
 * check excludes exactly the noted fields and treats any other difference as a defect.
 * Read {@link DeckIr.fidelity} before trusting the output of a conversion.
 */
export { readModelToIr } from './script/from-read/deck.js'
export { printScript } from './script/print/script.js'
export type { PrintScriptOptions } from './script/print/script.js'
export { printStandaloneScript } from './script/print/standalone.js'
export type { PrintStandaloneScriptOptions } from './script/print/standalone.js'
export type { AssetMode, PrintedScript } from './script/print/common.js'
export { canonicalDeckIr } from './script/verify/canonical.js'
export type { CanonicalCall, CanonicalChrome, CanonicalDeck, CanonicalSlide } from './script/verify/canonical.js'
export { diffDeckIr, knownNoteConstructs } from './script/verify/diff.js'
export type { DifferenceKind, IrDifference, RoundTripReport } from './script/verify/diff.js'
export { isAssetRef } from './script/ir.js'
export type {
	AssetIr,
	AssetRef,
	BackgroundIr,
	CallIr,
	ChromeIr,
	DeckIr,
	DeckPropsIr,
	IrValue,
	MasterIr,
	SlideIr,
	SlideLayoutIr,
	SlideSource,
	ThemeIr,
	TransitionIr,
	TransitionSoundIr,
} from './script/ir.js'
export type { Cause, Disposition, FidelityNote } from './script/fidelity.js'
