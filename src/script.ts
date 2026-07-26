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
export { isAssetRef } from './script/ir.js'
export type {
	AssetIr,
	AssetRef,
	BackgroundIr,
	CallIr,
	DeckIr,
	DeckPropsIr,
	IrValue,
	SlideIr,
	SlideSource,
} from './script/ir.js'
export type { Cause, Disposition, FidelityNote } from './script/fidelity.js'
