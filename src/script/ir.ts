/**
 * The deck IR — a serializable description of a deck as a sequence of write-API calls.
 *
 * This is the seam between reading a `.pptx` and printing TypeScript for it. Everything
 * upstream of the IR knows about OOXML and the read model; everything downstream knows
 * only about strings. Neither half can see the other, which is the point: the decision
 * "this `a:ln` becomes `line: { width: 2 }`" is made once, here, and is testable without
 * a printer, while "how a number is spelled in source text" is a printer concern that
 * cannot silently change what the deck means.
 *
 * Two rules keep that seam honest, and both are load-bearing:
 *
 * 1. **The IR is data, not objects.** No DOM nodes, no read-model instances, no
 *    functions — `structuredClone` and `JSON.stringify` both round-trip it. A round-trip
 *    test can therefore compare two IRs directly (see {@link FidelityNote}), and an IR
 *    can be cached or shipped without dragging a whole `Presentation` along.
 * 2. **`args` are literal write-API option objects.** A {@link CallIr} is exactly what
 *    the emitted source will pass, so the printer never invents, reshapes, or defaults a
 *    value — it spells out what is already there. If the printer had to interpret, the
 *    same construct could mean two things depending on which half you asked.
 *
 * Media is the one thing that cannot be a literal, since bytes are not source text. It
 * is referenced by {@link AssetRef} and carried out-of-band in {@link DeckIr.assets}, so
 * the IR itself stays serializable and the printer decides between a file path and an
 * inline base64 literal.
 */
import type { FidelityNote } from './fidelity.js'

/**
 * A reference to bytes held in {@link DeckIr.assets}, standing where a write-API option
 * wants image/media data. Printed as a `path` or an inline `data:` literal depending on
 * the print options — an IR-level decision would bake one choice into the IR and force a
 * re-read to change it.
 *
 * The `$asset` sigil is deliberate: no write-API option object has a `$`-prefixed key,
 * so a value carrying one is unambiguously a reference rather than user data.
 */
export interface AssetRef {
	$asset: string
}

/** `true` when `value` is an {@link AssetRef} rather than a plain IR object. */
export function isAssetRef(value: unknown): value is AssetRef {
	return typeof value === 'object' && value !== null && typeof (value as AssetRef).$asset === 'string'
}

/**
 * Any value an IR may hold. Deliberately narrower than `unknown`: it excludes `undefined`
 * so that "absent" has exactly one spelling (a missing key), which is what makes two IRs
 * comparable field-by-field in a round-trip diff. `readModelToIr` never writes an
 * `undefined` value.
 */
export type IrValue = null | boolean | number | string | AssetRef | IrValue[] | { [key: string]: IrValue }

/** Bytes referenced by an {@link AssetRef}, plus what a printer needs to name and type them. */
export interface AssetIr {
	/** Stable key an {@link AssetRef} resolves against; also the emitted filename. */
	name: string
	/** MIME type from the source package's content types, for an inline `data:` URI. */
	contentType: string
	bytes: Uint8Array
}

/**
 * One write-API call on a slide, in the order it must be made. `method` names a real
 * method on the authored slide and `args` are its positional arguments — so a printer is
 * a formatter, and `readModelToIr` owns every semantic choice.
 *
 * Ordering is z-order: OOXML paints a slide's shape tree front-to-back in document order,
 * and so does the write path, so preserving call order preserves overlap.
 */
export interface CallIr {
	method: 'addText' | 'addShape' | 'addImage' | 'addTable' | 'addChart' | 'addConnector' | 'addGroup' | 'addNotes'
	args: IrValue[]
	/**
	 * The source shape's name (`p:cNvPr/@name`), when it had one. Not an argument —
	 * carried so a printer can comment the call and a fidelity note can point at it.
	 */
	sourceName?: string
}

/**
 * How a slide's content gets into the output deck.
 *
 * `authored` — every shape is transcribed into {@link CallIr}s and regenerated through
 * the write API. This is the goal: the emitted source is editable, and a reader can see
 * what the deck contains.
 *
 * `carried` — the slide holds at least one construct the write API cannot express at all
 * (`chartEx` is the clear case: a full reader, no emitter), so transcribing it would
 * silently produce a different deck. The slide is copied from the source package instead.
 * Fidelity is *higher* than transcription here, but the emitted source says nothing about
 * what is on the slide, so this is a real cost and always carries a {@link FidelityNote}.
 */
export type SlideSource = 'authored' | 'carried'

/** A slide's background, already reduced to what the write API's `background` accepts. */
export interface BackgroundIr {
	color?: string
	transparency?: number
	image?: AssetRef
}

/**
 * The layout a slide should be bound to in the destination deck, identified two ways
 * because neither alone is sufficient.
 *
 * The **name** is the portable identity: it is what lets a script built from one deck be
 * re-pointed at a different template whose layouts happen to be named the same, and it is
 * what a reader recognises. But `p:cSld@name` is not unique — a deck with several masters
 * routinely carries several "Title and Content" layouts, and binding by an ambiguous name
 * is an error, not a coin flip. The **index** into the deck's layout gallery always
 * resolves, so it is the fallback, and {@link nameIsUnique} says which to trust.
 */
export interface SlideLayoutIr {
	/** `p:cSld@name` of the source layout (`''` when the layout is unnamed). */
	name: string
	/** Zero-based position in the source deck's layout gallery, in master-then-layout order. */
	index: number
	/** `false` when another layout in that gallery shares {@link name}. */
	nameIsUnique: boolean
}

/** One slide of the deck. */
export interface SlideIr {
	/** 1-based index in the source deck; the identity a {@link FidelityNote} points at. */
	number: number
	source: SlideSource
	/** The layout to bind to, or `null` when the source slide resolves none. */
	layout: SlideLayoutIr | null
	/** `p:sld/@show="0"` — a slide hidden from presentation but present in the deck. */
	hidden: boolean
	/** `p:cSld/@name`, when the source slide had one. */
	name?: string
	background?: BackgroundIr
	/** Speaker notes as plain text. Notes-slide geometry and placeholders do not survive. */
	notesText?: string
	/** Write-API calls in z-order. Empty when {@link source} is `carried`. */
	calls: CallIr[]
}

/** Deck-level properties, reduced to the five `docProps` fields the write API sets. */
export interface DeckPropsIr {
	title?: string
	author?: string
	company?: string
	subject?: string
	revision?: string
}

/**
 * A whole deck, ready to print.
 *
 * `widthEmu`/`heightEmu` are raw EMU rather than inches because that is what the read
 * model reports and what the write API's `Coord` accepts verbatim; converting here would
 * introduce a rounding decision at the wrong layer.
 */
export interface DeckIr {
	slideSize: { widthEmu: number; heightEmu: number }
	props: DeckPropsIr
	slides: SlideIr[]
	assets: AssetIr[]
	/**
	 * Every construct that did not survive into {@link slides} intact. This is a
	 * *contract*, not a log: a round-trip check excludes exactly these fields from its
	 * diff, so an unnoted difference is a bug and a note nothing else corroborates is a
	 * stale claim. See {@link FidelityNote}.
	 */
	fidelity: FidelityNote[]
}
