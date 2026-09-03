/**
 * Fidelity notes — the declared, machine-checkable list of what a deck loses on the way
 * to a script.
 *
 * A converter that warns about losses to a log has no way to be tested: the warnings are
 * prose, nothing consumes them, and a missing one looks exactly like a clean run. Making
 * the notes part of the IR inverts that. A note says "field X of shape Y will not come
 * back", so a round-trip check can exclude precisely those fields and treat **any**
 * remaining difference as a defect. Both directions then have teeth:
 *
 * - an *undeclared* loss fails the round-trip, because nothing excluded it;
 * - a *declared* loss that actually survives is a stale note, and the check can say so.
 *
 * That is the whole reason notes are first-class rather than a side channel. It also
 * means {@link FidelityNote.construct} is an identifier, not a sentence: it has to be
 * matchable against a field path, so it is a stable dotted key and the human explanation
 * lives in {@link FidelityNote.detail}.
 */

/**
 * Option names each note construct is a promise about.
 *
 * This table is the fidelity catalogue made mechanical. A note claims a construct will not
 * survive; without a mapping from that claim to the fields it covers, the claim cannot
 * exclude anything and the round trip degenerates into a snapshot.
 *
 * Each entry is a dotted option PATH, matched as a suffix of where the difference sits: a bare
 * `fill` means "the fill option wherever it appears", and `line.width` means that width and not
 * a table cell border's. Write the bare form where the word is unambiguous and qualify it the
 * moment the same word names two different things — matching the terminal key alone is how
 * `type`, written about a fill's solid default, came to excuse a character bullet that returned
 * as a numbered list.
 *
 * `'*'` covers every difference inside the call it is scoped to — correct only where the
 * note says the whole shape is gone or was copied wholesale, never as a shortcut for a
 * construct whose fields are merely tedious to enumerate.
 *
 * An empty list is meaningful and common: the construct is invisible to the IR on both
 * sides (a slide's build animations, a paragraph's `a:pPr/@marL`, a connector's shape
 * binding), so the note is a caveat for a human and there is nothing here to exclude. Saying
 * so explicitly is what keeps it from looking like an oversight. A construct can leave that
 * state — `slide.transition` was an empty entry until the converter learned to transcribe
 * transitions, and the IR gaining a field is exactly when its note gains a mapping.
 */
export const NOTE_CONSTRUCTS = {
	'chart.blanks': ['values'],
	'chart.combo': ['type'],
	// A chart with no cached plottable series is not emitted at all, so the whole frame is
	// missing from the output. Distinct from `chart.workbook`, which is about a chart that IS
	// emitted, rebuilt from the cache.
	'chart.data': ['*'],
	'chart.type': ['type'],
	'chart.workbook': ['*'],
	'chartEx.all': ['*'],
	'connector.binding': [],
	'connector.line': ['color', 'width', 'dashType', 'beginArrowType', 'endArrowType'],
	'connector.rotation': ['rotate'],
	// Empty, and the temptation to list title/author/subject/revision here has to be
	// resisted: this note is about the *other* seven docProps, which have no setter and are
	// absent from the IR. Listing the four that do have setters would excuse a printer that
	// stopped writing them — measured, by exactly that mutation.
	'deck.docProps': [],
	// Nothing to exclude: the differences it predicts are `added`, which WRITER_DEFAULTS covers
	// by kind. The note exists so a reader of the emitted script learns the deck gained them.
	'deck.docPropsDefault': [],
	'deck.slideSize': ['widthEmu', 'heightEmu'],
	'diagram.all': ['*'],
	// A gradient that cannot be expressed falls back to no gradient, so the difference lands on
	// the fill option itself. The `line.` twins below are the same construct on a stroke —
	// `gradientStops` is shared by both surfaces and scopes its notes by which one it is on.
	'fill.gradient': ['gradient', 'fill'],
	'fill.gradient.path': ['gradient', 'fill'],
	'line.gradient': ['gradient', 'line'],
	'line.gradient.path': ['gradient', 'line'],
	'line.gradient.schemeToken': ['gradient', 'line'],
	// Recorded only when an image-filled surface cannot carry its *bytes* — a linked blip, an
	// SVG the write path refuses, a part missing from the package. The fill option is then
	// absent from the output entirely, which is the same two keys `fill.schemeToken` covers.
	'fill.picture': ['fill', 'color'],
	// Empty, and deliberately so: this note declares that a picture fill's tiling, crop, DPI
	// and rotWithShape do not survive, and *none of them is in the IR on either side* — the
	// write API expresses a picture fill as bytes plus transparency, so the converter never
	// emits them. Widening this to `fill` would be the mistake it looks like a fix for: it
	// would excuse an image fill that failed to come back at all, which is the thing the
	// round trip is here to catch.
	'fill.picture.geometry': [],
	'fill.gradient.schemeToken': ['gradient', 'fill'],
	'fill.schemeToken': ['fill', 'color'],
	'graphicFrame.unknown': ['*'],
	'group.child': ['*'],
	'group.childSpace': ['x', 'y', 'w', 'h', 'rotate', 'flipH', 'flipV'],
	'group.empty': ['*'],
	'group.transform': ['rotate', 'flipH', 'flipV'],
	'image.data': ['data', '$asset'],
	'image.recolor': ['duotone', 'grayscale', 'biLevel', 'clrChange'],
	// Covers the picture's bytes, not just an `svg` option: an SVG picture's raster fallback
	// is regenerated rather than carried, so the blip the round trip compares is a different
	// image from the source's — which is precisely the loss this note is about.
	'image.svg': ['svg', 'data', '$asset'],
	'line.arrowSize': ['beginArrowType', 'endArrowType'],
	// The chrome notes. Most are empty for the reason stated above and it is the common case
	// here rather than the exception: `a:fmtScheme`, `p:txStyles`, a *master's* decoration and a
	// layout's placeholder definitions are all absent from the IR on *both* sides — the first two
	// because nothing reads them, the other two because nothing writes them — so there is nothing
	// for a note to exclude and the note exists for a human. A *layout's* decoration is the one
	// that left this state: it is now in the IR as `objects` and is genuinely compared, which is
	// what makes the `layout.` entries below worth stating separately.
	// `p:clrMap` is the subtle one: a remapped token changes what every scheme colour in the
	// deck resolves to, and the round trip still cannot see it, because the IR reports the token
	// verbatim rather than its resolved hex. That is exactly the blind spot this file's header
	// describes, and writing `[]` is the honest spelling of it.
	'master.background': ['background', 'color', 'data', 'master'],
	'master.colorMap': [],
	'master.decoration': [],
	'master.default': ['master'],
	'master.multiple': [],
	'master.name': ['title', 'layoutName'],
	'master.nameCollision': ['title', 'layoutName'],
	'master.placeholders': [],
	'master.txStyles': [],
	'theme.fmtScheme': [],
	// The two layout-shape notes with no slide counterpart to inherit a mapping from. Both are
	// empty for the same reason, and it is worth spelling out because both *look* like they
	// should exclude something. A table on a layout is absent from the `objects` array on both
	// sides — the source's because this converter skips it, the output's because it was never
	// written — so there is no difference to excuse. A flattened group is stronger still: the
	// source layout's group becomes N loose objects here and the output layout genuinely *has*
	// N loose objects, so the two agree exactly and the note is a caveat for a human reading the
	// emitted script, not an exclusion.
	'layout.decoration': [],
	'layout.group': [],
	// `a:ln/@algn="in"` has no write option, so the stroke comes back centred on the edge. The
	// IR carries no alignment on either side, so there is nothing to exclude and the note is a
	// caveat for a human reading the emitted script.
	'line.align': [],
	// Path-qualified: the entries are matched as a suffix of the difference's own path, and a
	// bare `width` also excused a table cell's bevel or border width.
	'line.dash': ['line.dashType'],
	'line.width': ['line.width'],
	'media.audioVideo': ['*'],
	'notes.formatting': ['notesText'],
	'shape.custGeom.guides': ['points'],
	'shape.effects': ['shadow', 'glow'],
	'shape.empty': ['*'],
	'shape.frameInherited': ['x', 'y', 'w', 'h'],
	'shape.hidden': ['*'],
	'shape.placeholder': ['placeholder'],
	'slide.animation': [],
	'slide.background': ['background', 'color', 'transparency', 'image'],
	'slide.carried': ['*'],
	'slide.layout': ['layoutName'],
	'slide.name': [],
	// A transition the write vocabulary cannot name is dropped whole, so the difference lands
	// on the slide's `transition` key itself. Deliberately *not* widened to the keys inside it:
	// this note is only ever recorded when the whole transition is gone.
	'slide.transition': ['transition'],
	// The sound alone, one level down. Scoped to `sound` so it cannot also excuse a transition
	// whose type or timing came back wrong — the loss it declares is exactly the missing
	// `p:sndAc`, and `data`/`$asset` cover the case where the sound survives with other bytes.
	'slide.transitionSound': ['sound', 'data', '$asset'],
	// A dash outside `ST_PresetLineDashVal` cannot be written back, so the edge comes out as
	// a plain dashed rule. Scoped to `border`, which is where that difference lands.
	'table.cell.borders.dash': ['border', 'diagonal'],
	// `table.cell.fill` used to live here, for a cell whose own fill could not be told apart
	// from the one it inherited from the table style. `TableCell.hasOwnFill` tells them apart,
	// so the mapper emits the right one and records nothing — the note is gone rather than
	// unmapped. Its `.gradient` / `.picture` children below are separate constructs and stay.
	// Empty, and correctly so: `a:tc/@id` and `a:tcPr/a:headers` have no write option — the
	// IR has nowhere to put them on either side, so there is nothing to exclude. The note
	// exists so a reader of the emitted script learns the association was there and is gone.
	'table.cell.headers': [],
	// The cell-side twins of `fill.picture` / `fill.picture.geometry`, and mapped for the
	// same reasons.
	'table.cell.fill.picture': ['fill'],
	'table.cell.fill.picture.geometry': [],
	// The table-side twins of `fill.schemeToken` / `text.color.schemeToken`: one of the seven
	// `ST_SchemeColorVal` tokens the write path's `clrMap` does not carry, baked to the literal
	// it resolves to. Three sites used to pass one through RAW, so the generated script warned
	// `color/invalid-value` and painted the default text colour instead.
	'table.cell.fill.schemeToken': ['fill', 'color'],
	'table.cell.borders.schemeToken': ['border', 'diagonal'],
	// The table-background twins. Scoped to `tableFill` rather than `fill`, because those are
	// two different options: one lands on `a:tblPr`, the other is stamped onto every cell.
	'table.fill.schemeToken': ['tableFill', 'color'],
	'table.fill.picture': ['tableFill'],
	'table.fill.picture.geometry': [],
	// A gradient that cannot be expressed falls back to no gradient, so the difference lands
	// on the fill option itself — `tableFill` for the background, `fill` for a cell.
	'table.fill.gradient': ['tableFill'],
	'table.fill.gradient.path': ['tableFill'],
	// A gradient STOP whose scheme token the write path cannot carry, baked to a literal. Scoped
	// to the surface the stop sits on, like the two above it: the stop note used to hardcode the
	// shape spelling, so a table gradient recorded `fill.gradient.schemeToken` while its
	// difference landed on `tableFill` and the note could never match it.
	'table.fill.gradient.schemeToken': ['tableFill'],
	'table.cell.fill.gradient': ['fill'],
	'table.cell.fill.gradient.path': ['fill'],
	'table.cell.fill.gradient.schemeToken': ['fill'],
	// Narrowed to the East-Asian `ST_TextVerticalType` modes `textDirection` cannot spell —
	// the four it can now round-trip, so this no longer excuses every vertical cell.
	'table.cell.vert': ['textDirection', 'vert'],
	'table.rowAuto': ['rowH'],
	'table.style': ['tableStyle'],
	'text.align': ['align'],
	// `text.bullet.numberStartAt` and `text.bullet.style` used to live here, for
	// `a:buAutoNum/@startAt` and for a bullet's own font/size/colour. `Paragraph.bulletDetail`
	// reads all four, so the mapper emits `numberStartAt` / `fontFace` / `size` / `color` and
	// records nothing — the notes are gone rather than unmapped. What remains of the size half
	// is `text.bullet.sizePt`, which is a genuinely unwritable unit rather than an unread value.
	'text.bullet.numberType': ['bullet'],
	// An absolute bullet size (`a:buSzPts`) has no write option at all — `bullet.size` is a
	// percentage of the run size — so the difference lands on the bullet option.
	'text.bullet.sizePt': ['bullet'],
	// A percentage outside 25–400%, which the write path rejects with a warning and replaces
	// with the run's own size.
	'text.bullet.sizePct': ['bullet'],
	// A bullet colour outside the ten scheme tokens the write path maps, baked to a literal.
	'text.bullet.schemeToken': ['bullet'],
	// A picture bullet (`a:buBlip`): readable, and `bullet.image` could author it, but the
	// paragraph mapper carries no asset resolver to re-embed the bytes with.
	'text.bullet.picture': ['bullet'],
	'text.color.default': ['color'],
	'text.color.inherited': ['color'],
	'text.color.schemeToken': ['color'],
	'text.equation': ['*'],
	'text.field': ['*'],
	'text.bullet.glyph': ['bullet'],
	// `text.bullet.inherited` used to live here — a paragraph stating no bullet of its own was
	// re-emitted with an explicit `a:buNone`, because omitting the write API's `bullet` is that
	// element rather than silence. `bullet: 'inherit'` says silence now, so the mapper carries
	// the state instead of excusing its loss and the note is gone rather than unmapped.
	//
	// `text.indent` used to live here, empty, because a paragraph's `a:pPr/@marL` and `@indent`
	// were in neither IR — the write API had no option for them, so the mapper never emitted one
	// and the diff compared two models both missing the field. `paraMarginLeft` / `paraIndent`
	// put it in both, which is what turns an empty entry into a comparison rather than a caveat,
	// so the note is gone rather than unmapped.
	// A baked `a:normAutofit` percentage outside 0-100, which the write API rejects: the frame
	// re-emits a bare `<a:normAutofit/>`, so the difference lands on the `fit` option.
	'text.autofit.fontScale': ['fit'],
	'text.autofit.lnSpcReduction': ['fit'],
	'text.paraSpaceZero': ['paraSpaceBefore', 'paraSpaceAfter'],
	'text.vert': ['vert'],
} as const satisfies Record<string, readonly string[]>

/**
 * Every construct a note may name.
 *
 * Derived from the table above rather than declared beside it, so a construct coined at a call
 * site and never mapped is a COMPILE error instead of a note that silently declares nothing --
 * which is what it was: seven constructs were emitted with no entry at all, and `declaringNote`
 * walked straight past each of them, so the difference the note predicted came back as a defect.
 */
export type NoteConstruct = keyof typeof NOTE_CONSTRUCTS

/** A construct with {@link LAYOUT_NOTE_PREFIX} stripped, for the ones declared only under it. */
type StripLayoutPrefix<T> = T extends `layout.${infer Rest}` ? Rest : never

/**
 * What a call site may name: the catalogue, plus the `layout.`-prefixed constructs spelled
 * relative to the prefix, which is how the layout walk records them.
 *
 * One union rather than two scope types, because a TypeScript method parameter is bivariant —
 * a narrower `LayoutNoteScope` would have been assignable to `NoteScope` in both directions and
 * enforced nothing. The cost is that a slide mapper could name a layout-relative construct; the
 * thing worth catching is a construct with no entry AT ALL, which this does catch.
 */
export type RecordableConstruct = NoteConstruct | StripLayoutPrefix<NoteConstruct>

/** What happened to the construct. */
export type Disposition =
	/** Gone: nothing in the output carries it. */
	| 'dropped'
	/**
	 * Present but structurally simplified — a value survives while the structure around
	 * it does not (threaded comment replies flattened into separate comments, a per-script
	 * font scheme reduced to latin, a placeholder demoted to a plain shape).
	 */
	| 'flattened'
	/**
	 * Present and structurally intact, but not the same value — a re-encoding that a
	 * viewer may or may not be able to tell apart (a modern transition mapped to the
	 * nearest named one, a chart rebuilt from cached points rather than its workbook).
	 */
	| 'approximated'

/**
 * Which side of the library is responsible. Recording this is what makes the notes
 * actionable rather than merely honest: `unread` and `unwritable` are bugs waiting to be
 * fixed in a specific subsystem, while `unsupported` is a property of OOXML or of the
 * chosen output tier and will not be fixed by more converter work.
 */
export type Cause =
	/** The read API exposes no accessor for it, so the converter never saw it. */
	| 'unread'
	/** Read fine, but the write API has no option that expresses it. */
	| 'unwritable'
	/** Both sides can handle it; this output tier structurally cannot carry it. */
	| 'unsupported'

/** One declared loss. */
export interface FidelityNote {
	/** 1-based source slide, or `null` for a deck-level loss. */
	slideNumber: number | null
	/** Source shape name (`p:cNvPr/@name`), or `null` when the loss is not shape-scoped. */
	shapeName: string | null
	/**
	 * Stable dotted identifier for the lost construct — `line.width`, `text.tabStops`,
	 * `chart.workbook`. Matched mechanically by the round-trip check, so it must name a
	 * field path rather than describe one. Reuse an existing key for the same construct
	 * instead of coining a synonym.
	 *
	 * A `string` rather than {@link NoteConstruct} because a note carries the `layout.`-prefixed
	 * spelling too; the recording end is the typed one.
	 */
	construct: string
	disposition: Disposition
	cause: Cause
	/** Why, in a sentence, for a human reading the emitted script's header. */
	detail: string
}

/**
 * Accumulates notes while a deck is walked, so the walk itself stays a plain
 * read-model → IR mapping and does not thread a growing array through every function.
 *
 * Deduplicates on the full note identity. Without that, a construct that is lost on every
 * run of a 400-run text body would produce 400 identical notes and bury the one-off losses
 * that actually need reading.
 */
export class NoteCollector {
	readonly #seen = new Set<string>()
	readonly #notes: FidelityNote[] = []

	/** Record a loss. A repeat of an identical note is ignored. */
	add(note: FidelityNote): void {
		const key = `${note.slideNumber}\0${note.shapeName}\0${note.construct}\0${note.disposition}\0${note.cause}\0${note.detail}`
		if (this.#seen.has(key)) return
		this.#seen.add(key)
		this.#notes.push(note)
	}

	/**
	 * The notes, in the order first recorded — which is deck order, since the walk is
	 * ordered, so a reader meets them the way they meet the slides.
	 */
	get notes(): FidelityNote[] {
		return this.#notes
	}
}

/**
 * Namespace for a loss recorded against a shape on a **slide layout** rather than on a slide,
 * prepended to the construct by {@link layoutShapeScope}.
 *
 * The chrome mapper reuses the slide shape mapper wholesale — a rectangle on a layout is
 * transcribed by the same code that transcribes one on a slide, which is what keeps the two
 * from drifting. That sharing brings the slide vocabulary with it: a themed outline on a
 * layout records `line.width` exactly as a slide's would. Left unmarked, those notes would be
 * wrong twice over. A template-anchored script would report them, though it never rebuilds a
 * layout and loses nothing; and the round trip would let one excuse a *slide* difference,
 * because a layout note carries no slide number and a name like "Rectangle 7" repeats between
 * the tiers. The prefix keeps the shared vocabulary and states which tier it is about.
 *
 * `layout.` rather than `master.`, which is how the other chrome constructs are spelled: those
 * are named after the `defineSlideMaster` call that authors them, while these name the source
 * tier the loss is on — and a master's own decoration is not re-authored at all, so a note
 * under this prefix is always about a layout.
 */
export const LAYOUT_NOTE_PREFIX = 'layout.'

/**
 * A {@link NoteCollector} bound to one slide and shape, so a mapping function can record
 * a loss without knowing where in the deck it sits. Obtained via {@link scopeNotes}.
 */
export interface NoteScope {
	/** Record a loss at this scope. */
	note(construct: RecordableConstruct, disposition: Disposition, cause: Cause, detail: string): void
	/** Re-scope to a shape within the same slide. */
	forShape(shapeName: string | null): NoteScope
}

/** Bind a collector to a slide (and optionally a shape). */
export function scopeNotes(
	collector: NoteCollector,
	slideNumber: number | null,
	shapeName: string | null = null
): NoteScope {
	return {
		note(construct, disposition, cause, detail) {
			collector.add({ slideNumber, shapeName, construct, disposition, cause, detail })
		},
		forShape(name) {
			return scopeNotes(collector, slideNumber, name)
		},
	}
}

/**
 * Wrap a scope so every construct recorded through it is namespaced under
 * {@link LAYOUT_NOTE_PREFIX} — see there for why.
 *
 * Constructs written at a call site inside the layout walk are therefore *relative*: the
 * decoration mapper records `decoration`, and what lands in the note list is
 * `layout.decoration`. The wrapper survives `forShape`, since a layout loss is scoped to a
 * shape name the same way a slide loss is.
 */
export function layoutShapeScope(notes: NoteScope): NoteScope {
	return {
		note(construct, disposition, cause, detail) {
			// The prefixed spelling is a `NoteConstruct` for the constructs declared under it and a
			// borrowed slide one otherwise, which `noteFields` resolves by stripping the prefix.
			notes.note(`${LAYOUT_NOTE_PREFIX}${construct}` as RecordableConstruct, disposition, cause, detail)
		},
		forShape(name) {
			return layoutShapeScope(notes.forShape(name))
		},
	}
}
