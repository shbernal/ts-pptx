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

/** The two printers a note can apply to: template-anchored (`printScript`) and standalone (`printStandaloneScript`). */
export type PrintTier = 'template' | 'standalone'

/** One catalogue entry: the option paths a note excuses, and the printers whose output it describes. */
export interface NoteConstructEntry {
	/** Option paths the note is a promise about; see {@link NOTE_CONSTRUCTS}. */
	readonly fields: readonly string[]
	/** The printers that report the note. */
	readonly tiers: readonly PrintTier[]
}

/** A loss both printers have. */
const BOTH_TIERS = ['template', 'standalone'] as const

/**
 * A loss only the standalone printer has, because the template-anchored one keeps the construct in
 * its template.
 *
 * Document properties are the plain case: the read half notes that only five of the twelve have
 * write-API setters, which matters for a script that rebuilds the deck, but the template-anchored
 * tier never authors them. They ride in the template, all twelve. The chrome is the same story on
 * a larger scale: the theme's format scheme, the master's text styles and colour map, the layouts'
 * decoration, names and placeholder definitions are genuine losses for a script that rebuilds the
 * design, and untouched by one whose template *is* the design. A caveat that does not apply to the
 * output in front of you is worse than none: it teaches the reader to skim the ones that do.
 */
const STANDALONE_ONLY = ['standalone'] as const

/**
 * A loss only the template-anchored printer has. `slide.carried` says a slide is copied from the
 * source package rather than transcribed. The standalone tier has no source package, so it
 * transcribes the slide, and the construct that made it uncarryable records its own note; reporting
 * both would count one loss twice and describe a behaviour that script does not have.
 */
const TEMPLATE_ONLY = ['template'] as const

/**
 * Option names each note construct is a promise about.
 *
 * This table is the fidelity catalogue made mechanical. A note claims a construct will not
 * survive; without a mapping from that claim to the fields it covers, the claim cannot
 * exclude anything and the round trip degenerates into a snapshot.
 *
 * Each entry is a dotted option PATH, matched one of two ways depending on the note that names it:
 *
 * - **A note scoped to a shape** is already confined to that shape's call, so its entries are
 *   matched as a SUFFIX of where the difference sits inside it: a bare `fill` means "the fill
 *   option wherever it appears in the call", and `line.width` means that width and not a table
 *   cell border's. Write the bare form where the word is unambiguous and qualify it the moment
 *   the same word names two different things. Matching the terminal key alone is how `type`,
 *   written about a fill's solid default, came to excuse a character bullet that returned as a
 *   numbered list, and how a bare `color` in a run-colour note excused the shape's fill and
 *   outline colours, and in a cell note every colour in the table.
 * - **A note with no shape** (a slide or deck loss) has nothing confining it to one call, so its
 *   entries are anchored at the root of the diff. Each names one node by its full path
 *   (`background`, `chrome.masters.title`) and covers that node only, and a trailing `.*` extends
 *   it to everything beneath (`transition.sound.*`). As suffixes they matched inside every call in
 *   scope, and a deck's `master.background` excused every run colour and every lost image in the
 *   deck.
 *
 * `'*'` covers every difference inside the note's scope — correct only where the note says the
 * whole shape or slide is gone or was copied wholesale, never as a shortcut for a construct
 * whose fields are merely tedious to enumerate.
 *
 * An empty list is meaningful and common: the construct is invisible to the IR on both
 * sides (a slide's build animations, a paragraph's `a:pPr/@marL`, a connector's shape
 * binding), so the note is a caveat for a human and there is nothing here to exclude. Saying
 * so explicitly is what keeps it from looking like an oversight. A construct can leave that
 * state — `slide.transition` was an empty entry until the converter learned to transcribe
 * transitions, and the IR gaining a field is exactly when its note gains a mapping.
 *
 * Each entry also names the printers whose output the note describes (`tiers`), and each printer
 * reports only the notes that apply to it. Two hand-kept sets used to decide that, one per printer,
 * so a new chrome construct was reported by the template-anchored tier unless someone remembered
 * to list it. An entry without `tiers` no longer satisfies the catalogue's type.
 */
export const NOTE_CONSTRUCTS = {
	'chart.blanks': { fields: ['values'], tiers: BOTH_TIERS },
	'chart.combo': { fields: ['type'], tiers: BOTH_TIERS },
	// A chart with no cached plottable series is not emitted at all, so the whole frame is
	// missing from the output. Distinct from `chart.workbook`, which is about a chart that IS
	// emitted, rebuilt from the cache.
	'chart.data': { fields: ['*'], tiers: BOTH_TIERS },
	// Label flags the writer has no spelling for: the legend key on any plot, the category name on any
	// but a pie or a doughnut. Absent from the IR on both sides, so there is nothing to exclude.
	'chart.labels': { fields: [], tiers: BOTH_TIERS },
	'chart.type': { fields: ['type'], tiers: BOTH_TIERS },
	// A 3-D line, area or pie rebuilt flat: `type` names the flat chart on both sides, and the 3-D
	// view is in neither IR.
	'chart.type3D': { fields: [], tiers: BOTH_TIERS },
	// Scatter or bubble series with X values of their own, rebuilt against the first series' X row.
	// The IR carries only that row on both sides, so the per-series X values are in neither.
	'chart.xValues': { fields: [], tiers: BOTH_TIERS },
	// Scatter or bubble series plotted against text X labels, rebuilt against the 1..n positions
	// PowerPoint plots them at. Both IRs carry those positions, and the labels are in neither.
	'chart.xLabels': { fields: [], tiers: BOTH_TIERS },
	'chart.workbook': { fields: ['*'], tiers: BOTH_TIERS },
	'chartEx.all': { fields: ['*'], tiers: BOTH_TIERS },
	'connector.binding': { fields: [], tiers: BOTH_TIERS },
	'connector.line': { fields: ['color', 'width', 'dashType', 'beginArrowType', 'endArrowType'], tiers: BOTH_TIERS },
	'connector.rotation': { fields: ['rotate'], tiers: BOTH_TIERS },
	// Empty, and the temptation to list title/author/subject/revision here has to be
	// resisted: this note is about the *other* seven docProps, which have no setter and are
	// absent from the IR. Listing the four that do have setters would excuse a printer that
	// stopped writing them — measured, by exactly that mutation.
	'deck.docProps': { fields: [], tiers: STANDALONE_ONLY },
	// Nothing to exclude: the differences it predicts are `added`, which WRITER_DEFAULTS covers
	// by kind. The note exists so a reader of the emitted script learns the deck gained them.
	'deck.docPropsDefault': { fields: [], tiers: STANDALONE_ONLY },
	'deck.slideSize': { fields: ['slideSize.*'], tiers: BOTH_TIERS },
	'diagram.all': { fields: ['*'], tiers: BOTH_TIERS },
	// A gradient that cannot be expressed falls back to no gradient, so the difference lands on
	// the fill option itself. The `line.` twins below are the same construct on a stroke —
	// `gradientStops` is shared by both surfaces and scopes its notes by which one it is on.
	'fill.gradient': { fields: ['gradient', 'fill'], tiers: BOTH_TIERS },
	'fill.gradient.path': { fields: ['gradient', 'fill'], tiers: BOTH_TIERS },
	'line.gradient': { fields: ['gradient', 'line'], tiers: BOTH_TIERS },
	'line.gradient.path': { fields: ['gradient', 'line'], tiers: BOTH_TIERS },
	'line.gradient.schemeToken': { fields: ['gradient', 'line'], tiers: BOTH_TIERS },
	// Recorded only when an image-filled surface cannot carry its *bytes* — a linked blip, an
	// SVG the write path refuses, a part missing from the package. The fill option is then
	// absent from the output entirely, so the difference lands on `fill` itself.
	'fill.picture': { fields: ['fill'], tiers: BOTH_TIERS },
	// Empty, and deliberately so: this note declares that a picture fill's tiling, crop, DPI
	// and rotWithShape do not survive, and *none of them is in the IR on either side* — the
	// write API expresses a picture fill as bytes plus transparency, so the converter never
	// emits them. Widening this to `fill` would be the mistake it looks like a fix for: it
	// would excuse an image fill that failed to come back at all, which is the thing the
	// round trip is here to catch.
	'fill.picture.geometry': { fields: [], tiers: BOTH_TIERS },
	'fill.gradient.schemeToken': { fields: ['gradient', 'fill'], tiers: BOTH_TIERS },
	// A hatch colour outside the tokens the write path maps, baked to a literal. The table twins
	// below are scoped to their own option, like the gradient-stop ones.
	'fill.pattern.schemeToken': { fields: ['fill.pattern.fgColor', 'fill.pattern.bgColor'], tiers: BOTH_TIERS },
	'fill.schemeToken': { fields: ['fill', 'fill.color'], tiers: BOTH_TIERS },
	// The outline's twin on a glow, and `shadow.schemeToken` below on a shadow: a colour outside the
	// ten tokens, baked to a literal.
	'glow.schemeToken': { fields: ['glow.color'], tiers: BOTH_TIERS },
	'graphicFrame.unknown': { fields: ['*'], tiers: BOTH_TIERS },
	'group.child': { fields: ['*'], tiers: BOTH_TIERS },
	'group.childSpace': { fields: ['x', 'y', 'w', 'h', 'rotate', 'flipH', 'flipV'], tiers: BOTH_TIERS },
	'group.empty': { fields: ['*'], tiers: BOTH_TIERS },
	'group.transform': { fields: ['rotate', 'flipH', 'flipV'], tiers: BOTH_TIERS },
	// A picture crop with no `crop` spelling is left off, so the difference lands on the option.
	'image.crop': { fields: ['crop'], tiers: BOTH_TIERS },
	'image.data': { fields: ['data', '$asset'], tiers: BOTH_TIERS },
	'image.recolor': { fields: ['duotone', 'grayscale', 'biLevel', 'clrChange'], tiers: BOTH_TIERS },
	// Covers the picture's bytes, not just an `svg` option: an SVG picture's raster fallback
	// is regenerated rather than carried, so the blip the round trip compares is a different
	// image from the source's — which is precisely the loss this note is about.
	'image.svg': { fields: ['svg', 'data', '$asset'], tiers: BOTH_TIERS },
	'line.arrowSize': { fields: ['beginArrowType', 'endArrowType'], tiers: BOTH_TIERS },
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
	// Deck-scoped, so anchored at the root: a layout's background and title sit under
	// `chrome.masters`, and a renamed layout also reaches every slide bound to it as `layoutName`.
	'master.background': { fields: ['chrome.masters.background.*'], tiers: STANDALONE_ONLY },
	// A background's gradient and pattern take the shape fill's notes, under the tier's construct.
	'master.background.gradient': { fields: ['chrome.masters.background.*'], tiers: STANDALONE_ONLY },
	'master.background.gradient.path': { fields: ['chrome.masters.background.*'], tiers: STANDALONE_ONLY },
	'master.background.gradient.schemeToken': { fields: ['chrome.masters.background.*'], tiers: STANDALONE_ONLY },
	'master.background.pattern.schemeToken': { fields: ['chrome.masters.background.*'], tiers: STANDALONE_ONLY },
	'master.colorMap': { fields: [], tiers: STANDALONE_ONLY },
	'master.decoration': { fields: [], tiers: STANDALONE_ONLY },
	// Scoped to the `DEFAULT` layout's title, so a suffix: the whole added layout.
	'master.default': { fields: ['master'], tiers: STANDALONE_ONLY },
	'master.multiple': { fields: [], tiers: STANDALONE_ONLY },
	// Both are about a layout *title* the standalone tier has to invent: deduplicated because it doubles
	// as a lookup key, whitespace-collapsed because the write path emits it as a raw XML attribute
	// value. The template-anchored tier keeps the layout's own `p:cSld@name`.
	'master.name': { fields: ['chrome.masters.title', 'layoutName'], tiers: STANDALONE_ONLY },
	'master.nameCollision': { fields: ['chrome.masters.title', 'layoutName'], tiers: STANDALONE_ONLY },
	'master.placeholders': { fields: [], tiers: STANDALONE_ONLY },
	'master.txStyles': { fields: [], tiers: STANDALONE_ONLY },
	'theme.fmtScheme': { fields: [], tiers: STANDALONE_ONLY },
	// The two layout-shape notes with no slide counterpart to inherit a mapping from. Both are
	// empty for the same reason, and it is worth spelling out because both *look* like they
	// should exclude something. A table on a layout is absent from the `objects` array on both
	// sides — the source's because this converter skips it, the output's because it was never
	// written — so there is no difference to excuse. A flattened group is stronger still: the
	// source layout's group becomes N loose objects here and the output layout genuinely *has*
	// N loose objects, so the two agree exactly and the note is a caveat for a human reading the
	// emitted script, not an exclusion.
	'layout.decoration': { fields: [], tiers: STANDALONE_ONLY },
	'layout.group': { fields: [], tiers: STANDALONE_ONLY },
	// `a:ln/@algn="in"` has no write option, so the stroke comes back centred on the edge. The
	// IR carries no alignment on either side, so there is nothing to exclude and the note is a
	// caveat for a human reading the emitted script.
	'line.align': { fields: [], tiers: BOTH_TIERS },
	// Path-qualified: the entries are matched as a suffix of the difference's own path, and a
	// bare `width` also excused a table cell's bevel or border width.
	'line.dash': { fields: ['line.dashType'], tiers: BOTH_TIERS },
	// An outline colour outside the ten scheme tokens the write path maps, baked to a literal.
	'line.schemeToken': { fields: ['line.color'], tiers: BOTH_TIERS },
	'line.width': { fields: ['line.width'], tiers: BOTH_TIERS },
	'media.audioVideo': { fields: ['*'], tiers: BOTH_TIERS },
	'notes.formatting': { fields: ['notesText'], tiers: BOTH_TIERS },
	'shadow.schemeToken': { fields: ['shadow.color'], tiers: BOTH_TIERS },
	'shape.custGeom.guides': { fields: ['points'], tiers: BOTH_TIERS },
	'shape.effects': { fields: ['shadow', 'glow'], tiers: BOTH_TIERS },
	'shape.empty': { fields: ['*'], tiers: BOTH_TIERS },
	'shape.frameInherited': { fields: ['x', 'y', 'w', 'h'], tiers: BOTH_TIERS },
	// A shape nothing places is dropped whole, so every difference on it is the loss.
	'shape.frameUnresolved': { fields: ['*'], tiers: BOTH_TIERS },
	'shape.hidden': { fields: ['*'], tiers: BOTH_TIERS },
	'shape.placeholder': { fields: ['placeholder'], tiers: BOTH_TIERS },
	'slide.animation': { fields: [], tiers: BOTH_TIERS },
	// Slide-scoped, so anchored at the root. The diff compares the slide's `background` as one
	// value, so the colour, transparency and picture bytes all land on that one path.
	'slide.background': { fields: ['background'], tiers: BOTH_TIERS },
	'slide.background.gradient': { fields: ['background'], tiers: BOTH_TIERS },
	'slide.background.gradient.path': { fields: ['background'], tiers: BOTH_TIERS },
	'slide.background.gradient.schemeToken': { fields: ['background'], tiers: BOTH_TIERS },
	'slide.background.pattern.schemeToken': { fields: ['background'], tiers: BOTH_TIERS },
	'slide.carried': { fields: ['*'], tiers: TEMPLATE_ONLY },
	'slide.layout': { fields: ['layoutName'], tiers: BOTH_TIERS },
	'slide.name': { fields: [], tiers: BOTH_TIERS },
	// A transition the write vocabulary cannot name is dropped whole, so the difference lands
	// on the slide's `transition` key itself. Deliberately *not* widened to the keys inside it:
	// this note is only ever recorded when the whole transition is gone.
	'slide.transition': { fields: ['transition'], tiers: BOTH_TIERS },
	// The sound alone, one level down. Scoped to `transition.sound` so it cannot also excuse a
	// transition whose type or timing came back wrong — the loss it declares is exactly the missing
	// `p:sndAc`, and the `.*` also covers the case where the sound survives with other bytes.
	'slide.transitionSound': { fields: ['transition.sound.*'], tiers: BOTH_TIERS },
	// A dash outside `ST_PresetLineDashVal` cannot be written back, so the edge comes out as
	// a plain dashed rule. Scoped to `border`, which is where that difference lands.
	'table.cell.borders.dash': { fields: ['border', 'diagonal'], tiers: BOTH_TIERS },
	// `table.cell.fill` used to live here, for a cell whose own fill could not be told apart
	// from the one it inherited from the table style. `TableCell.hasOwnFill` tells them apart,
	// so the mapper emits the right one and records nothing — the note is gone rather than
	// unmapped. Its `.gradient` / `.picture` children below are separate constructs and stay.
	// Empty, and correctly so: `a:tc/@id` and `a:tcPr/a:headers` have no write option — the
	// IR has nowhere to put them on either side, so there is nothing to exclude. The note
	// exists so a reader of the emitted script learns the association was there and is gone.
	'table.cell.headers': { fields: [], tiers: BOTH_TIERS },
	// A field or an equation in a cell contributes no run, so its text is in neither IR. Not
	// `text.field`'s `*`, which on a cell's note would excuse the whole table.
	'table.cell.field': { fields: [], tiers: BOTH_TIERS },
	'table.cell.equation': { fields: [], tiers: BOTH_TIERS },
	// The cell-side twins of `fill.picture` / `fill.picture.geometry`, and mapped for the
	// same reasons.
	'table.cell.fill.picture': { fields: ['fill'], tiers: BOTH_TIERS },
	'table.cell.fill.picture.geometry': { fields: [], tiers: BOTH_TIERS },
	// The table-side twins of `fill.schemeToken` / `text.color.schemeToken`: one of the seven
	// `ST_SchemeColorVal` tokens the write path's `clrMap` does not carry, baked to the literal
	// it resolves to. Three sites used to pass one through RAW, so the generated script warned
	// `color/invalid-value` and painted the default text colour instead.
	'table.cell.fill.schemeToken': { fields: ['fill', 'fill.color'], tiers: BOTH_TIERS },
	'table.cell.borders.schemeToken': { fields: ['border', 'diagonal'], tiers: BOTH_TIERS },
	// The table-background twins. Scoped to `tableFill` rather than `fill`, because those are
	// two different options: one lands on `a:tblPr`, the other is stamped onto every cell.
	'table.fill.schemeToken': { fields: ['tableFill', 'tableFill.color'], tiers: BOTH_TIERS },
	'table.fill.picture': { fields: ['tableFill'], tiers: BOTH_TIERS },
	'table.fill.picture.geometry': { fields: [], tiers: BOTH_TIERS },
	// A gradient that cannot be expressed falls back to no gradient, so the difference lands
	// on the fill option itself — `tableFill` for the background, `fill` for a cell.
	'table.fill.gradient': { fields: ['tableFill'], tiers: BOTH_TIERS },
	'table.fill.gradient.path': { fields: ['tableFill'], tiers: BOTH_TIERS },
	// A gradient STOP whose scheme token the write path cannot carry, baked to a literal. Scoped
	// to the surface the stop sits on, like the two above it: the stop note used to hardcode the
	// shape spelling, so a table gradient recorded `fill.gradient.schemeToken` while its
	// difference landed on `tableFill` and the note could never match it.
	'table.fill.gradient.schemeToken': { fields: ['tableFill'], tiers: BOTH_TIERS },
	'table.fill.pattern.schemeToken': {
		fields: ['tableFill.pattern.fgColor', 'tableFill.pattern.bgColor'],
		tiers: BOTH_TIERS,
	},
	'table.cell.fill.gradient': { fields: ['fill'], tiers: BOTH_TIERS },
	'table.cell.fill.gradient.path': { fields: ['fill'], tiers: BOTH_TIERS },
	'table.cell.fill.gradient.schemeToken': { fields: ['fill'], tiers: BOTH_TIERS },
	'table.cell.fill.pattern.schemeToken': {
		fields: ['fill.pattern.fgColor', 'fill.pattern.bgColor'],
		tiers: BOTH_TIERS,
	},
	// Narrowed to the East-Asian `ST_TextVerticalType` modes `textDirection` cannot spell —
	// the four it can now round-trip, so this no longer excuses every vertical cell.
	'table.cell.vert': { fields: ['textDirection', 'vert'], tiers: BOTH_TIERS },
	'table.rowAuto': { fields: ['rowH'], tiers: BOTH_TIERS },
	'table.style': { fields: ['tableStyle'], tiers: BOTH_TIERS },
	'text.align': { fields: ['align'], tiers: BOTH_TIERS },
	// `text.bullet.numberStartAt` and `text.bullet.style` used to live here, for
	// `a:buAutoNum/@startAt` and for a bullet's own font/size/colour. `Paragraph.bulletDetail`
	// reads all four, so the mapper emits `numberStartAt` / `fontFace` / `size` / `color` and
	// records nothing — the notes are gone rather than unmapped. What remains of the size half
	// is `text.bullet.sizePt`, which is a genuinely unwritable unit rather than an unread value.
	'text.bullet.numberType': { fields: ['bullet'], tiers: BOTH_TIERS },
	// An absolute bullet size (`a:buSzPts`) has no write option at all — `bullet.size` is a
	// percentage of the run size — so the difference lands on the bullet option.
	'text.bullet.sizePt': { fields: ['bullet'], tiers: BOTH_TIERS },
	// A percentage outside 25–400%, which the write path rejects with a warning and replaces
	// with the run's own size.
	'text.bullet.sizePct': { fields: ['bullet'], tiers: BOTH_TIERS },
	// A bullet colour outside the ten scheme tokens the write path maps, baked to a literal.
	'text.bullet.schemeToken': { fields: ['bullet'], tiers: BOTH_TIERS },
	// A picture bullet (`a:buBlip`): readable, and `bullet.image` could author it, but the
	// paragraph mapper carries no asset resolver to re-embed the bytes with.
	'text.bullet.picture': { fields: ['bullet'], tiers: BOTH_TIERS },
	// A run's colour, wherever the run sits: a shape, a group child or a table cell.
	'text.color.default': { fields: ['options.color'], tiers: BOTH_TIERS },
	'text.color.inherited': { fields: ['options.color'], tiers: BOTH_TIERS },
	'text.color.schemeToken': { fields: ['options.color'], tiers: BOTH_TIERS },
	'text.equation': { fields: ['*'], tiers: BOTH_TIERS },
	'text.field': { fields: ['*'], tiers: BOTH_TIERS },
	// A run link the run-level `hyperlink` cannot spell is left off, so the difference lands there.
	'text.hyperlink': { fields: ['hyperlink'], tiers: BOTH_TIERS },
	// A link that states no underline comes back stating `u="sng"`, the write path's default for a link.
	'text.hyperlink.underline': { fields: ['options.underline'], tiers: BOTH_TIERS },
	'text.bullet.glyph': { fields: ['bullet'], tiers: BOTH_TIERS },
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
	'text.autofit.fontScale': { fields: ['fit'], tiers: BOTH_TIERS },
	'text.autofit.lnSpcReduction': { fields: ['fit'], tiers: BOTH_TIERS },
	'text.paraSpaceZero': { fields: ['paraSpaceBefore', 'paraSpaceAfter'], tiers: BOTH_TIERS },
	'text.vert': { fields: ['vert'], tiers: BOTH_TIERS },
} as const satisfies Record<string, NoteConstructEntry>

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
	/**
	 * Source shape name (`p:cNvPr/@name`), `''` for a shape that has none, or `null` when the loss is
	 * not shape-scoped.
	 */
	shapeName: string | null
	/**
	 * Stable dotted identifier for the lost construct — `line.width`, `text.tabStops`,
	 * `chart.workbook`. Matched mechanically by the round-trip check, so it must name a
	 * field path rather than describe one. Reuse an existing key for the same construct
	 * instead of coining a synonym.
	 *
	 * A `string` rather than `NoteConstruct` because a note carries the `layout.`-prefixed
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
 * prepended to the construct by `layoutShapeScope`.
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
 * `true` when a note describes the output of the printer for `tier`.
 *
 * A construct's catalogue entry decides. A `layout.`-prefixed construct with no entry of its own is
 * a slide construct borrowed for a loss in re-authoring a layout's decoration, so whatever its
 * slide twin says, it applies only to the standalone tier: the template-anchored one does not
 * rebuild layouts.
 */
export function noteAppliesTo(construct: string, tier: PrintTier): boolean {
	const entry = (NOTE_CONSTRUCTS as Record<string, NoteConstructEntry | undefined>)[construct]
	if (entry) return entry.tiers.includes(tier)
	if (construct.startsWith(LAYOUT_NOTE_PREFIX)) return tier === 'standalone'
	return true
}

/**
 * `a`, `a and b`, `a, b and c` — a note is a sentence a human reads, and a bare
 * `join(', ')` made a two-item one read as a list that had lost its last member.
 */
export function andList(items: readonly string[]): string {
	if (items.length < 2) return items.join('')
	return `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`
}

/** A graphic frame payload with no write-API emitter. */
export type UnwritableFramePayload = 'chartEx' | 'diagram' | 'unknown'

/**
 * The note each {@link UnwritableFramePayload} records.
 *
 * Here rather than beside the shape mapper that records them, because the printer reads it too
 * and the printer does not reach into `from-read/`. `graphicFrameCall` picks its note from it. The
 * template-anchored printer keeps these on a slide it copies whole and drops every other note
 * recorded there: these say which construct forced the copy, and the rest describe a
 * transcription that tier never prints.
 */
export const UNWRITABLE_FRAME_CONSTRUCTS = {
	chartEx: 'chartEx.all',
	diagram: 'diagram.all',
	unknown: 'graphicFrame.unknown',
} as const satisfies Record<UnwritableFramePayload, NoteConstruct>

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
