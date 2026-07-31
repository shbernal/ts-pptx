/**
 * The round-trip oracle: does the deck a printed script builds match the deck it was read
 * from, everywhere the fidelity notes did not say it would not?
 *
 * **Why a projection diff and not byte identity.** The output package is not the input
 * package and never will be — different rel ids, different shape ids, a regenerated shape
 * tree, `dirty="0"` on every run. Comparing bytes would report a total mismatch for every
 * deck and prove nothing. So the comparison happens one level up, on the IR: run the same
 * `readModelToIr` over the source deck and over the deck the script produced, and compare
 * those. Two decks that produce the same IR make the same write-API calls, which is exactly
 * the claim a generated script makes.
 *
 * **What a note is worth here.** A {@link FidelityNote} says a named construct will not
 * survive. That makes it an *exclusion with a signature*: the difference it predicts is
 * allowed, and every other difference is a defect. Without the notes there would be no way
 * to tell a known cost from a regression, and the check could only ever be a snapshot.
 *
 * **Two blind spots, both stated plainly, because a clean report is worth exactly what its
 * limits allow.**
 *
 * 1. *The reader is shared.* Both IRs come from `readModelToIr`, so a construct the read
 *    model cannot see is missing from both and compares equal. This check cannot certify
 *    anything the read path does not read; `pnpm run read:census` measures that surface, and
 *    the two are complementary rather than overlapping.
 * 2. *The converter need not be injective.* Two different source constructs can map to the
 *    same call, and then the round trip agrees while the deck has changed. Measured, not
 *    hypothetical: the mapper once read an explicit `a:buNone` as the glyph `n`, so a
 *    bulletless paragraph gained a bullet — and because the output's real `n` bullet read
 *    back through the same wrong mapping, the diff was clean. Mutation testing is what
 *    catches that class, and a mutation that survives here is a statement about this
 *    check's reach rather than a mistake in it.
 *
 * A clean report means "nothing the converter can distinguish was lost", never "nothing was
 * lost".
 */
import type { FidelityNote } from '../fidelity.js'
import type { CanonicalCall, CanonicalChrome, CanonicalDeck, CanonicalSlide } from './canonical.js'
import type { IrValue } from '../ir.js'

/**
 * Which way a difference points.
 *
 * `lost` and `changed` are the ones the fidelity contract governs — the source said
 * something the output does not, or says differently. `added` is the write path being
 * explicit where the source was implicit, which is *usually* harmless and is reported
 * separately for a reason worth stating: this diff cannot tell whether the value the write
 * path chose is the one the source was inheriting. An `added` `color` is benign if the
 * shape was inheriting that colour and a silent repaint if it was not.
 */
export type DifferenceKind = 'lost' | 'added' | 'changed'

/**
 * Options the write path materialises that the source deck left implicit, with the reason
 * each is not judgeable from here.
 *
 * This list exists because `added` differences cannot simply be waved through. Mutation
 * testing settled that: dropping the converter's explicit `valign` left the write path to
 * centre every text body, which showed up as an `added` and nothing failed. So `added` is
 * held to the same standard as the rest — a note must declare it, or its field must appear
 * here with a reason.
 *
 * Every entry names a write-path default whose *correctness* the round trip has no way to
 * assess, because the source's counterpart was inherited and the read model cannot resolve
 * what it was inheriting. That is an honest "unknown", not an "ignore" — each one is a
 * candidate for the read-side backlog, and closing any of them turns its entry into a real
 * comparison.
 */
const WRITER_DEFAULTS: Record<string, string> = {
	// The write path copies the object-level and first-run character properties onto every
	// run it emits. The source's bare runs inherited theirs from the list style, which is
	// unread, so there is nothing to compare the copied value against.
	fontSize: 'runs inherit size from an unread list style; the write path copies the shape size onto each run',
	options: 'a run or cell that carried no options at all in the source now carries the ones below',
	// `a:tcBdr` borders come from the table style graph, which `resolvedFill`/`borders` fold
	// together rather than separating; the write path spells out an explicit border per edge.
	border:
		'table cell borders are resolved through the style graph, which the read model cannot separate from the cell’s own',
	// `a:ln/@w` defaults to a hairline in OOXML and to 1pt here, and a source outline with no
	// explicit width is exactly the theme-`lnRef` case `line.width` already declares unread.
	width:
		'an outline with no explicit width takes the write path’s 1pt default; the source’s came from the unread theme line style',
	// A shape with no geometry of its own — an unfilled placeholder — has no preset to name.
	shape: 'a shape that inherits its geometry from a layout is emitted as a plain rectangle',
	type: 'a fill or line with no explicit kind takes the write path’s solid default',
	// Chart options the write path always emits; the read side reports only what the chart
	// part contains, and these are defaults rather than authored values.
	chartColors: 'the write path assigns a series palette; the source’s came from the theme',
	dataLabelFormatCode: 'a data label with no explicit number format takes the write path’s default',
	showLeaderLines: 'leader lines are on by default in the write path',
	// The five docProps the write path seeds in its constructor ('ts-pptx',
	// 'ts-pptx Presentation', '1'). A source deck that declared none of a given property gets
	// the library's, and there is no way to unset one — assigning `''` writes an empty element
	// rather than removing it. Unlike the entries above, this is a *write*-side candidate, not
	// a read-side one: the fix is a way to author a deck with a property genuinely absent.
	// Restricted to `added` by construction, so a printer that stopped writing docProps still
	// fails: the source's own values become `changed`, which nothing here excuses.
	title: 'a deck that declares no title gets the write path’s own',
	author: 'a deck that declares no author gets the write path’s own',
	subject: 'a deck that declares no subject gets the write path’s own',
	company: 'a deck that declares no company gets the write path’s own',
	revision: 'a deck that declares no revision gets the write path’s own',
}

export interface IrDifference {
	/** 1-based source slide. */
	slideNumber: number
	/** Source shape name, when the difference sits inside a call that has one. */
	shapeName: string | null
	/**
	 * Shape names nested inside the same call — a group's children. A note scoped to one of
	 * these also declares differences reported against the call that carries it.
	 */
	nestedNames: string[]
	/** Where in the slide's projection, e.g. `calls[4].args[1].line.width`. */
	path: string
	/**
	 * The option name the difference is about — the last object key on {@link path}, or
	 * `'*'` when a whole call is present on one side only. This is what a note's `construct`
	 * is matched against, which is why `construct` has to name a field path rather than
	 * describe one.
	 */
	field: string
	kind: DifferenceKind
	/** The source value, JSON-encoded and truncated for reporting. `'—'` when absent. */
	expected: string
	/** The output value, likewise. */
	actual: string
	/** The note that declares this difference, when one does. */
	declaredBy: FidelityNote | null
}

export interface RoundTripReport {
	slideCount: number
	/** Every difference found, in deck order. */
	differences: IrDifference[]
	/**
	 * Differences neither a note nor a known write-path default accounts for.
	 * **These are the defects**, and the number this check gates on.
	 */
	undeclared: IrDifference[]
	/** Differences a note declared. The contract working. */
	declared: IrDifference[]
	/** `added` differences: the write path spelling out what the source left implicit. */
	added: IrDifference[]
	/**
	 * Notes that matched no difference.
	 *
	 * **Not a defect signal, despite the obvious reading.** Most notes name constructs the
	 * read model cannot see at all — a bullet's `a:buFont`, a shape's theme `effectRef` —
	 * so they are absent from *both* IRs and there is nothing for them to match. A note
	 * here is a candidate for review, not a stale claim.
	 */
	unmatchedNotes: FidelityNote[]
}

/**
 * Option names each note construct is a promise about.
 *
 * This table is the fidelity catalogue made mechanical. A note claims a construct will not
 * survive; without a mapping from that claim to the fields it covers, the claim cannot
 * exclude anything and the round trip degenerates into a snapshot.
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
const NOTE_FIELDS: Record<string, readonly string[]> = {
	'chart.blanks': ['values'],
	'chart.combo': ['type'],
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
	'fill.gradient.path': ['gradient', 'fill'],
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
	// here rather than the exception: `a:fmtScheme`, `p:txStyles`, master/layout decoration and
	// layout placeholder definitions are all absent from the IR on *both* sides — nothing reads
	// them — so there is nothing for a note to exclude and the note exists for a human.
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
	'line.dash': ['dashType'],
	'line.width': ['width'],
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
	// The table-background twins. Scoped to `tableFill` rather than `fill`, because those are
	// two different options: one lands on `a:tblPr`, the other is stamped onto every cell.
	'table.fill.picture': ['tableFill'],
	'table.fill.picture.geometry': [],
	// A gradient that cannot be expressed falls back to no gradient, so the difference lands
	// on the fill option itself — `tableFill` for the background, `fill` for a cell.
	'table.fill.gradient': ['tableFill'],
	'table.fill.gradient.path': ['tableFill'],
	'table.cell.fill.gradient': ['fill'],
	'table.cell.fill.gradient.path': ['fill'],
	// Narrowed to the East-Asian `ST_TextVerticalType` modes `textDirection` cannot spell —
	// the four it can now round-trip, so this no longer excuses every vertical cell.
	'table.cell.vert': ['textDirection', 'vert'],
	'table.rowAuto': ['rowH'],
	'table.style': ['tableStyle'],
	'text.align': ['align'],
	'text.bullet.numberStartAt': ['bullet'],
	'text.bullet.numberType': ['bullet'],
	'text.bullet.style': ['bullet'],
	'text.color.default': ['color'],
	'text.color.inherited': ['color'],
	'text.color.schemeToken': ['color'],
	'text.equation': ['*'],
	'text.field': ['*'],
	'text.bullet.glyph': ['bullet'],
	'text.bullet.inherited': ['bullet'],
	'text.indent': [],
	'text.paraSpaceZero': ['paraSpaceBefore', 'paraSpaceAfter'],
	'text.vert': ['vert'],
}

/** Every note construct this table knows how to match, for a test to check against. */
export function knownNoteConstructs(): string[] {
	return Object.keys(NOTE_FIELDS).sort()
}

/**
 * Compare the deck a script was generated from against the deck it produced.
 *
 * `notes` must be the set that actually applies to the printed output — `printScript`'s
 * returned `notes`, not `DeckIr.fidelity`. The two differ in both directions: the printer
 * suppresses notes its tier rescues and adds ones its tier causes, so diffing against the
 * IR's own list would both over- and under-exclude.
 */
export function diffDeckIr(expected: CanonicalDeck, actual: CanonicalDeck, notes: FidelityNote[]): RoundTripReport {
	const differences: IrDifference[] = []

	if (expected.slideSize.widthEmu !== actual.slideSize.widthEmu) {
		differences.push(scalarDifference(0, null, [], 'slideSize.widthEmu', 'widthEmu', expected.slideSize.widthEmu, actual.slideSize.widthEmu)) // prettier-ignore
	}
	if (expected.slideSize.heightEmu !== actual.slideSize.heightEmu) {
		differences.push(scalarDifference(0, null, [], 'slideSize.heightEmu', 'heightEmu', expected.slideSize.heightEmu, actual.slideSize.heightEmu)) // prettier-ignore
	}

	diffValue(expected.props, actual.props, 'props', 'props', 0, null, [], differences)
	diffChrome(expected.chrome, actual.chrome, differences)

	const slideCount = Math.max(expected.slides.length, actual.slides.length)
	for (let index = 0; index < slideCount; index++) {
		const before = expected.slides[index]
		const after = actual.slides[index]
		const number = before?.number ?? after?.number ?? index + 1
		if (!before || !after) {
			differences.push({
				slideNumber: number,
				shapeName: null,
				nestedNames: [],
				path: `slides[${index}]`,
				field: '*',
				kind: before ? 'lost' : 'added',
				expected: before ? 'slide' : '—',
				actual: after ? 'slide' : '—',
				declaredBy: null,
			})
			continue
		}
		diffSlide(before, after, differences)
	}

	for (const difference of differences) difference.declaredBy = declaringNote(difference, notes)
	const matched = new Set(differences.map((difference) => difference.declaredBy).filter(Boolean))
	const accounted = (difference: IrDifference): boolean =>
		difference.declaredBy !== null || (difference.kind === 'added' && difference.field in WRITER_DEFAULTS)

	return {
		slideCount: expected.slides.length,
		differences,
		undeclared: differences.filter((difference) => !accounted(difference)),
		declared: differences.filter((difference) => difference.declaredBy !== null),
		added: differences.filter((difference) => difference.kind === 'added'),
		unmatchedNotes: notes.filter((note) => !matched.has(note)),
	}
}

/**
 * The theme and the layout gallery.
 *
 * Reported at `slideNumber: 0`, the same deck-level bucket the slide size uses, so a
 * deck-scoped note (which carries `slideNumber: null` and matches any slide) can declare one.
 * Masters align by **title**, not by position, for the same reason calls align by shape name:
 * the write path always emits a blank layout of its own at gallery position 0, so positional
 * alignment reports every layout in the deck as renamed and buries the one real difference —
 * that extra layout — under a shift.
 *
 * And, exactly as with shape names, a title is only usable as a key where it is *unique on
 * both sides*. Uniqueness is supposed to be an invariant here — the converter deduplicates
 * titles because a title is also what `addSlide({ masterTitle })` binds on — but keying on it
 * unconditionally makes the check assume the thing it should be testing: a mutation that
 * removed the deduplication left two layouts sharing a title, both expected entries matched
 * the same actual one, and the round trip came back clean. Falling back to position when the
 * key repeats costs nothing when the invariant holds and catches it when it does not.
 */
function diffChrome(expected: CanonicalChrome, actual: CanonicalChrome, out: IrDifference[]): void {
	diffValue(expected.theme, actual.theme, 'chrome.theme', 'theme', 0, null, [], out)

	// A whole master appearing or disappearing is reported *scoped to its title*, using the
	// same `shapeName` channel a shape-scoped note uses. Without that, one note declaring the
	// blank layout the write path always adds would carry `field: 'master'` and excuse every
	// added layout — which is what a mutation that stopped deduplicating layout titles proved.
	const report = (
		path: string,
		kind: DifferenceKind,
		before: IrValue | undefined,
		after: IrValue | undefined
	): void => {
		out.push({
			slideNumber: 0,
			shapeName: titleOf(before ?? after ?? null),
			nestedNames: [],
			path,
			field: 'master',
			kind,
			expected: brief(before),
			actual: brief(after),
			declaredBy: null,
		})
	}

	const byTitle = uniqueByTitle(actual.masters)
	const expectedTitles = uniqueByTitle(expected.masters)
	const claimed = new Set<IrValue>()
	let cursor = 0
	const nextUnclaimed = (): IrValue | undefined => {
		while (cursor < actual.masters.length) {
			const candidate = actual.masters[cursor++]
			if (candidate !== undefined && !claimed.has(candidate)) return candidate
		}
		return undefined
	}

	expected.masters.forEach((before, index) => {
		const title = titleOf(before)
		// Usable only where the title identifies exactly one master on each side.
		const named = title !== null && expectedTitles.has(title) ? byTitle.get(title) : undefined
		const after = named !== undefined && !claimed.has(named) ? named : nextUnclaimed()
		if (after === undefined) {
			report(`chrome.masters[${index}]`, 'lost', before, undefined)
			return
		}
		claimed.add(after)
		diffValue(before, after, `chrome.masters[${index}]`, 'master', 0, null, [], out)
	})

	actual.masters.forEach((after, index) => {
		if (!claimed.has(after)) report(`chrome.masters[${index}]`, 'added', undefined, after)
	})
}

/** Masters keyed by title, keeping only titles that identify exactly one of them. */
function uniqueByTitle(masters: IrValue[]): Map<string, IrValue> {
	const counts = new Map<string, number>()
	for (const master of masters) {
		const title = titleOf(master)
		if (title !== null) counts.set(title, (counts.get(title) ?? 0) + 1)
	}
	const out = new Map<string, IrValue>()
	for (const master of masters) {
		const title = titleOf(master)
		if (title !== null && counts.get(title) === 1) out.set(title, master)
	}
	return out
}

/** A canonical master's `title`, which is its identity across the round trip. */
function titleOf(master: IrValue): string | null {
	if (master === null || typeof master !== 'object' || Array.isArray(master)) return null
	const title = (master as Record<string, IrValue>)['title']
	return typeof title === 'string' ? title : null
}

function diffSlide(expected: CanonicalSlide, actual: CanonicalSlide, out: IrDifference[]): void {
	const number = expected.number
	const at = (path: string, field: string, a: IrValue, b: IrValue): void => {
		if (JSON.stringify(a) !== JSON.stringify(b)) out.push(scalarDifference(number, null, [], path, field, a, b))
	}
	at('hidden', 'hidden', expected.hidden, actual.hidden)
	at('layoutName', 'layoutName', expected.layoutName, actual.layoutName)
	at('background', 'background', expected.background, actual.background)
	at('notesText', 'notesText', expected.notesText, actual.notesText)
	// Structural rather than whole-value, unlike the four above: a transition is a small object
	// whose parts are lost independently — the append path drops an embedded sound while
	// keeping the effect — and a whole-value compare would force the note that declares the
	// sound to carry `field: 'transition'`, which would then excuse a wrong type or duration too.
	diffValue(expected.transition, actual.transition, 'transition', 'transition', number, null, [], out)

	for (const [before, after] of alignCalls(expected.calls, actual.calls)) {
		const name = before?.shapeName ?? after?.shapeName ?? null
		// Union of both sides: a child that dropped out is named only by the source's call.
		const nested = [...new Set([...(before?.containedNames ?? []), ...(after?.containedNames ?? [])])]
		const path = `calls[${before ? expected.calls.indexOf(before) : actual.calls.indexOf(after as CanonicalCall)}]`
		if (!before || !after) {
			out.push({
				slideNumber: number,
				shapeName: name,
				nestedNames: nested,
				path,
				field: '*',
				kind: before ? 'lost' : 'added',
				expected: before ? before.method : '—',
				actual: after ? after.method : '—',
				declaredBy: null,
			})
			continue
		}
		if (before.method !== after.method) {
			out.push(scalarDifference(number, name, nested, `${path}.method`, 'method', before.method, after.method))
		}
		diffValue(before.args, after.args, `${path}.args`, 'args', number, name, nested, out)
	}
}

/**
 * Pair up the two slides' calls.
 *
 * Position alone is wrong the moment one shape drops out: every later call shifts by one
 * and the report fills with mismatches that are really one missing shape. Shape names
 * survive the round trip (the converter passes `objectName` through), so they are the
 * alignment key wherever they are present and unambiguous, with position as the fallback
 * for unnamed or duplicate-named shapes.
 */
function alignCalls(
	expected: CanonicalCall[],
	actual: CanonicalCall[]
): Array<[CanonicalCall | null, CanonicalCall | null]> {
	const usable = (calls: CanonicalCall[]): Map<string, CanonicalCall> => {
		const counts = new Map<string, number>()
		for (const call of calls) {
			if (call.shapeName === null) continue
			counts.set(call.shapeName, (counts.get(call.shapeName) ?? 0) + 1)
		}
		const out = new Map<string, CanonicalCall>()
		for (const call of calls) {
			if (call.shapeName !== null && counts.get(call.shapeName) === 1) out.set(call.shapeName, call)
		}
		return out
	}

	const byName = usable(actual)
	const claimed = new Set<CanonicalCall>()
	const pairs: Array<[CanonicalCall | null, CanonicalCall | null]> = []
	// Positional cursor for calls that names cannot align, advanced only past claimed calls
	// so a named match does not consume an unnamed call's slot.
	let cursor = 0
	const nextUnclaimed = (): CanonicalCall | null => {
		while (cursor < actual.length) {
			const candidate = actual[cursor++]
			if (candidate && !claimed.has(candidate)) return candidate
		}
		return null
	}

	for (const call of expected) {
		const named = call.shapeName === null ? undefined : byName.get(call.shapeName)
		if (named && !claimed.has(named)) {
			claimed.add(named)
			pairs.push([call, named])
			continue
		}
		const positional = nextUnclaimed()
		if (positional) claimed.add(positional)
		pairs.push([call, positional])
	}
	for (const call of actual) if (!claimed.has(call)) pairs.push([null, call])
	return pairs
}

function diffValue(
	expected: IrValue,
	actual: IrValue,
	path: string,
	field: string,
	slideNumber: number,
	shapeName: string | null,
	nestedNames: string[],
	out: IrDifference[]
): void {
	if (expected === actual) return
	const kindOf = (value: IrValue): string => (value === null ? 'null' : Array.isArray(value) ? 'array' : typeof value)
	if (kindOf(expected) !== kindOf(actual)) {
		out.push(scalarDifference(slideNumber, shapeName, nestedNames, path, field, expected, actual))
		return
	}

	if (Array.isArray(expected) && Array.isArray(actual)) {
		const length = Math.max(expected.length, actual.length)
		for (let index = 0; index < length; index++) {
			const before = expected[index]
			const after = actual[index]
			if (before === undefined || after === undefined) {
				out.push({
					slideNumber,
					shapeName,
					nestedNames,
					path: `${path}[${index}]`,
					field,
					kind: before === undefined ? 'added' : 'lost',
					expected: brief(before),
					actual: brief(after),
					declaredBy: null,
				})
				continue
			}
			diffValue(before, after, `${path}[${index}]`, field, slideNumber, shapeName, nestedNames, out)
		}
		return
	}

	if (expected !== null && typeof expected === 'object' && actual !== null && typeof actual === 'object') {
		const record = expected as Record<string, IrValue>
		const other = actual as Record<string, IrValue>
		for (const key of new Set([...Object.keys(record), ...Object.keys(other)])) {
			const before = record[key]
			const after = other[key]
			if (before === undefined || after === undefined) {
				out.push({
					slideNumber,
					shapeName,
					nestedNames,
					path: `${path}.${key}`,
					field: key,
					kind: before === undefined ? 'added' : 'lost',
					expected: brief(before),
					actual: brief(after),
					declaredBy: null,
				})
				continue
			}
			diffValue(before, after, `${path}.${key}`, key, slideNumber, shapeName, nestedNames, out)
		}
		return
	}

	out.push(scalarDifference(slideNumber, shapeName, nestedNames, path, field, expected, actual))
}

function scalarDifference(
	slideNumber: number,
	shapeName: string | null,
	nestedNames: string[],
	path: string,
	field: string,
	expected: IrValue,
	actual: IrValue
): IrDifference {
	return {
		slideNumber,
		shapeName,
		nestedNames,
		path,
		field,
		kind: 'changed',
		expected: brief(expected),
		actual: brief(actual),
		declaredBy: null,
	}
}

/** The first note that covers this difference, or `null` if none does. */
function declaringNote(difference: IrDifference, notes: FidelityNote[]): FidelityNote | null {
	for (const note of notes) {
		if (note.slideNumber !== null && note.slideNumber !== difference.slideNumber) continue
		// A shape-scoped note covers only that shape; a slide- or deck-scoped one covers any.
		if (
			note.shapeName !== null &&
			note.shapeName !== difference.shapeName &&
			!difference.nestedNames.includes(note.shapeName)
		) {
			continue
		}
		const fields = NOTE_FIELDS[note.construct]
		if (!fields) continue
		if (fields.includes('*') || fields.includes(difference.field)) return note
	}
	return null
}

function brief(value: IrValue | undefined): string {
	if (value === undefined) return '—'
	const text = JSON.stringify(value)
	return text.length > 80 ? `${text.slice(0, 80)}…` : text
}
