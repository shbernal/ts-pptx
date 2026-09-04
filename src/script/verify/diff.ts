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
import { alignByKey } from './align.js'
import { LAYOUT_NOTE_PREFIX, NOTE_CONSTRUCTS, type FidelityNote } from '../fidelity.js'
import {
	collectObjectNames,
	type CanonicalCall,
	type CanonicalChrome,
	type CanonicalDeck,
	type CanonicalSlide,
} from './canonical.js'
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
 * what it was inheriting. That is an honest "unknown", not an "ignore" — each one names a
 * gap in the reader that could be closed later, and teaching the reader to resolve one turns
 * its entry here into a real comparison.
 */
const WRITER_DEFAULTS: Record<string, string> = {
	// Keys are dotted option PATHS, matched as a suffix of a difference's own path. A bare key
	// still means "this option wherever it appears", which is right for the ones below that are
	// unambiguous; qualify one the moment the same word names two different things.
	// The write path copies the object-level and first-run character properties onto every
	// run it emits. The source's bare runs inherited theirs from the list style, which is
	// unread, so there is nothing to compare the copied value against.
	fontSize: 'runs inherit size from an unread list style; the write path copies the shape size onto each run',
	options: 'a run or cell that carried no options at all in the source now carries the ones below',
	// `a:tcBdr` borders come from the table style graph, which `resolvedFill`/`borders` fold
	// together rather than separating; the write path spells out an explicit border per edge.
	border:
		'table cell borders are resolved through the style graph, which the read model cannot separate from the cell’s own',
	// The same unknown one level down, for a cell whose border array survives while an individual
	// EDGE gains the write path's completed `{ type, width }`. Under key-only matching these were
	// covered by the bare `type`/`width` entries, which is the accident this qualification removes:
	// those are about a fill's kind and an outline's weight, on shapes.
	'border.type':
		'a table cell border edge the write path completes; the source’s came from the style graph the read model cannot separate',
	'border.width':
		'a table cell border edge the write path completes; the source’s came from the style graph the read model cannot separate',
	// `a:ln/@w` defaults to a hairline in OOXML and to 1pt here, and a source outline with no
	// explicit width is exactly the theme-`lnRef` case `line.width` already declares unread.
	'line.width':
		'an outline with no explicit width takes the write path’s 1pt default; the source’s came from the unread theme line style',
	// A shape with no geometry of its own — an unfilled placeholder — has no preset to name.
	shape: 'a shape that inherits its geometry from a layout is emitted as a plain rectangle',
	'fill.type': 'a fill with no explicit kind takes the write path’s solid default',
	'line.type': 'an outline with no explicit kind takes the write path’s solid default',
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
	// Path-qualified so they reach `docProps` and nothing else: a bare `title` also excused an
	// added chart title and an added layout title in `chrome.masters`.
	'props.title': 'a deck that declares no title gets the write path’s own',
	'props.author': 'a deck that declares no author gets the write path’s own',
	'props.subject': 'a deck that declares no subject gets the write path’s own',
	'props.revision': 'a deck that declares no revision gets the write path’s own',
	'props.company': 'a deck that declares no company gets the write path’s own',
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
 * Every note construct this table knows how to match, for a test to check against. The
 * `layout.` spellings are not listed — they resolve through {@link isKnownNoteConstruct},
 * which is what a corpus check should use.
 */
export function knownNoteConstructs(): string[] {
	return Object.keys(NOTE_CONSTRUCTS).sort()
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

	diffValue(expected.props, actual.props, 'props', 'props', {
		slideNumber: 0,
		shapeName: null,
		nestedNames: [],
		out: differences,
	})
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
		difference.declaredBy !== null ||
		(difference.kind === 'added' && Object.keys(WRITER_DEFAULTS).some((spec) => specCovers(spec, difference)))

	return {
		slideCount: Math.max(expected.slides.length, actual.slides.length),
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
	diffValue(expected.theme, actual.theme, 'chrome.theme', 'theme', {
		slideNumber: 0,
		shapeName: null,
		nestedNames: [],
		out,
	})

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
			nestedNames: decorationNames(before, after),
			path,
			field: 'master',
			kind,
			expected: brief(before),
			actual: brief(after),
			declaredBy: null,
		})
	}

	// Titles align the masters where a title names exactly one on each side; `alignByKey` owns
	// that rule, and the call aligner reads the same one.
	alignByKey(expected.masters, actual.masters, titleOf).forEach(([before, after], index) => {
		if (before === null) {
			report(`chrome.masters[${index}]`, 'added', undefined, after ?? undefined)
			return
		}
		if (after === null) {
			report(`chrome.masters[${index}]`, 'lost', before, undefined)
			return
		}
		// A layout's decoration rides inside this one value, so a note scoped to one of those
		// shapes has no call of its own to be matched against — the same problem a group's child
		// has on a slide, and the same answer.
		diffValue(before, after, `chrome.masters[${index}]`, 'master', {
			slideNumber: 0,
			shapeName: null,
			nestedNames: decorationNames(before, after),
			out,
		})
	})
}

/**
 * The `objectName`s of every decorative shape in a master's props, taking both sides so a
 * shape that dropped out is still named by the source's.
 */
function decorationNames(before: IrValue | undefined, after: IrValue | undefined): string[] {
	const names = new Set<string>()
	for (const master of [before, after]) if (master !== undefined) collectObjectNames(master, names)
	return [...names]
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
	diffValue(expected.transition, actual.transition, 'transition', 'transition', {
		slideNumber: number,
		shapeName: null,
		nestedNames: [],
		out,
	})

	for (const [before, after] of alignByKey(expected.calls, actual.calls, (call) => call.shapeName)) {
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
		diffValue(before.args, after.args, `${path}.args`, 'args', {
			slideNumber: number,
			shapeName: name,
			nestedNames: nested,
			out,
		})
	}
}

/**
 * Where in the deck a difference is being reported, and where the report goes.
 *
 * Four of `diffValue`'s eight positional parameters were this, threaded unchanged through
 * every recursion; only `path` and `field` move as it walks. Bundling them is what makes the
 * recursive calls read as "same site, deeper path".
 */
interface DiffSite {
	slideNumber: number
	shapeName: string | null
	nestedNames: string[]
	out: IrDifference[]
}

function diffValue(expected: IrValue, actual: IrValue, path: string, field: string, site: DiffSite): void {
	const { slideNumber, shapeName, nestedNames, out } = site
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
			diffValue(before, after, `${path}[${index}]`, field, site)
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
			diffValue(before, after, `${path}.${key}`, key, site)
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

/**
 * The dotted option path a difference sits at, array indices removed: `calls[4].args[1].line.width`
 * becomes `['calls', 'args', 'line', 'width']`.
 *
 * Both exclusion mechanisms used to match the terminal key alone, at any depth anywhere in the
 * IR — so `type` (written about a fill's solid default) excused an added `bullet.type`, waving
 * through a character bullet that came back as a numbered list; `title` (written about
 * `docProps`) excused an added chart title; and `width` (written about `a:ln/@w`) excused an
 * added table-cell bevel or border width. Matching a suffix keeps every bare key working where
 * it is unambiguous and lets a spec say WHICH `width` it means.
 */
function pathKeys(difference: IrDifference): string[] {
	const keys = difference.path.split('.').map((segment) => segment.replace(/\[[^\]]*\]/g, ''))
	// A handful of differences are synthesised rather than walked, and name a `field` the path
	// does not end in — a whole master reported at `chrome.masters[11]` with `field: 'master'`.
	// Appending it keeps those matchable while leaving every walked difference untouched.
	return keys[keys.length - 1] === difference.field ? keys : [...keys, difference.field]
}

/**
 * Whether one exclusion spec covers a difference: `'*'` covers the whole call it is scoped to,
 * and anything else must be a dotted SUFFIX of the difference's own path.
 */
function specCovers(spec: string, difference: IrDifference): boolean {
	if (spec === '*') return true
	const wanted = spec.split('.')
	const keys = pathKeys(difference)
	if (wanted.length > keys.length) return false
	return wanted.every((segment, index) => segment === keys[keys.length - wanted.length + index])
}

/** The first note that covers this difference, or `null` if none does. */
function declaringNote(difference: IrDifference, notes: FidelityNote[]): FidelityNote | null {
	for (const note of notes) {
		if (note.slideNumber !== null && note.slideNumber !== difference.slideNumber) continue
		// A layout-shape note is about the chrome, which the diff reports in the deck-level bucket.
		// Without this a themed outline on a layout — recorded with no slide number, and named
		// after a shape whose name repeats between the layout and the slides bound to it — would
		// excuse the same difference on a *slide*, which is a loss nothing declared.
		if (note.construct.startsWith(LAYOUT_NOTE_PREFIX) && difference.slideNumber !== 0) continue
		// A shape-scoped note covers only that shape; a slide- or deck-scoped one covers any.
		if (
			note.shapeName !== null &&
			note.shapeName !== difference.shapeName &&
			!difference.nestedNames.includes(note.shapeName)
		) {
			continue
		}
		const fields = noteFields(note.construct)
		if (!fields) continue
		if (fields.some((field) => specCovers(field, difference))) return note
	}
	return null
}

/**
 * The fields a construct is a promise about.
 *
 * A layout-shape construct falls back to its slide counterpart, which is the whole point of
 * borrowing the vocabulary: `layout.line.width` is `line.width` on a layout and covers exactly
 * the same field. Only the two constructs with no slide twin need entries of their own.
 */
function noteFields(construct: string): readonly string[] | undefined {
	const table: Record<string, readonly string[] | undefined> = NOTE_CONSTRUCTS
	return table[construct] ?? (construct.startsWith(LAYOUT_NOTE_PREFIX) ? table[construct.slice(LAYOUT_NOTE_PREFIX.length)] : undefined) // prettier-ignore
}

/**
 * `true` when the coverage table can match this construct — directly or through the
 * layout-shape fallback. The check a corpus test runs over every note both printers emit.
 */
export function isKnownNoteConstruct(construct: string): boolean {
	return noteFields(construct) !== undefined
}

function brief(value: IrValue | undefined): string {
	if (value === undefined) return '—'
	const text = JSON.stringify(value)
	return text.length > 80 ? `${text.slice(0, 80)}…` : text
}
