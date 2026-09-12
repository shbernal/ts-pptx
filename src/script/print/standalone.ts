/**
 * {@link DeckIr} → a runnable TypeScript module that depends on nothing but this package.
 * **Tier A.**
 *
 * **The trade, stated once.** The template-anchored printer (`printScript`) reuses the source
 * `.pptx` and gets the deck's entire design back byte for byte, at the cost of shipping that
 * file alongside the script and of leaving the design uneditable. This one ships a single
 * `.ts` file. Everything the source deck's chrome contained has to be *re-authored* from what
 * the read model exposes, and the read model deliberately exposes the property tiers a slide
 * inherits — colours, fonts, placeholder geometry — not the parts a slide merely sits on top
 * of. So the output is a deck that renders its slides faithfully and wears a different suit.
 *
 * **The ceiling is measured, not estimated, and no amount of printer work moves it.** Two
 * constructs are unreachable from *both* directions: `a:fmtScheme` (nothing reads it, and the
 * write path emits a hardcoded Office one) and `p:txStyles` (no reader, though
 * `SlideMasterProps.textStyles` could author it). A third, `p:clrMap`, is readable and has no
 * setter. Each is a fidelity note here and a `pass` in the other tier, which is the whole
 * reason the other tier shipped first. A **master's** own decoration joins them for a
 * structural reason rather than a reading one: `defineSlideMaster` creates a layout, so there
 * is no master shape tree to write to. A *layout's* decoration used to be on this list and is
 * not any more — it is transcribed into that layout's `objects`.
 *
 * **One structural difference beyond the chrome.** A slide marked `carried` has no source
 * package to be carried from here, so it is transcribed like any other and loses only the
 * construct that made it uncarryable — which each shape has already declared for itself. That
 * is why {@link DeckIr}'s calls are populated even for a carried slide.
 */
import { uniqueTitle, type DeckIr, type IrValue, type MasterIr } from '../ir.js'
import { NoteCollector, noteAppliesTo, scopeNotes } from '../fidelity.js'
import { printArguments, printString, printValue, type AssetPrinter } from './literal.js'
import {
	assetIdentifiers,
	assetPrinter,
	header,
	printAssetBindings,
	printLayoutSetup,
	printPreamble,
	printSlide,
	printedScript,
	resolvePrintOptions,
	type CommonPrintOptions,
	type PrintedScript,
} from './common.js'

/** The standalone printer's options: exactly the ones both tiers share. */
export type PrintStandaloneScriptOptions = CommonPrintOptions

/** Turn a deck IR into a runnable TypeScript module that needs no template. */
export function printStandaloneScript(ir: DeckIr, options: PrintStandaloneScriptOptions = {}): PrintedScript {
	const { outputPath, assetDir, assetMode, packageName } = resolvePrintOptions(options)

	const collector = new NoteCollector()
	const assetNames = assetIdentifiers(ir)
	const assets = assetPrinter(assetNames)
	const printAsset = assets.print

	// EVERY part that records a note is walked before the header is built: this tier's own notes
	// are only known afterwards, and `notes` below is a snapshot of the collector rather than a
	// live view. `printDocProps` used to run after that snapshot was taken, so the one note it
	// records reached neither the header nor the returned list — it was collected and dropped.
	const masters = printMasters(ir, collector, printAsset)
	const body = printSlides(ir, collector, printAsset)
	const props = printDocProps(ir, collector)
	// Printed up front too, for its asset references rather than its notes: the header counts, and
	// the bindings declare, only the assets something printed.
	const theme = Object.keys(ir.chrome.theme).length > 0 ? printValue(ir.chrome.theme as IrValue, 0, printAsset) : null
	const notes = [...ir.fidelity, ...collector.notes].filter((note) => noteAppliesTo(note.construct, 'standalone'))

	const needsReadFile = assetMode === 'file' && assets.printed.size > 0
	const lines: string[] = [
		header(
			[
				'Generated from a .pptx by ts-pptx/script.',
				'',
				`Standalone: this script needs nothing but ${packageName}${
					needsReadFile ? ` and ${assets.printed.size} media file(s) in ${assetDir}` : ''
				}.`,
				"The deck's theme, layouts and slide content are all re-authored through the public write",
				'API, so every one of them is editable here — and the parts of the original design the read',
				'model cannot see are gone. They are listed below. For a byte-identical design at the cost',
				'of shipping the source deck alongside the script, print the template-anchored variant',
				'instead.',
				'',
				`Writes ${outputPath}. Needs an ESM context — it uses top-level await.`,
			],
			notes
		),
		'',
		...printPreamble(packageName, { fs: needsReadFile ? ['readFile'] : [], read: false }),
		'',
		...printLayoutSetup(ir.slideSize),
	]

	if (theme !== null) {
		lines.push(
			'',
			"// The theme's colour scheme and font faces. Its format scheme — the fill, line and effect",
			'// style lists a shape references through p:style — has no counterpart on either side of the',
			"// library, so the output carries Office's.",
			`pptx.theme = ${theme}`
		)
	}

	if (props.length > 0) lines.push('', ...props)

	const assetLines = printAssetBindings(
		ir,
		assetNames,
		assetDir,
		assetMode,
		assets.printed,
		'// Media bytes, inlined so this script is a single self-contained file.'
	)
	if (assetLines.length > 0) lines.push('', ...assetLines)

	if (masters.length > 0) lines.push('', ...masters)
	lines.push(...body, '', `await pptx.writeFile({ fileName: here(${printString(outputPath)}) })`, '')

	return printedScript(lines, ir, assetMode, notes, assets.printed)
}

/**
 * The document properties the write API can set.
 *
 * Four of the twelve the read model exposes, which is the loss `deck.docProps` already
 * declares — and unlike the template-anchored tier, here it is a real one, because nothing
 * else carries them.
 */
/**
 * `a`, `a and b`, `a, b and c` — a note is a sentence a human reads, and a bare
 * `join(', ')` made a two-item one read as a list that had lost its last member.
 */
function andList(items: readonly string[]): string {
	if (items.length < 2) return items.join('')
	return `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`
}

function printDocProps(ir: DeckIr, collector: NoteCollector): string[] {
	const entries = Object.entries(ir.props).filter(([, value]) => typeof value === 'string')

	// A deck built through the write API is stamped with the library's own author, subject,
	// title, revision and company in the constructor, and a property cannot be unset — writing
	// `''` emits an empty element rather than removing it. So whatever the source left blank
	// comes back filled in, which is a change to the deck rather than a formatting detail.
	//
	// `company` sat outside this list while nothing populated it in the IR: it was
	// unconditionally "missing", so every standalone script carried a note claiming the source
	// deck declared no company — a claim the converter had never checked and one that is false
	// for most real decks. Now that `Presentation.appProperties` reads `docProps/app.xml`, the
	// key is populated when the source states one and the note is only raised when it does not.
	const stamped = ['title', 'author', 'subject', 'revision', 'company'].filter((key) => !(key in ir.props))
	if (stamped.length > 0) {
		scopeNotes(collector, null).note(
			'deck.docPropsDefault',
			'approximated',
			'unwritable',
			`the source deck declares no ${andList(stamped)}, and a deck built through the write API is stamped with the library's own value for each in its constructor with no way to unset it, so the output declares ${stamped.length === 1 ? 'one' : 'these'} the source did not`
		)
	}

	if (entries.length === 0) return []
	return [
		'// Document properties. Only these have write-API setters; see the fidelity notes.',
		...entries.map(([key, value]) => `pptx.${key} = ${printString(value as string)}`),
	]
}

/**
 * One `defineSlideMaster` per source layout.
 *
 * The naming is the write API's rather than OOXML's: a `defineSlideMaster` call creates a
 * *layout* under the single shared master, which is the right granularity for a source layout.
 * It carries the layout's name, its background, and its decoration — the bands, rules and
 * wordmarks that make the gallery recognisable rather than a list of blank rectangles. What it
 * still cannot carry is the layout's *placeholder definitions*, and that is a write-path
 * consequence rather than a reading one (see `from-read/chrome.ts`). The name earns its place
 * twice over: it is also what `addSlide({ masterTitle })` binds on.
 */
function printMasters(ir: DeckIr, collector: NoteCollector, printAsset: AssetPrinter): string[] {
	// A deck written by this library opens its gallery with the layout the constructor seeds, and
	// the output's constructor seeds it again. Re-authoring it gave the output two layouts named
	// `DEFAULT`, so `addSlide({ masterTitle: 'DEFAULT' })` was ambiguous there; the slides bound to
	// it bind to the seed instead, which is the same layout.
	const seeded = ir.chrome.masters.find(isSeedLayout)

	// Otherwise the extra layout is there whatever is printed below, so the note is recorded before
	// the early return. Scoped to that layout's own title rather than deck-wide, because the round
	// trip matches a note's `shapeName` against the identity of what differed, and an unscoped note
	// would declare *any* extra layout. That title is `DEFAULT` unless a source layout that is not
	// the seed already carries it; the reader then makes the second unique the way it made the
	// source's titles unique, so the extra one is reported under that spelling.
	if (!seeded) {
		const sourceTitles = new Set(ir.chrome.masters.map((master) => master.props['title']).filter((t): t is string => typeof t === 'string')) // prettier-ignore
		scopeNotes(collector, null, uniqueTitle('DEFAULT', sourceTitles)).note(
			'master.default',
			'approximated',
			'unsupported',
			"the write path seeds every presentation with a blank layout of its own named DEFAULT, and there is no way to remove it, so the output deck's layout gallery carries one extra entry ahead of the source's layouts; nothing binds to it, but it is visible in PowerPoint's layout picker"
		)
	}
	const authored = ir.chrome.masters.filter((master) => master !== seeded)
	if (authored.length === 0) return []
	return [
		'// One master per source layout, carrying its name, background and decoration. What a',
		'// layout does not bring across is its placeholder definitions; see the fidelity notes.',
		...authored.map((master) => printArguments('pptx.defineSlideMaster', [master.props], 0, printAsset)),
	]
}

/**
 * What the layout the write path seeds every presentation with reads back as: blank, titled
 * `DEFAULT`, on the white its master's `p:bgRef` resolves to. `compact` sorts an IR object's keys,
 * so the serialized form is stable.
 */
const SEED_LAYOUT_PROPS = JSON.stringify({ background: { color: 'FFFFFF' }, title: 'DEFAULT' })

/**
 * Whether a source layout is the one the output's constructor seeds: first in the gallery, where
 * the seed sits, and identical to it. Anything else is re-authored, even a layout named `DEFAULT`,
 * because binding its slides to the seed would lose what sets it apart.
 */
function isSeedLayout(master: MasterIr): boolean {
	return master.layoutIndex === 0 && JSON.stringify(master.props) === SEED_LAYOUT_PROPS
}

/** Every slide, in source order, each bound to the master its source layout became. */
function printSlides(ir: DeckIr, collector: NoteCollector, printAsset: AssetPrinter): string[] {
	const byLayout = new Map<number, MasterIr>()
	for (const master of ir.chrome.masters) byLayout.set(master.layoutIndex, master)

	const lines: string[] = []
	for (const slide of ir.slides) {
		const master = slide.layout === null ? undefined : byLayout.get(slide.layout.index)
		const title = typeof master?.props['title'] === 'string' ? master.props['title'] : null

		if (title === null) {
			scopeNotes(collector, slide.number).note(
				'slide.layout',
				'approximated',
				'unsupported',
				slide.layout === null
					? "this slide resolves no layout of its own, so it binds to the write path's default blank layout"
					: `this slide's source layout (gallery position ${slide.layout.index}) produced no master to bind to, so it binds to the write path's default blank layout`
			)
		}

		lines.push(
			...printSlide(
				slide,
				{
					comment: `Slide ${slide.number} — ${title === null ? 'no master' : `master ${JSON.stringify(title)}`}`,
					construction: title === null ? 'pptx.addSlide()' : `pptx.addSlide({ masterTitle: ${printString(title)} })`,
				},
				collector,
				printAsset
			)
		)
	}
	return lines
}
