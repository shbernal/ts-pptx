/**
 * {@link DeckIr} → a runnable TypeScript module, template-anchored. **Tier B.**
 *
 * **The tier.** The emitted script reuses the *source deck itself* as its template:
 * `Presentation.fromTemplate` strips a package's slides while leaving its masters, layouts,
 * theme, and document properties byte-identical, so the chrome is never regenerated — it is
 * the original. Only the slides are rebuilt, through the public write API, and grafted back
 * with `appendSlides`. That is why this tier's fidelity note list is short: everything the
 * converter cannot see or cannot write is confined to slide content, and the whole of the
 * deck's design survives because nothing ever tried to reproduce it.
 *
 * It also means the plan's "strip the source slides" step needs no helper. `fromTemplate`
 * already does exactly that, and does it without pruning shared parts, so the template
 * asset and the source deck are the same file.
 *
 * **Order is not rearranged.** Slides are emitted in source order and every operation
 * appends, so `p:sldIdLst` comes out in the original order without any position arithmetic.
 * Contiguous slides sharing a layout share one generator, because `appendSlides` binds
 * every slide in one call to a single layout — the batching is forced by that signature,
 * not chosen for tidiness.
 *
 * **This module may add fidelity notes.** A loss can belong to the *tier* rather than to
 * the conversion — a slide's `p:cSld@name` is read fine and would survive a byte copy, but
 * has no public write-API setter, so it dies here and nowhere else. {@link PrintedScript}
 * therefore returns the notes that actually apply to *this* output, which is also the set a
 * round-trip check should exclude from its diff.
 *
 * The standalone counterpart is `printStandaloneScript` (Tier A), which prints the same IR
 * with no template at all and pays for it in the note list.
 */
import type { DeckIr, SlideIr } from '../ir.js'
import { NoteCollector, noteAppliesTo, scopeNotes } from '../fidelity.js'
import { commentText, printString, type AssetPrinter } from './literal.js'
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

export type { AssetMode, PrintedScript } from './common.js'

export interface PrintScriptOptions extends CommonPrintOptions {
	/**
	 * Path the emitted script loads its template from, resolved against the script's own
	 * location. This is the **source deck unchanged** — `fromTemplate` strips its slides.
	 * @default './template.pptx'
	 */
	templatePath?: string
}

/** Turn a deck IR into a runnable, template-anchored TypeScript module. */
export function printScript(ir: DeckIr, options: PrintScriptOptions = {}): PrintedScript {
	const templatePath = options.templatePath ?? './template.pptx'
	const { outputPath, assetDir, assetMode, packageName } = resolvePrintOptions(options)

	const collector = new NoteCollector()
	const assetNames = assetIdentifiers(ir)
	const assets = assetPrinter(assetNames)
	const printAsset = assets.print

	const body = printSlides(ir, collector, printAsset)
	const needsSource = ir.slides.some((slide) => slide.source === 'carried')
	const needsFallbackLayout = ir.slides.some((slide) => slide.source === 'authored' && slide.layout === null)

	// Built last: the tier's own notes are only known once the slides have been walked.
	// Each construct's catalogue entry says which tiers it applies to; the chrome's losses, for one,
	// are carried by the template rather than lost here.
	const notes = [...ir.fidelity, ...collector.notes].filter((note) => noteAppliesTo(note.construct, 'template'))

	const lines: string[] = [
		header(
			[
				'Generated from a .pptx by ts-pptx/script.',
				'',
				'Template-anchored: the source deck is reused as the template, so its slide masters,',
				'layouts, theme and document properties are the originals, byte for byte. Only the',
				'slide content below was rebuilt through the public write API.',
				'',
				`Expects ${templatePath} (the source deck, unmodified) beside this file${
					// Counted after the slides were printed: only what they reference ships. The chrome's
					// pictures stay in the template, and a carried slide is copied with its own.
					assetMode === 'file' && assets.printed.size > 0
						? `, plus ${assets.printed.size} media file(s) in ${assetDir}`
						: ''
				}.`,
				`Writes ${outputPath}. Needs an ESM context — it uses top-level await.`,
			],
			notes
		),
		'',
		...printPreamble(packageName, { fs: ['readFile', 'writeFile'], read: true }),
		'',
		'// The source deck is the template: fromTemplate strips its slides and leaves its masters,',
		'// layouts, theme and document properties exactly as they were.',
		`const deck = await Presentation.fromTemplate(here(${printString(templatePath)}))`,
	]

	if (needsSource) {
		lines.push(
			'',
			'// A second, unstripped handle on the same file, for the slides copied verbatim.',
			`const source = await Presentation.load(here(${printString(templatePath)}))`
		)
	}

	if (needsFallbackLayout) {
		lines.push(
			'',
			'// Some source slides resolved no layout of their own; they bind to the first one.',
			'const fallbackLayout = deck.layouts()[0]',
			"if (!fallbackLayout) throw new Error('the template declares no slide layouts to bind to')"
		)
	}

	const assetLines = printAssetBindings(ir, assetNames, assetDir, assetMode, assets.printed)
	if (assetLines.length > 0) lines.push('', ...assetLines)

	lines.push(
		'',
		'/**',
		' * A generator sized to the template. appendSlides compares slide sizes exactly and',
		' * throws when they differ, so the inches here must round-trip to the source EMU.',
		' */',
		'function generator(): TsPptx {',
		...printLayoutSetup(ir.slideSize, '\t'),
		'\treturn pptx',
		'}',
		...body,
		'',
		`await writeFile(here(${printString(outputPath)}), await deck.save())`,
		''
	)

	return printedScript(lines, ir, assetMode, notes, assets.printed)
}

/**
 * The slide statements, batched into generators, preceded by the layout handles they bind to.
 *
 * A batch ends when the layout changes or a carried slide interrupts, because
 * `appendSlides` binds one layout per call and a carried slide is not authored at all.
 *
 * Every layout `const` is hoisted above the first statement rather than emitted where it is
 * first used. `deck.layouts()` is a snapshot of the gallery, and `importSlide` — the call a
 * carried slide prints — grows that gallery with the source slide's own layout and master.
 * Resolving a handle after one has run is therefore resolving against a deck that has already
 * moved; capturing every handle up front is the only ordering that cannot.
 */
function printSlides(ir: DeckIr, collector: NoteCollector, printAsset: AssetPrinter): string[] {
	const declarations: string[] = []
	const lines: string[] = []
	let batch: SlideIr[] = []
	let generatorIndex = 0
	// Two non-contiguous batches can share one layout, so its `const` must be emitted once.
	const declaredLayouts = new Set<number>()

	const flush = (): void => {
		if (batch.length === 0) return
		const first = batch[0]
		if (!first) return
		const name = `gen${++generatorIndex}`
		lines.push('', `const ${name} = generator()`)
		for (const slide of batch) lines.push(...printAuthoredSlide(slide, name, collector, printAsset))
		lines.push('', printAppend(name, first, collector, declaredLayouts, declarations))
		batch = []
	}

	for (const slide of ir.slides) {
		if (slide.source === 'carried') {
			flush()
			lines.push(
				'',
				`// Slide ${slide.number} — copied from the source deck rather than rebuilt; see the fidelity notes.`,
				`deck.importSlide(source, ${slide.number - 1})`
			)
			continue
		}
		// A batch is one appendSlides call, which binds every slide in it to one layout.
		const previous = batch[0]
		if (previous && previous.layout?.index !== slide.layout?.index) flush()
		batch.push(slide)
	}
	flush()

	return declarations.length > 0 ? ['', ...declarations, ...lines] : lines
}

/**
 * The `appendSlides` call closing a batch, binding by layout name where that is
 * unambiguous and by gallery position where it is not. Any handle it needs is appended to
 * `declarations`, which the caller hoists above every statement.
 *
 * Binding by name is preferred because it survives being re-pointed at a different
 * template and because it is legible. But `appendSlides` *throws* on an ambiguous name
 * rather than picking one, so the positional form resolves the case where the name does not
 * identify one layout: a multi-master deck that repeats layout names.
 *
 * A carried slide used to demote every binding in the script to a position too, because its
 * `importSlide` copied the source layout in under a second name-alike entry. `importSlide`
 * now binds to the chrome the template already holds (`read/api/ops/part-reuse.ts`), and
 * this deck's template *is* the source file, so the gallery no longer grows and the names
 * stay unambiguous.
 */
function printAppend(
	generatorName: string,
	first: SlideIr,
	collector: NoteCollector,
	declared: Set<number>,
	declarations: string[]
): string {
	const layout = first.layout
	if (!layout) return `await deck.appendSlides(${generatorName}, { layout: fallbackLayout })`
	if (layout.nameIsUnique) {
		return `await deck.appendSlides(${generatorName}, { layout: ${printString(layout.name)} })`
	}

	scopeNotes(collector, first.number).note(
		'slide.layout',
		'approximated',
		'unsupported',
		`more than one layout in the source deck is named ${JSON.stringify(layout.name)}, so this slide binds to gallery position ${layout.index} instead; re-pointing the script at a different template will not track the name`
	)
	const identifier = `layout${layout.index}`
	if (!declared.has(layout.index)) {
		declared.add(layout.index)
		declarations.push(
			`const ${identifier} = deck.layouts()[${layout.index}] // ${commentText(JSON.stringify(layout.name))}`,
			`if (!${identifier}) throw new Error('the template has no layout at position ${layout.index}')`
		)
	}
	return `await deck.appendSlides(${generatorName}, { layout: ${identifier} })`
}

/** One authored slide, plus the layout-binding note this tier is responsible for. */
function printAuthoredSlide(
	slide: SlideIr,
	generatorName: string,
	collector: NoteCollector,
	printAsset: AssetPrinter
): string[] {
	if (slide.layout === null) {
		scopeNotes(collector, slide.number).note(
			'slide.layout',
			'approximated',
			'unsupported',
			"this slide resolves no layout of its own, so it binds to the template's first layout; that governs theme and clrMap resolution, which may differ from the source"
		)
	}

	// A tier loss, and a narrow one: a transition's *stop-previous* form needs no relationship
	// and rides across fine, while an embedded start sound needs an audio part wired to the
	// appended slide. `extractSlides` never runs the registration pass that assigns it a
	// relationship id (`registerTransitionSounds` belongs to the package-assembly path), so the
	// emitter finds no `_sndRId` and writes no `p:sndAc` at all. That is a silent drop rather
	// than a dangling reference — which is the safe failure of the two, and still a loss the
	// reader has to be told about. The standalone tier writes a real package and keeps it.
	if (slide.transition?.sound?.data !== undefined) {
		scopeNotes(collector, slide.number).note(
			'slide.transitionSound',
			'dropped',
			'unsupported',
			"this transition's start sound is dropped: the append path this tier rides does not register a transition's embedded audio part, so the sound does not reach the template. The assignment below still spells it out, and the WAV still ships beside the script, because both describe the source deck faithfully and would take effect the day the append path carries them — but this output is silent. Print the standalone variant to keep the sound"
		)
	}

	const layout = slide.layout === null ? 'no layout' : `layout ${JSON.stringify(slide.layout.name)}`
	return printSlide(
		slide,
		{ comment: `Slide ${slide.number} — ${layout}`, construction: `${generatorName}.addSlide()` },
		collector,
		printAsset
	)
}
