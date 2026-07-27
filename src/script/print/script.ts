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
import { NoteCollector, scopeNotes } from '../fidelity.js'
import { inches } from '../from-read/values.js'
import { printString, type AssetPrinter } from './literal.js'
import {
	assetIdentifiers,
	assetPrinter,
	header,
	printAssetBindings,
	printSlide,
	type AssetMode,
	type PrintedScript,
} from './common.js'

export type { AssetMode, PrintedScript } from './common.js'

export interface PrintScriptOptions {
	/**
	 * Path the emitted script loads its template from, resolved against the script's own
	 * location. This is the **source deck unchanged** — `fromTemplate` strips its slides.
	 * @default './template.pptx'
	 */
	templatePath?: string
	/** Path the emitted script writes to, resolved against the script's own location. @default './output.pptx' */
	outputPath?: string
	/** Directory the emitted script reads image assets from, when {@link assets} is `'file'`. @default './assets' */
	assetDir?: string
	/** @default 'file' */
	assets?: AssetMode
	/**
	 * Import specifier the emitted script uses, with `/read` appended for the read half.
	 * Defaults to this package's own published name; override it to point a generated
	 * script at a local build or a fork.
	 * @default '@shbernal/ts-pptx'
	 */
	packageName?: string
}

/**
 * Constructs the IR declares lost that a template-anchored output keeps anyway.
 *
 * Document properties are the obvious case: the read half notes that only five of the twelve
 * have write-API setters, which is true and matters for a standalone output — but this tier
 * never authors them at all. They ride in the template, untouched, all twelve of them.
 *
 * The whole of the chrome is the same story on a larger scale. The theme's format scheme, the
 * master's text styles and colour map, the layouts' decoration and placeholder definitions —
 * every one of them is a genuine loss for a script that has to rebuild the deck's design, and
 * none of them is touched here, because the design *is* the template. A caveat that does not
 * apply to the output in front of you is worse than no caveat: it teaches the reader to skim
 * the ones that do.
 */
const TEMPLATE_CARRIED_CONSTRUCTS = new Set([
	'deck.docProps',
	'master.background',
	'master.colorMap',
	'master.decoration',
	'master.multiple',
	// Both are about a layout *title* the standalone tier has to invent — deduplicated because
	// it doubles as a lookup key, whitespace-collapsed because the write path emits it as a raw
	// XML attribute value. Here the layout keeps its own `p:cSld@name`, untouched.
	'master.name',
	'master.nameCollision',
	'master.placeholders',
	'master.txStyles',
	'theme.fmtScheme',
])

/**
 * Default import specifier for the emitted script — this package's *published* name, which
 * is not its directory name. Getting it wrong produces a script that prints and typechecks
 * and then fails at `import`, so it is pinned by a test rather than left to a literal here.
 */
const PACKAGE_NAME = '@shbernal/ts-pptx'

/** Turn a deck IR into a runnable, template-anchored TypeScript module. */
export function printScript(ir: DeckIr, options: PrintScriptOptions = {}): PrintedScript {
	const templatePath = options.templatePath ?? './template.pptx'
	const outputPath = options.outputPath ?? './output.pptx'
	const assetDir = (options.assetDir ?? './assets').replace(/\/$/, '')
	const assetMode = options.assets ?? 'file'
	const packageName = options.packageName ?? PACKAGE_NAME

	const collector = new NoteCollector()
	const assetNames = assetIdentifiers(ir)
	const printAsset = assetPrinter(assetNames)

	const body = printSlides(ir, collector, printAsset)
	const needsSource = ir.slides.some((slide) => slide.source === 'carried')
	const needsFallbackLayout = ir.slides.some((slide) => slide.source === 'authored' && slide.layout === null)

	// Built last: the tier's own notes are only known once the slides have been walked.
	const notes = [...ir.fidelity.filter((note) => !TEMPLATE_CARRIED_CONSTRUCTS.has(note.construct)), ...collector.notes]

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
					assetMode === 'file' && ir.assets.length > 0 ? `, plus ${ir.assets.length} media file(s) in ${assetDir}` : ''
				}.`,
				`Writes ${outputPath}. Needs an ESM context — it uses top-level await.`,
			],
			notes
		),
		'',
		"import { readFile, writeFile } from 'node:fs/promises'",
		"import { fileURLToPath } from 'node:url'",
		`import TsPptx from ${printString(packageName)}`,
		`import { Presentation } from ${printString(`${packageName}/read`)}`,
		'',
		'/** Resolve a path against this script rather than against the working directory. */',
		'const here = (name: string): string => fileURLToPath(new URL(name, import.meta.url))',
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

	const assetLines = printAssetBindings(ir, assetNames, assetDir, assetMode)
	if (assetLines.length > 0) lines.push('', ...assetLines)

	lines.push(
		'',
		'/**',
		' * A generator sized to the template. appendSlides compares slide sizes exactly and',
		' * throws when they differ, so the inches here must round-trip to the source EMU.',
		' */',
		'function generator(): TsPptx {',
		'\tconst pptx = new TsPptx()',
		`\tpptx.defineLayout({ name: 'source', width: ${inches(ir.slideSize.widthEmu)}, height: ${inches(ir.slideSize.heightEmu)} })`,
		"\tpptx.layout = 'source'",
		'\treturn pptx',
		'}',
		...body,
		'',
		`await writeFile(here(${printString(outputPath)}), await deck.save())`,
		''
	)

	return {
		code: lines.join('\n'),
		assets: new Map(
			assetMode === 'file' ? ir.assets.map((asset): [string, Uint8Array] => [asset.name, asset.bytes]) : []
		),
		notes,
	}
}

/**
 * The slide statements, batched into generators.
 *
 * A batch ends when the layout changes or a carried slide interrupts, because
 * `appendSlides` binds one layout per call and a carried slide is not authored at all.
 */
function printSlides(ir: DeckIr, collector: NoteCollector, printAsset: AssetPrinter): string[] {
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
		lines.push('', ...printAppend(name, first, collector, declaredLayouts))
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

	return lines
}

/**
 * The `appendSlides` call closing a batch, binding by layout name where that is
 * unambiguous and by gallery position where it is not.
 *
 * Binding by name is preferred because it survives being re-pointed at a different
 * template and because it is legible. But `appendSlides` *throws* on an ambiguous name
 * rather than picking one, and a multi-master deck routinely repeats layout names, so the
 * positional form is the only thing that resolves there.
 */
function printAppend(generatorName: string, first: SlideIr, collector: NoteCollector, declared: Set<number>): string[] {
	const layout = first.layout
	if (!layout) return [`await deck.appendSlides(${generatorName}, { layout: fallbackLayout })`]
	if (layout.nameIsUnique) {
		return [`await deck.appendSlides(${generatorName}, { layout: ${printString(layout.name)} })`]
	}

	scopeNotes(collector, first.number).note(
		'slide.layout',
		'approximated',
		'unsupported',
		`more than one layout in the source deck is named ${JSON.stringify(layout.name)}, so this slide binds to gallery position ${layout.index} instead; re-pointing the script at a different template will not track the name`
	)
	const identifier = `layout${layout.index}`
	const append = `await deck.appendSlides(${generatorName}, { layout: ${identifier} })`
	if (declared.has(layout.index)) return [append]
	declared.add(layout.index)
	return [
		`const ${identifier} = deck.layouts()[${layout.index}] // ${JSON.stringify(layout.name)}`,
		`if (!${identifier}) throw new Error('the template has no layout at position ${layout.index}')`,
		append,
	]
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
