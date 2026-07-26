/**
 * {@link DeckIr} → a runnable TypeScript module, template-anchored.
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
 */
import type { DeckIr, IrValue, SlideIr } from '../ir.js'
import type { FidelityNote } from '../fidelity.js'
import { NoteCollector, scopeNotes } from '../fidelity.js'
import { inches } from '../from-read/values.js'
import { printArguments, printString, printValue, type AssetPrinter } from './literal.js'

/** How image bytes reach the emitted script. */
export type AssetMode =
	/** Written beside the script and read at run time. Keeps the script readable. */
	| 'file'
	/** Inlined as `data:` literals. One self-contained file, at roughly 4/3 the byte size. */
	| 'inline'

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

export interface PrintedScript {
	/** The TypeScript module source. */
	code: string
	/**
	 * Image bytes the script expects to find in its asset directory, keyed by filename.
	 * Empty when {@link PrintScriptOptions.assets} is `'inline'`, since the bytes are then
	 * in {@link code}.
	 */
	assets: Map<string, Uint8Array>
	/**
	 * The losses that apply to **this** output: the IR's own notes, minus any the template
	 * anchoring rescues, plus the ones this tier causes. Reproduced as a comment block at
	 * the top of {@link code}, so the artifact carries its own caveats.
	 */
	notes: FidelityNote[]
}

/**
 * Constructs the IR declares lost that a template-anchored output keeps anyway.
 *
 * Document properties are the case: the read half notes that only five of the twelve have
 * write-API setters, which is true and matters for a standalone output — but this tier
 * never authors them at all. They ride in the template, untouched, all twelve of them. A
 * caveat that does not apply to the output in front of you is worse than no caveat, because
 * it teaches the reader to skim the ones that do.
 */
const TEMPLATE_CARRIED_CONSTRUCTS = new Set(['deck.docProps'])

/**
 * Default import specifier for the emitted script — this package's *published* name, which
 * is not its directory name. Getting it wrong produces a script that prints and typechecks
 * and then fails at `import`, so it is pinned by a test rather than left to a literal here.
 */
const PACKAGE_NAME = '@shbernal/ts-pptx'

/** Width the header comment's prose wraps at, leaving room for the ` * ` prefix. */
const COMMENT_WIDTH = 110

/** Turn a deck IR into a runnable, template-anchored TypeScript module. */
export function printScript(ir: DeckIr, options: PrintScriptOptions = {}): PrintedScript {
	const templatePath = options.templatePath ?? './template.pptx'
	const outputPath = options.outputPath ?? './output.pptx'
	const assetDir = (options.assetDir ?? './assets').replace(/\/$/, '')
	const assetMode = options.assets ?? 'file'
	const packageName = options.packageName ?? PACKAGE_NAME

	const collector = new NoteCollector()
	const assetNames = assetIdentifiers(ir)
	const printAsset: AssetPrinter = (ref) => {
		const identifier = assetNames.get(ref.$asset)
		if (!identifier) throw new Error(`Asset reference ${ref.$asset} has no entry in DeckIr.assets`)
		return identifier
	}

	const body = printSlides(ir, collector, printAsset)
	const needsSource = ir.slides.some((slide) => slide.source === 'carried')
	const needsFallbackLayout = ir.slides.some((slide) => slide.source === 'authored' && slide.layout === null)

	// Built last: the tier's own notes are only known once the slides have been walked.
	const notes = [...ir.fidelity.filter((note) => !TEMPLATE_CARRIED_CONSTRUCTS.has(note.construct)), ...collector.notes]

	const lines: string[] = [
		header(ir, notes, { templatePath, outputPath, assetDir, assetMode }),
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
 * A stable JavaScript identifier per asset.
 *
 * Derived from the asset's own name rather than its index, so a reader can match a `const`
 * in the script against a file in the asset directory. The dedupe suffix is defensive: the
 * read half currently names assets `image1.png`, `image2.jpg`, … which cannot collide after
 * the extension is stripped, but nothing in the IR contract promises that.
 */
function assetIdentifiers(ir: DeckIr): Map<string, string> {
	const out = new Map<string, string>()
	const used = new Set<string>()
	for (const asset of ir.assets) {
		const base = asset.name.replace(/\.[^.]*$/, '').replace(/[^A-Za-z0-9_$]/g, '_')
		let identifier = /^[A-Za-z_$]/.test(base) ? base : `_${base}`
		let suffix = 2
		while (used.has(identifier)) identifier = `${base}_${suffix++}`
		used.add(identifier)
		out.set(asset.name, identifier)
	}
	return out
}

/**
 * One `const` per asset, each holding a complete `data:` URI.
 *
 * Both modes bind the same *kind* of value — a data URI string — so the call sites are
 * identical and `addImage`'s `data` option is used either way. Only where the bytes come
 * from changes. Rewriting the option key to `path` in file mode would have the printer
 * reshaping an argument, which is precisely the thing that keeps the IR checkable.
 */
function printAssetBindings(ir: DeckIr, identifiers: Map<string, string>, assetDir: string, mode: AssetMode): string[] {
	if (ir.assets.length === 0) return []

	const lines =
		mode === 'file'
			? ['// Image bytes, loaded from the asset directory beside this script.']
			: ['// Image bytes, inlined so this script needs no files but its template.']

	for (const asset of ir.assets) {
		const identifier = identifiers.get(asset.name)
		if (!identifier) continue
		if (mode === 'file') {
			const path = printString(`${assetDir}/${asset.name}`)
			lines.push(
				`const ${identifier} = \`data:${asset.contentType};base64,\${(await readFile(here(${path}))).toString('base64')}\``
			)
		} else {
			lines.push(`const ${identifier} = ${printString(`data:${asset.contentType};base64,${toBase64(asset.bytes)}`)}`)
		}
	}
	return lines
}

/** Base64 without `Buffer`, matching the rest of the library's isomorphic media handling. */
function toBase64(bytes: Uint8Array): string {
	let binary = ''
	// Chunked because String.fromCharCode is applied to the whole slice at once, and a
	// megabyte-sized spread overflows the argument limit.
	const CHUNK = 0x8000
	for (let i = 0; i < bytes.length; i += CHUNK) {
		binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
	}
	return btoa(binary)
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

/** One authored slide: its slide-level properties, then its calls in z-order. */
function printAuthoredSlide(
	slide: SlideIr,
	generatorName: string,
	collector: NoteCollector,
	printAsset: AssetPrinter
): string[] {
	const notes = scopeNotes(collector, slide.number)
	const identifier = `slide${slide.number}`
	const layout = slide.layout === null ? 'no layout' : `layout ${JSON.stringify(slide.layout.name)}`
	const lines = ['', `// Slide ${slide.number} — ${layout}`, `const ${identifier} = ${generatorName}.addSlide()`]

	if (slide.layout === null) {
		notes.note(
			'slide.layout',
			'approximated',
			'unsupported',
			"this slide resolves no layout of its own, so it binds to the template's first layout; that governs theme and clrMap resolution, which may differ from the source"
		)
	}
	if (slide.name !== undefined) {
		notes.note(
			'slide.name',
			'dropped',
			'unwritable',
			`the slide's own name (p:cSld@name, here ${JSON.stringify(slide.name)}) has no public write-API setter, so the appended slide is unnamed`
		)
	}

	if (slide.hidden) lines.push(`${identifier}.hidden = true`)
	if (slide.background) {
		lines.push(`${identifier}.background = ${printValue(slide.background as IrValue, 0, printAsset)}`)
	}
	if (slide.notesText !== undefined) {
		lines.push(printArguments(`${identifier}.addNotes`, [slide.notesText], 0, printAsset))
	}

	for (const call of slide.calls) {
		// The source shape name makes a fidelity note navigable: a note carries the same
		// name, and without it a reader has no way to find the call it refers to.
		if (call.sourceName) lines.push(`// ${call.sourceName.replace(/\s+/g, ' ')}`)
		lines.push(printArguments(`${identifier}.${call.method}`, call.args, 0, printAsset))
	}

	return lines
}

/** The banner comment: how to run the script, and every loss that applies to it. */
function header(
	ir: DeckIr,
	notes: FidelityNote[],
	paths: { templatePath: string; outputPath: string; assetDir: string; assetMode: AssetMode }
): string {
	const lines = [
		'Generated from a .pptx by ts-pptx/script.',
		'',
		'Template-anchored: the source deck is reused as the template, so its slide masters,',
		'layouts, theme and document properties are the originals, byte for byte. Only the',
		'slide content below was rebuilt through the public write API.',
		'',
		`Expects ${paths.templatePath} (the source deck, unmodified) beside this file${
			paths.assetMode === 'file' && ir.assets.length > 0
				? `, plus ${ir.assets.length} image(s) in ${paths.assetDir}`
				: ''
		}.`,
		`Writes ${paths.outputPath}. Needs an ESM context — it uses top-level await.`,
		'',
	]

	if (notes.length === 0) {
		lines.push('FIDELITY: nothing was lost that this converter can detect.')
	} else {
		lines.push(
			`FIDELITY: ${notes.length} declared loss(es). Everything not listed here is expected to`,
			'survive unchanged; a difference that is not in this list is a defect, not a caveat.',
			''
		)
		for (const [scope, group] of groupNotes(notes)) {
			lines.push(`${scope}:`)
			for (const note of group) {
				const where = note.shapeName ? ` (${note.shapeName})` : ''
				lines.push(
					...wrap(
						`- ${note.construct}${where} [${note.disposition}, ${note.cause}] — ${note.detail}`,
						COMMENT_WIDTH,
						'  '
					)
				)
			}
		}
	}

	return ['/*', ...lines.map((line) => (line ? ` * ${line}` : ' *')), ' */'].join('\n')
}

/** Notes bucketed by slide, in deck order, so a reader meets them the way they meet the deck. */
function groupNotes(notes: FidelityNote[]): Array<[string, FidelityNote[]]> {
	const deckLevel = notes.filter((note) => note.slideNumber === null)
	const bySlide = new Map<number, FidelityNote[]>()
	for (const note of notes) {
		if (note.slideNumber === null) continue
		const existing = bySlide.get(note.slideNumber)
		if (existing) existing.push(note)
		else bySlide.set(note.slideNumber, [note])
	}

	const out: Array<[string, FidelityNote[]]> = []
	if (deckLevel.length > 0) out.push(['Deck', deckLevel])
	for (const number of [...bySlide.keys()].sort((a, b) => a - b)) {
		out.push([`Slide ${number}`, bySlide.get(number) ?? []])
	}
	return out
}

/** Wrap prose to `width`, indenting every line after the first by `continuation`. */
function wrap(text: string, width: number, continuation: string): string[] {
	const words = text.split(/\s+/).filter(Boolean)
	const lines: string[] = []
	let current = ''
	for (const word of words) {
		const candidate = current ? `${current} ${word}` : `${continuation}${word}`
		if (current && candidate.length > width) {
			lines.push(current)
			current = `${continuation}${word}`
		} else {
			current = current ? candidate : word
		}
	}
	if (current) lines.push(current)
	return lines
}
