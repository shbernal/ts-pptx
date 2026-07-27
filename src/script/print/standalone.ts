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
 * **The ceiling is measured, not estimated, and no amount of printer work moves it.** Three
 * constructs are unreachable from *both* directions: `a:fmtScheme` (nothing reads it, and the
 * write path emits a hardcoded Office one), `p:txStyles` (no reader, though
 * `SlideMasterProps.textStyles` could author it), and master/layout decoration (documented as
 * out of the read model's scope, carried byte-for-byte by the import paths instead). A fourth,
 * `p:clrMap`, is readable and has no setter. Each is a fidelity note here and a `pass` in the
 * other tier, which is the whole reason the other tier shipped first.
 *
 * **One structural difference beyond the chrome.** A slide marked `carried` has no source
 * package to be carried from here, so it is transcribed like any other and loses only the
 * construct that made it uncarryable — which each shape has already declared for itself. That
 * is why {@link DeckIr}'s calls are populated even for a carried slide.
 */
import type { DeckIr, IrValue, MasterIr } from '../ir.js'
import { NoteCollector, scopeNotes } from '../fidelity.js'
import { inches } from '../from-read/values.js'
import { printArguments, printString, printValue, type AssetPrinter } from './literal.js'
import {
	assetIdentifiers,
	assetPrinter,
	header,
	printAssetBindings,
	printSlide,
	type AssetMode,
	type PrintedScript,
} from './common.js'

export interface PrintStandaloneScriptOptions {
	/** Path the emitted script writes to, resolved against the script's own location. @default './output.pptx' */
	outputPath?: string
	/** Directory the emitted script reads image assets from, when {@link assets} is `'file'`. @default './assets' */
	assetDir?: string
	/** @default 'file' */
	assets?: AssetMode
	/**
	 * Import specifier the emitted script uses. Defaults to this package's own published
	 * name; override it to point a generated script at a local build or a fork.
	 * @default '@shbernal/ts-pptx'
	 */
	packageName?: string
}

/**
 * Notes the IR declares that do not describe *this* output.
 *
 * Only one, and it is the mirror of the other tier's list. `slide.carried` says a slide will
 * be copied from the source package rather than transcribed; there is no source package here,
 * so the slide is transcribed and the construct that made it uncarryable has already recorded
 * its own note. Reporting both would double-count one loss and describe a behaviour this
 * script does not have.
 */
const NOT_APPLICABLE = new Set(['slide.carried'])

/** See `printScript` — the published name is not the directory name, and a test pins it. */
const PACKAGE_NAME = '@shbernal/ts-pptx'

/** Turn a deck IR into a runnable TypeScript module that needs no template. */
export function printStandaloneScript(ir: DeckIr, options: PrintStandaloneScriptOptions = {}): PrintedScript {
	const outputPath = options.outputPath ?? './output.pptx'
	const assetDir = (options.assetDir ?? './assets').replace(/\/$/, '')
	const assetMode = options.assets ?? 'file'
	const packageName = options.packageName ?? PACKAGE_NAME

	const collector = new NoteCollector()
	const assetNames = assetIdentifiers(ir)
	const printAsset = assetPrinter(assetNames)

	// Walked before the header is built: this tier's own notes are only known afterwards.
	const masters = printMasters(ir, collector, printAsset)
	const body = printSlides(ir, collector, printAsset)
	const notes = [...ir.fidelity.filter((note) => !NOT_APPLICABLE.has(note.construct)), ...collector.notes]

	const needsReadFile = assetMode === 'file' && ir.assets.length > 0
	const lines: string[] = [
		header(
			[
				'Generated from a .pptx by ts-pptx/script.',
				'',
				`Standalone: this script needs nothing but ${packageName}${
					needsReadFile ? ` and ${ir.assets.length} media file(s) in ${assetDir}` : ''
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
		...(needsReadFile ? ["import { readFile } from 'node:fs/promises'"] : []),
		"import { fileURLToPath } from 'node:url'",
		`import TsPptx from ${printString(packageName)}`,
		'',
		'/** Resolve a path against this script rather than against the working directory. */',
		'const here = (name: string): string => fileURLToPath(new URL(name, import.meta.url))',
		'',
		'const pptx = new TsPptx()',
		`pptx.defineLayout({ name: 'source', width: ${inches(ir.slideSize.widthEmu)}, height: ${inches(ir.slideSize.heightEmu)} })`,
		"pptx.layout = 'source'",
	]

	const theme = ir.chrome.theme as IrValue
	if (Object.keys(ir.chrome.theme).length > 0) {
		lines.push(
			'',
			"// The theme's colour scheme and font faces. Its format scheme — the fill, line and effect",
			'// style lists a shape references through p:style — has no counterpart on either side of the',
			"// library, so the output carries Office's.",
			`pptx.theme = ${printValue(theme, 0, printAsset)}`
		)
	}

	const props = printDocProps(ir, collector)
	if (props.length > 0) lines.push('', ...props)

	const assetLines = printAssetBindings(
		ir,
		assetNames,
		assetDir,
		assetMode,
		'// Media bytes, inlined so this script is a single self-contained file.'
	)
	if (assetLines.length > 0) lines.push('', ...assetLines)

	if (masters.length > 0) lines.push('', ...masters)
	lines.push(...body, '', `await pptx.writeFile({ fileName: here(${printString(outputPath)}) })`, '')

	return {
		code: lines.join('\n'),
		assets: new Map(
			assetMode === 'file' ? ir.assets.map((asset): [string, Uint8Array] => [asset.name, asset.bytes]) : []
		),
		notes,
	}
}

/**
 * The document properties the write API can set.
 *
 * Four of the twelve the read model exposes, which is the loss `deck.docProps` already
 * declares — and unlike the template-anchored tier, here it is a real one, because nothing
 * else carries them.
 */
function printDocProps(ir: DeckIr, collector: NoteCollector): string[] {
	const entries = Object.entries(ir.props).filter(([, value]) => typeof value === 'string')

	// A deck built through the write API is stamped with the library's own author, company,
	// subject, title and revision in the constructor, and a property cannot be unset — writing
	// `''` emits an empty element rather than removing it. So whatever the source left blank
	// comes back filled in, which is a change to the deck rather than a formatting detail.
	const stamped = ['title', 'author', 'company', 'subject', 'revision'].filter((key) => !(key in ir.props))
	if (stamped.length > 0) {
		scopeNotes(collector, null).note(
			'deck.docPropsDefault',
			'approximated',
			'unwritable',
			`the source deck declares no ${stamped.join(', ')}, and a deck built through the write API is stamped with the library's own value for each in its constructor with no way to unset it, so the output declares ${stamped.length === 1 ? 'one' : 'these'} the source did not`
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
 * What it carries is thin by measurement, not by choice — a layout's decoration and its
 * placeholder definitions are both out of reach (see `from-read/chrome.ts`), leaving the name
 * and the background. The name still earns its place: it is what `addSlide({ masterTitle })`
 * binds on, so the deck keeps a layout gallery a reader recognises.
 */
function printMasters(ir: DeckIr, collector: NoteCollector, printAsset: AssetPrinter): string[] {
	// Recorded before the early return, because the extra layout is there either way: it is
	// seeded in the constructor, not by anything printed below. Scoped to that layout's own
	// title rather than deck-wide, because the round trip matches a note's `shapeName` against
	// the identity of what differed, and an unscoped note would declare *any* extra layout.
	scopeNotes(collector, null, 'DEFAULT').note(
		'master.default',
		'approximated',
		'unsupported',
		"the write path seeds every presentation with a blank layout of its own named DEFAULT, and there is no way to remove it, so the output deck's layout gallery carries one extra entry ahead of the source's layouts; nothing binds to it, but it is visible in PowerPoint's layout picker"
	)
	if (ir.chrome.masters.length === 0) return []
	return [
		'// One master per source layout. A layout that carried decoration or placeholder',
		'// definitions has lost them — the read model decodes neither; see the fidelity notes.',
		...ir.chrome.masters.map((master) => printArguments('pptx.defineSlideMaster', [master.props], 0, printAsset)),
	]
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
