/**
 * The parts of printing a script that do not depend on which tier is being printed.
 *
 * Two printers target the same IR — one anchored to the source deck as a template, one
 * standalone — and they differ only in how a deck and its chrome are set up. Everything after
 * that is identical: a slide's statements, how image bytes reach the script, and how the
 * fidelity notes are rendered into the banner. Keeping that here is not tidiness. A slide body
 * printed two slightly different ways would mean the round-trip oracle certifies one tier and
 * only resembles a check on the other, and the drift would be invisible in review.
 */
import { asIrValue, type AssetRef, type DeckIr, type IrValue, type SlideIr } from '../ir.js'
import { scopeNotes, type FidelityNote, type NoteCollector } from '../fidelity.js'
import { printArguments, printString, printValue, type AssetPrinter } from './literal.js'
import { InvalidOptionError } from '../../errors.js'

/** How image bytes reach the emitted script. */
export type AssetMode =
	/** Written beside the script and read at run time. Keeps the script readable. */
	| 'file'
	/** Inlined as `data:` literals. One self-contained file, at roughly 4/3 the byte size. */
	| 'inline'

export interface PrintedScript {
	/** The TypeScript module source. */
	code: string
	/**
	 * Image bytes the script expects to find in its asset directory, keyed by filename.
	 * Empty when the print options asked for `'inline'`, since the bytes are then in
	 * {@link code}.
	 */
	assets: Map<string, Uint8Array>
	/**
	 * The losses that apply to **this** output: the IR's own notes, minus any this tier
	 * rescues, plus the ones this tier causes. Reproduced as a comment block at the top of
	 * {@link code}, so the artifact carries its own caveats.
	 */
	notes: FidelityNote[]
}

/**
 * Default import specifier for an emitted script — this package's *published* name, which is
 * not its directory name. Getting it wrong produces a script that prints and typechecks and
 * then fails at `import`, so it is pinned by a test rather than left to a literal.
 */
export const PACKAGE_NAME = '@shbernal/ts-pptx'

/** The print options both tiers share, with their defaults applied. */
export interface ResolvedPrintOptions {
	/** Where the emitted script writes its deck, resolved against the script's own location. */
	outputPath: string
	/** Where the emitted script reads image assets from, with any trailing slash stripped. */
	assetDir: string
	/** Whether image bytes are written beside the script or inlined as `data:` literals. */
	assetMode: AssetMode
	/** The import specifier the emitted script uses. */
	packageName: string
}

/**
 * Apply the defaults both printers share.
 *
 * `templatePath` is deliberately not here: only the template-anchored tier has one, and that
 * difference is the whole reason there are two printers.
 * @param options - the caller's options, any of which may be absent
 */
export function resolvePrintOptions(options: {
	outputPath?: string
	assetDir?: string
	assets?: AssetMode
	packageName?: string
}): ResolvedPrintOptions {
	return {
		outputPath: options.outputPath ?? './output.pptx',
		// A trailing slash would print `./assets/` and then join another, so it is stripped once
		// here rather than defended against at each of the places the directory is interpolated.
		assetDir: (options.assetDir ?? './assets').replace(/\/$/, ''),
		assetMode: options.assets ?? 'file',
		packageName: options.packageName ?? PACKAGE_NAME,
	}
}

/** Width the header comment's prose wraps at, leaving room for the ` * ` prefix. */
const COMMENT_WIDTH = 110

/**
 * A stable JavaScript identifier per asset.
 *
 * Derived from the asset's own name rather than its index, so a reader can match a `const`
 * in the script against a file in the asset directory. The dedupe suffix is defensive: the
 * read half currently names assets `image1.png`, `image2.jpg`, … which cannot collide after
 * the extension is stripped, but nothing in the IR contract promises that.
 */
export function assetIdentifiers(ir: DeckIr): Map<string, string> {
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

/** Resolve an {@link AssetRef} to the `const` identifier standing for its bytes. */
export function assetPrinter(identifiers: Map<string, string>): AssetPrinter {
	return (ref: AssetRef): string => {
		const identifier = identifiers.get(ref.$asset)
		if (!identifier)
			throw new InvalidOptionError(
				'script/unresolved-asset-reference',
				`Asset reference ${ref.$asset} has no entry in DeckIr.assets`
			)
		return identifier
	}
}

/**
 * One `const` per asset, each holding a complete `data:` URI.
 *
 * Both modes bind the same *kind* of value — a data URI string — so the call sites are
 * identical and `addImage`'s `data` option is used either way. Only where the bytes come
 * from changes. Rewriting the option key to `path` in file mode would have the printer
 * reshaping an argument, which is precisely the thing that keeps the IR checkable.
 */
export function printAssetBindings(
	ir: DeckIr,
	identifiers: Map<string, string>,
	assetDir: string,
	mode: AssetMode,
	inlineComment = '// Media bytes, inlined so this script needs no files but its template.'
): string[] {
	if (ir.assets.length === 0) return []

	const lines =
		mode === 'file' ? ['// Media bytes, loaded from the asset directory beside this script.'] : [inlineComment]

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
 * One slide: a comment, the statement that creates it, then its slide-level properties and
 * its calls in z-order.
 *
 * `comment` and `construction` are the tier's business — one binds a layout by name at append
 * time, the other names a master up front — and everything after them is not.
 */
export function printSlide(
	slide: SlideIr,
	parts: { comment: string; construction: string },
	collector: NoteCollector,
	printAsset: AssetPrinter
): string[] {
	const notes = scopeNotes(collector, slide.number)
	const identifier = `slide${slide.number}`
	const lines = ['', `// ${parts.comment}`, `const ${identifier} = ${parts.construction}`]

	if (slide.name !== undefined) {
		notes.note(
			'slide.name',
			'dropped',
			'unwritable',
			`the slide's own name (p:cSld@name, here ${JSON.stringify(slide.name)}) has no public write-API setter, so the generated slide is unnamed`
		)
	}

	if (slide.hidden) lines.push(`${identifier}.hidden = true`)
	if (slide.background) {
		lines.push(`${identifier}.background = ${printValue(slide.background as IrValue, 0, printAsset)}`)
	}
	if (slide.transition) {
		lines.push(`${identifier}.transition = ${printValue(asIrValue(slide.transition), 0, printAsset)}`)
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

/** The banner comment: `prose` as written, then every loss that applies to this output. */
export function header(prose: string[], notes: FidelityNote[]): string {
	const lines = [...prose, '']

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

/**
 * Close out a printed script: the accumulated lines as one module source, plus the
 * asset side-car the caller has to write beside it.
 *
 * Both tiers end this way, and the `assetMode` test is the whole reason it is worth
 * naming. In `'inline'` mode the bytes are already in {@link PrintedScript.code} as
 * `data:` literals, so the map must come back **empty** — a tier that returned them
 * anyway would have its caller write every image to disk as well, and the script would
 * ignore the files it was handed.
 * @param {string[]} lines - the script's lines, in order
 * @param {DeckIr} ir - the IR being printed, for its assets
 * @param {AssetMode} assetMode - how image bytes reach the script
 * @param {FidelityNote[]} notes - the losses that apply to this output
 * @returns {PrintedScript} the finished script
 */
export function printedScript(lines: string[], ir: DeckIr, assetMode: AssetMode, notes: FidelityNote[]): PrintedScript {
	return {
		code: lines.join('\n'),
		assets: new Map(
			assetMode === 'file' ? ir.assets.map((asset): [string, Uint8Array] => [asset.name, asset.bytes]) : []
		),
		notes,
	}
}
