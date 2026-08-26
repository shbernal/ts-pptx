/**
 * Read-model proxy for one slide (`p:sld`), backed by its live part DOM.
 */
import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { relativePartName } from '../opc/partnames.js'
import type { Relationships } from '../opc/relationships.js'
import {
	OOXML_NS,
	attr,
	boolValue,
	createElement,
	firstChild,
	insertInOrder,
	intValue,
	ownerDocumentOf,
	removeAttr,
	setAttr,
	type Document,
	type Element,
} from '../oxml/dom.js'
import type { ThemeContext } from '../oxml/theme.js'
import { resolveSlideColorContext, resolveSlideThemeParts } from './theme-context.js'
import { backgroundElementOf, readSlideBackground, type SlideBackground } from './slide-background.js'
import type { Presentation } from './presentation.js'
import {
	AutoShape,
	GraphicFrame,
	GroupShape,
	Picture,
	buildShapes,
	findShapeByIdDeep,
	type AnyShape,
	type ShapeHost,
} from './shapes.js'
import { NotesSlide } from './notes.js'
import { authorNotes } from './ops/notes-author.js'
import { readModernSlideComments, readSlideComments, type Comment, type ModernComment } from './comments.js'
import { readTagsForPart, type Tag } from './tags.js'
import { SlideLayout, type SlideMaster, type Theme } from './chrome.js'
import type { TextFrame } from './text.js'
import {
	buildTransition,
	parseTransition,
	removeTransition,
	type TransitionInfo,
	type TransitionInput,
} from './transition.js'
import { enumerateSpids, flattenAnimations, hasAnimations, pruneSpids, remapSpids } from './animation.js'
import { IMAGE_REL, NOTES_SLIDE_REL, SLIDE_LAYOUT_REL } from '../../ooxml/rel-types.js'
import { InternalError, InvalidOptionError, PackageReadError } from '../../errors.js'
import { cSldOf, spTreeOf } from '../oxml/slide-dom.js'

/** Options for {@link Slide.addTextBox}. Geometry is in EMU. */
export interface AddTextBoxOptions {
	/** Left edge in EMU (`a:off/@x`). */
	left: number
	/** Top edge in EMU (`a:off/@y`). */
	top: number
	/** Width in EMU (`a:ext/@cx`); must be positive. */
	width: number
	/** Height in EMU (`a:ext/@cy`); must be positive. */
	height: number
	/** Initial text; omitted/empty yields an empty paragraph. */
	text?: string
	/** Shape name (`p:cNvPr/@name`); defaults to `TextBox <id>`. */
	name?: string
}

/** Options for {@link Slide.addPicture}. Geometry is in EMU. */
export interface AddPictureOptions {
	/** Left edge in EMU (`a:off/@x`). */
	left: number
	/** Top edge in EMU (`a:off/@y`). */
	top: number
	/** Width in EMU (`a:ext/@cx`); must be positive. */
	width: number
	/** Height in EMU (`a:ext/@cy`); must be positive. */
	height: number
	/** Shape name (`p:cNvPr/@name`); defaults to `Picture <id>`. */
	name?: string
	/**
	 * Image file extension (e.g. `png`). When omitted it is sniffed from the
	 * image's magic bytes; supply it (with `contentType`) for formats the
	 * sniffer does not recognize.
	 */
	extension?: string
	/** MIME content type (e.g. `image/png`); sniffed alongside `extension` when omitted. */
	contentType?: string
}

interface ImageType {
	extension: string
	contentType: string
}

/** Recognize a handful of common raster formats from their leading bytes. */
function sniffImageType(bytes: Uint8Array): ImageType | null {
	const b = bytes
	if (b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47)
		return { extension: 'png', contentType: 'image/png' }
	if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff)
		return { extension: 'jpeg', contentType: 'image/jpeg' }
	if (b.length >= 4 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38)
		return { extension: 'gif', contentType: 'image/gif' }
	if (b.length >= 2 && b[0] === 0x42 && b[1] === 0x4d) return { extension: 'bmp', contentType: 'image/bmp' }
	if (
		b.length >= 4 &&
		((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0x00) ||
			(b[0] === 0x4d && b[1] === 0x4d && b[2] === 0x00 && b[3] === 0x2a))
	)
		return { extension: 'tiff', contentType: 'image/tiff' }
	if (
		b.length >= 12 &&
		b[0] === 0x52 &&
		b[1] === 0x49 &&
		b[2] === 0x46 &&
		b[3] === 0x46 &&
		b[8] === 0x57 &&
		b[9] === 0x45 &&
		b[10] === 0x42 &&
		b[11] === 0x50
	)
		return { extension: 'webp', contentType: 'image/webp' }
	return null
}

export class Slide implements ShapeHost {
	constructor(
		readonly presentation: Presentation,
		/** The slide's OPC part (`/ppt/slides/slideN.xml`). */
		readonly part: Part,
		/** The slide id from `p:sldIdLst` (`p:sldId/@id`). */
		readonly slideId: number,
		/** Zero-based position in presentation order. */
		readonly index: number
	) {}

	/** Partname of this slide's part. */
	get partName(): string {
		return this.part.partName
	}

	/**
	 * Escape hatch: the underlying `p:sld` element. After mutating it call
	 * {@link markDirty}, or `save()` writes the original bytes. Slide-level DOM
	 * access is also reachable as `slide.part.dom`; this getter is the same node,
	 * on the same rung of the ladder as `Shape.element_` and the rest.
	 */
	get element_(): Element {
		const root = this.part.dom.documentElement
		if (!root)
			throw new PackageReadError('package/part-has-no-root', `Slide ${this.partName} has no root <p:sld> element`)
		return root
	}

	/** Mark this slide's part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/** This slide part's relationships (image embeds, layout, hyperlinks, …). */
	get relationships(): Relationships {
		return this.presentation.opc.relationshipsFor(this.partName)
	}

	/**
	 * The deck's OPC package. Convenience for `presentation.opc`, and the member
	 * {@link ShapeHost} names so a shape can reach a referenced part (an image, a
	 * chart) the same way whether it lives on a slide, a layout, or a master.
	 */
	get opc(): OpcPackage {
		return this.presentation.opc
	}

	#themeColors?: ThemeContext

	/**
	 * The slide's resolved theme context (`clrMap` + `clrScheme` + the theme
	 * `fmtScheme` + `fontScheme`, plus the layout/master roots), walked once from
	 * slide → layout → master → theme and cached on this proxy. Backs the read-model
	 * resolved getters ({@link Shape.resolvedFill}, `Run.resolvedColor`,
	 * `Run.resolvedSizePt`, `Run.resolvedFontFace`) so a `schemeClr` token — including
	 * one delivered through a shape's `p:style` `fillRef`/`lnRef` — resolves to a
	 * literal hex, and a placeholder run's inherited size/typeface (a `+mj-*`/`+mn-*`
	 * font token included) resolves to a literal value. The maps/roots are empty when
	 * the theme chain is incomplete, in which case tokens simply stay unresolved.
	 */
	themeContext(): ThemeContext {
		return (this.#themeColors ??= resolveSlideColorContext(this.presentation.opc, this.partName))
	}

	/** Authoring name of the slide (`p:cSld/@name`), or `null` if unnamed. */
	get name(): string | null {
		const cSld = this.#cSld()
		return cSld ? attr(cSld, 'name') : null
	}

	/**
	 * Whether this slide is hidden (`p:sld/@show="0"`). The attribute is
	 * `xsd:boolean` defaulting to `true`, so an absent attribute means shown.
	 * Hidden slides are dropped from PowerPoint/LibreOffice presentations and
	 * exported PDFs, so render order diverges from model order when any earlier
	 * slide is hidden.
	 */
	get hidden(): boolean {
		const root = this.part.dom.documentElement
		const show = root ? attr(root, 'show') : null
		return show === '0' || show === 'false'
	}

	/**
	 * Hide or show this slide. Hiding writes `p:sld/@show="0"`; showing removes
	 * the attribute, restoring PowerPoint's canonical shown form (absent ⇒ shown).
	 * Marks the slide part dirty.
	 */
	set hidden(value: boolean) {
		const root = this.part.dom.documentElement
		if (!root)
			throw new PackageReadError('package/part-has-no-root', `Slide ${this.partName} has no root <p:sld> element`)
		if (value) setAttr(root, 'show', '0')
		else removeAttr(root, 'show')
		this.part.markDirty()
	}

	/**
	 * Whether this slide draws the master's non-placeholder shapes
	 * (`p:sld/@showMasterSp`). `xsd:boolean` defaulting to `true`, so an absent
	 * attribute means shown — the same shape as {@link hidden}.
	 *
	 * This is the switch that decides whether {@link SlideMaster.shapes} belongs on
	 * *this* slide. A slide the author gave a full-bleed image or a section divider
	 * usually sets it to `0`, and a renderer that paints the master's band and logo
	 * anyway puts the template's furniture on top of a slide that deliberately hid
	 * it. Placeholders are unaffected: the flag suppresses only the master's
	 * decorative shapes.
	 */
	get showMasterSp(): boolean {
		const root = this.part.dom.documentElement
		return boolValue(root && attr(root, 'showMasterSp')) !== false
	}

	/**
	 * The slide's show transition (`p:transition`), decoded into a typed model, or
	 * `null` when the slide has none. Handles both PowerPoint forms: the bare
	 * `<p:transition>` and the `mc:AlternateContent` wrapper that carries the exact
	 * `p14:dur` duration (the `p14` Choice is preferred so `durationMs` is recovered).
	 */
	get transition(): TransitionInfo | null {
		const root = this.part.dom.documentElement
		return root ? parseTransition(root) : null
	}

	/**
	 * Set or clear the slide's show transition. Assigning `null` removes it.
	 * Writing a transition with `durationMs` emits the `mc:AlternateContent` form
	 * (a `p14` Choice carrying `p14:dur` plus a base `mc:Fallback`); otherwise the
	 * bare `<p:transition>` is written. The node is inserted at its schema slot —
	 * after `p:clrMapOvr`, before `p:timing`/`p:extLst`. Marks the slide part dirty.
	 */
	set transition(value: TransitionInput | null) {
		const root = this.part.dom.documentElement
		if (!root)
			throw new PackageReadError('package/part-has-no-root', `Slide ${this.partName} has no root <p:sld> element`)
		removeTransition(root)
		if (value) {
			const doc = this.part.dom
			insertInOrder(root, buildTransition(doc, value), ['p:timing', 'p:extLst'])
		}
		this.part.markDirty()
	}

	/**
	 * Whether the slide carries build animations (`p:timing` with a `<p:bldP>` or a
	 * `presetID`-bearing time node). The animation tree itself is preserved opaquely;
	 * see {@link animationSpids}.
	 */
	get hasAnimations(): boolean {
		const root = this.part.dom.documentElement
		return root ? hasAnimations(root) : false
	}

	/**
	 * @internal The sorted, de-duplicated set of shape ids (`spid`) referenced by
	 * the slide's animations (`<p:spTgt spid>` and `<p:bldP spid>`). Exposed for the
	 * import paths' spid-coherence checks.
	 */
	animationSpids(): number[] {
		const root = this.part.dom.documentElement
		return root ? enumerateSpids(root) : []
	}

	/**
	 * @internal Rewrite every animation `spid` per `mapping` (old → new), keeping
	 * the opaque timing tree coherent when shape ids are reassigned. Marks the slide
	 * part dirty only when a reference changed.
	 */
	remapAnimationSpids(mapping: Map<number, number>): void {
		const root = this.part.dom.documentElement
		if (root && remapSpids(root, mapping)) this.part.markDirty()
	}

	/**
	 * @internal Remove the build animations targeting the given shape ids (their
	 * `<p:bldP>` and effect nodes), so deleting a shape never leaves a dangling
	 * `spid` reference. Marks the slide part dirty only when something was removed.
	 */
	pruneAnimationSpids(spids: Iterable<number>): void {
		const root = this.part.dom.documentElement
		if (root && pruneSpids(root, spids)) this.part.markDirty()
	}

	/**
	 * Flatten the slide's build animations: remove the `<p:timing>` block so the
	 * slide renders and edits as its final static state, with every shape shown at
	 * once. Gated like {@link hasAnimations} — a `<p:timing>` that is purely a media
	 * loop (no `<p:bldP>` or `presetID`) is preserved so media playback survives.
	 * Marks the slide part dirty and returns `true` only when a timing block was
	 * removed.
	 *
	 * This removes click-through staging only; it does not delete shapes. If a slide
	 * animated alternating states over the same region, the flattened render shows
	 * them all at once — keep that distinct from removing staged/duplicate shapes.
	 */
	flattenAnimations(): boolean {
		const root = this.part.dom.documentElement
		if (root && flattenAnimations(root)) {
			this.part.markDirty()
			return true
		}
		return false
	}

	/** Top-level shapes in the slide's shape tree, in document order. */
	get shapes(): AnyShape[] {
		const spTree = this.#spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	/**
	 * All text on the slide, flattened in document order — the read-model
	 * counterpart to {@link TextFrame.text} one level up. Every text-bearing shape
	 * contributes its text, recursing into groups ({@link GroupShape}) and reading
	 * table cells ({@link GraphicFrame.table}); text-free shapes (pictures,
	 * connectors, empty boxes) contribute nothing. Blocks are joined by `\n`; within
	 * a table, cells in a row are joined by `\t` and rows by `\n`.
	 *
	 * This is deliberately scoped to the slide's own shape tree. It does **not**
	 * include speaker notes (read those via {@link notesText}) or chart data labels
	 * (read those via {@link GraphicFrame.chart}); folding those in would force
	 * ordering and separator choices a caller is better placed to make. Extract a
	 * whole deck with `deck.slides.map((s) => s.text)`.
	 */
	get text(): string {
		const blocks: string[] = []
		const walk = (shapes: AnyShape[]): void => {
			for (const shape of shapes) {
				if (shape instanceof GroupShape) {
					walk(shape.shapes)
				} else if (shape instanceof GraphicFrame) {
					const table = shape.table
					if (!table) continue
					for (const row of table.rows) {
						const cells = row.cells.map((cell) => cell.text)
						if (cells.some((cell) => cell.length > 0)) blocks.push(cells.join('\t'))
					}
				} else {
					const text = shape.text
					if (text.length > 0) blocks.push(text)
				}
			}
		}
		walk(this.shapes)
		return blocks.join('\n')
	}

	/**
	 * The slide's speaker-note text — the companion to {@link text} for the notes
	 * that {@link text} deliberately excludes. Delegates to {@link notesSlide}: finds
	 * the notes body placeholder (`p:ph` `type="body"`) and flattens its text frame
	 * the same way {@link TextFrame.text} does (paragraphs joined by `\n`).
	 *
	 * Returns `null` when the slide has **no notes slide part at all** — distinct
	 * from `''`, which means a notes slide exists but its body is empty (PowerPoint
	 * often attaches an empty notes slide to every slide). Only the body placeholder
	 * is read; the slide-thumbnail (`sldImg`) and slide-number (`sldNum`)
	 * placeholders a notes slide also carries are read via {@link notesSlide}.
	 */
	get notesText(): string | null {
		const notes = this.notesSlide
		return notes ? notes.text : null
	}

	/**
	 * The slide's speaker-note body as a navigable {@link TextFrame} — the rich
	 * companion to {@link notesText}, which flattens the same body to a plain string.
	 * Walk `paragraphs → runs` to recover per-run formatting (bold/italic/underline,
	 * colour, size, face) and any notes hyperlink that {@link notesText} discards.
	 *
	 * `null` when the slide has **no notes slide part** (the same boundary
	 * {@link notesText} returns `null` at), and also when a notes part exists but
	 * carries no body-placeholder text frame — there is then no frame to hand back
	 * (whereas {@link notesText} reports `''` for that empty-body case). Convenience
	 * for `notesSlide.textFrame`; see {@link notesSlide} for the whole modeled notes
	 * slide (its `sldImg`/`sldNum` placeholders and their geometry).
	 */
	get notesTextFrame(): TextFrame | null {
		return this.notesSlide?.textFrame ?? null
	}

	/**
	 * The slide's speaker-notes slide (`notesSlideN.xml`) as a modeled
	 * {@link NotesSlide}, or `null` when the slide has no notes slide part. Beyond the
	 * body text {@link notesText}/{@link notesTextFrame} surface, this exposes the
	 * notes slide's three placeholders — the slide thumbnail (`slideImage`), the notes
	 * body, and the slide-number field (`slideNumber`) — each with its geometry and
	 * (where present) a navigable text frame.
	 *
	 * The writer authors all three placeholders but leaves `sldImg`/`sldNum` with an
	 * empty `p:spPr`, so on an authored deck their geometry reads `null` while an
	 * imported deck carries the notesMaster-derived geometry. The body text and the
	 * slide-number field round-trip either way.
	 */
	get notesSlide(): NotesSlide | null {
		const notesRel = this.relationships.byType(NOTES_SLIDE_REL)[0]
		if (!notesRel) return null
		const notesPart = this.presentation.opc.part(this.relationships.resolveTarget(notesRel.id))
		if (!notesPart) return null
		return new NotesSlide(this.presentation.opc, notesPart)
	}

	/**
	 * The slide layout this slide is bound to (`slideLayout` relationship), as a
	 * modeled {@link SlideLayout}, or `null` when the slide has no layout. The layout
	 * carries the placeholder geometry and background a slide inherits; walk on to
	 * {@link master} and {@link theme} through it (`slide.layout.master.theme`).
	 */
	get layout(): SlideLayout | null {
		const rel = this.relationships.byType(SLIDE_LAYOUT_REL)[0]
		if (!rel) return null
		const part = this.presentation.opc.part(this.relationships.resolveTarget(rel.id))
		return part ? new SlideLayout(this.presentation.opc, part) : null
	}

	/**
	 * The slide master this slide resolves against, via its {@link layout}, as a
	 * modeled {@link SlideMaster}, or `null` when the layout/master chain is broken.
	 * The master owns the colour map (`schemeClr` token → theme slot) and the default
	 * text styles a placeholder inherits.
	 */
	get master(): SlideMaster | null {
		return this.layout?.master ?? null
	}

	/**
	 * The theme this slide resolves colour and font tokens against, via its
	 * {@link layout} → {@link master} → theme, as a modeled {@link Theme}, or `null`.
	 * Read its {@link Theme.colorScheme}/{@link Theme.fontScheme} to see the literal
	 * palette and faces a `schemeClr`/`+mj-*` token resolves to.
	 */
	get theme(): Theme | null {
		return this.layout?.master?.theme ?? null
	}

	/**
	 * The slide's **legacy** review comments (`p:cm` in its `comments/commentN.xml`
	 * part), `[]` when it has none. Each {@link Comment} carries its body text, marker
	 * position (EMU), timestamp, and its author resolved through the deck-wide
	 * {@link Presentation.commentAuthors} registry (`@authorId` → name/initials). These
	 * are the comments the writer authors via `slide.addComment(...)`; the 2018 modern
	 * comment parts (`p188:cm`) are a separate schema, preserved but not decoded.
	 */
	get comments(): Comment[] {
		return readSlideComments(this.presentation.opc, this.part, this.presentation.commentAuthors)
	}

	/**
	 * The slide's **modern** (2018) review comments (`p188:cm` in its
	 * `modernComment_*.xml` part), `[]` when it has none. Each {@link ModernComment}
	 * carries its body text, `created` timestamp, marker position (EMU), its author
	 * resolved through the deck-wide {@link Presentation.modernCommentAuthors}
	 * registry (`@authorId` GUID → name/initials), and its reply thread nested under
	 * `replies`. Read-only — the writer authors the legacy schema; see
	 * {@link Presentation.commentSchema} to tell which schema a deck uses.
	 */
	get modernComments(): ModernComment[] {
		return readModernSlideComments(this.presentation.opc, this.part, this.presentation.modernCommentAuthors)
	}

	/**
	 * The slide's programmatic tags (`p:custDataLst/p:tags` on the slide part,
	 * resolved to `ppt/tags/tagN.xml`) as `{ name, val }` string pairs, `[]` when it
	 * has none. Host/add-in metadata with no visible rendering and no writer —
	 * read-only, preserved byte-for-byte on round-trip. Deck-level tags are
	 * {@link Presentation.tags}.
	 */
	get tags(): Tag[] {
		return readTagsForPart(this.presentation.opc, this.partName)
	}

	/**
	 * The slide's effective background (`p:cSld/p:bg`), decoded into a typed
	 * {@link SlideBackground}. Resolved through the inheritance chain: the slide's
	 * own `p:bg` wins; failing that the slideLayout's, then the slideMaster's — the
	 * result's `source` records which supplied it. `null` when nothing in the chain
	 * defines a background.
	 *
	 * Colour tokens resolve against this slide's theme; an image background's
	 * `r:embed` resolves to an absolute part name through the *owning* part's
	 * relationships (the layout's rels for a layout-inherited image, etc.). Solid,
	 * gradient, and image backgrounds are the three the writer authors and so
	 * round-trip faithfully; `pattern`/`themeRef` are read-only for imported decks.
	 */
	get background(): SlideBackground | null {
		const opc = this.presentation.opc
		const parts = resolveSlideThemeParts(opc, this.partName)
		const ctx = this.themeContext()
		const candidates: { root: Element | null; source: 'slide' | 'layout' | 'master'; partName: string | null }[] = [
			{ root: parts.slideRoot, source: 'slide', partName: this.partName },
			{ root: parts.layoutRoot, source: 'layout', partName: parts.layoutPartName },
			{ root: parts.masterRoot, source: 'master', partName: parts.masterPartName },
		]
		const themeRels = parts.themePartName ? opc.relationshipsFor(parts.themePartName) : null
		for (const { root, source, partName } of candidates) {
			const bg = backgroundElementOf(root)
			if (!bg) continue
			const rels = partName ? opc.relationshipsFor(partName) : null
			return readSlideBackground(bg, source, ctx, rels, themeRels)
		}
		return null
	}

	/**
	 * The slide's own slide-number placeholder (`p:sp` with `p:ph type="sldNum"`,
	 * carrying an `<a:fld type="slidenum">`), or `null` when the slide shows no slide
	 * number of its own. This is the concrete shape the writer emits for
	 * `pptx.setSlideNumber(...)` on a slide's layout; its geometry and run formatting
	 * read off the returned {@link AutoShape} the same as any placeholder.
	 *
	 * Scoped to the slide's *own* shape tree — a slide number inherited purely from
	 * the master's `p:hf`/placeholder (with no shape on the slide) is a master-level
	 * concern this getter does not resolve.
	 */
	get slideNumberPlaceholder(): AutoShape | null {
		return this.placeholder('sldNum') ?? null
	}

	/** The first top-level shape with the given drawing id (`p:cNvPr/@id`), or `undefined`. */
	shapeById(id: number): AnyShape | undefined {
		return this.shapes.find((shape) => shape.id === id)
	}

	/**
	 * The first shape *anywhere* in the slide's shape tree with the given drawing id
	 * (`p:cNvPr/@id`), descending into groups — unlike {@link shapeById}, which scans
	 * only top-level shapes. Walked pre-order (a group is visited before its children),
	 * matching how the writer allocates ids. `undefined` when no shape carries that id.
	 *
	 * Drawing ids are unique within a slide, so the first match is the only match; the
	 * pre-order walk just fixes a deterministic order. This backs the connector-binding
	 * resolution ({@link import('./shapes.js').Connector.startConnection}), which must
	 * resolve a binding into a group that top-level {@link shapeById} cannot see.
	 */
	shapeByIdDeep(id: number): AnyShape | undefined {
		return findShapeByIdDeep(this.shapes, id)
	}

	/** The first top-level shape with the given name (`p:cNvPr/@name`), or `undefined`. */
	shapeByName(name: string): AnyShape | undefined {
		return this.shapes.find((shape) => shape.name === name)
	}

	/**
	 * The first placeholder of the given type (`p:ph/@type`, e.g. `title`,
	 * `ctrTitle`, `subTitle`, `body`), optionally narrowed by `idx`. Returns
	 * `undefined` when none match. Only `p:sp` shapes can be placeholders, so the
	 * result is an {@link AutoShape}.
	 */
	placeholder(type: string, idx?: string): AutoShape | undefined {
		return this.shapes.find((shape): shape is AutoShape => {
			const ph = shape instanceof AutoShape ? shape.placeholder : null
			return ph !== null && ph.type === type && (idx === undefined || ph.idx === idx)
		})
	}

	/**
	 * Append a text box (`p:sp` with `txBox="1"`) to the slide's shape tree and
	 * return it. Geometry is required (EMU); width and height must be positive.
	 * Allocates a drawing id unique within the slide. Marks the slide part dirty.
	 */
	addTextBox(options: AddTextBoxOptions): AutoShape {
		const { left, top, width, height } = options
		requireFinite(left, 'left')
		requireFinite(top, 'top')
		requirePositive(width, 'width')
		requirePositive(height, 'height')

		const spTree = this.#spTree()
		if (!spTree)
			throw new PackageReadError('slide/no-shape-tree', `Slide ${this.partName} has no spTree to add a shape to`)
		const doc = spTree.ownerDocument
		if (!doc) throw new InternalError('oxml/node-has-no-document', 'Slide DOM has no owner document')

		const id = this.#nextShapeId()
		const sp = buildTextBox(doc, {
			id,
			name: options.name ?? `TextBox ${id}`,
			text: options.text ?? '',
			left: Math.round(left),
			top: Math.round(top),
			width: Math.round(width),
			height: Math.round(height),
		})
		this.#appendShape(spTree, sp)
		return new AutoShape(sp, this)
	}

	/**
	 * Add a picture (`p:pic`) from raw image bytes and return it. Creates a media
	 * part under `/ppt/media/`, registers its content type, wires an `image`
	 * relationship from this slide, and appends the picture to the shape tree.
	 * Geometry is required (EMU); width and height must be positive. The image
	 * format is sniffed from the bytes unless `extension`/`contentType` are given.
	 */
	addPicture(image: Uint8Array, options: AddPictureOptions): Picture {
		const { left, top, width, height } = options
		requireFinite(left, 'left')
		requireFinite(top, 'top')
		requirePositive(width, 'width')
		requirePositive(height, 'height')

		const sniffed = sniffImageType(image)
		const extension = (options.extension ?? sniffed?.extension)?.toLowerCase().replace(/^\./, '')
		const contentType = options.contentType ?? sniffed?.contentType
		if (!extension || !contentType) {
			throw new InvalidOptionError(
				'image/undeterminable-type',
				'Could not determine image type; pass { extension, contentType } to addPicture'
			)
		}

		const spTree = this.#spTree()
		if (!spTree)
			throw new PackageReadError('slide/no-shape-tree', `Slide ${this.partName} has no spTree to add a picture to`)
		const doc = spTree.ownerDocument
		if (!doc) throw new InternalError('oxml/node-has-no-document', 'Slide DOM has no owner document')

		const opc = this.presentation.opc
		const mediaPartName = opc.reserveMediaPartName(extension)
		opc.addPart(mediaPartName, contentType, image)
		const relId = this.relationships.add(IMAGE_REL, relativePartName(this.partName, mediaPartName)).id

		const id = this.#nextShapeId()
		const pic = buildPicture(doc, {
			id,
			name: options.name ?? `Picture ${id}`,
			relId,
			left: Math.round(left),
			top: Math.round(top),
			width: Math.round(width),
			height: Math.round(height),
		})
		this.#appendShape(spTree, pic)
		return new Picture(pic, this)
	}

	/**
	 * Give this slide the speaker notes `text`, and return the resulting
	 * {@link NotesSlide}. A `
` starts a new paragraph, matching the write-side
	 * `addNotes`; the runs carry no formatting of their own, so style them
	 * afterwards through {@link notesTextFrame}.
	 *
	 * The counterpart to the read side's `notesText`/`notesTextFrame`/`notesSlide`
	 * getters, and the way to annotate a slide that has **no notes part at all** —
	 * the state an `importSlide` without `{ importNotes: true }` leaves behind, and
	 * the one a `notesTextFrame` edit cannot reach because there is no frame to
	 * hand back. Calling it on a slide that already has notes replaces the body
	 * text and leaves the rest of the part (its geometry, its other two
	 * placeholders) alone.
	 *
	 * Creating the part pulls in what a notes slide must bind to: a `notesMaster`,
	 * of which a presentation may hold at most one. This deck's own is reused when
	 * it has one; otherwise one is installed, bound to a clone of this deck's theme
	 * so it resolves against the destination palette. That is the same
	 * single-master rule `importSlide({ importNotes: true })` and `appendSlides`
	 * follow, so mixing the three cannot produce a second notes master.
	 */
	addNotes(text: string): NotesSlide {
		const partName = authorNotes(this, text)
		const part = this.presentation.opc.part(partName)
		if (!part) throw new InternalError('import/part-went-missing', `Authored notes part went missing: ${partName}`)
		return new NotesSlide(this.presentation.opc, part)
	}

	/**
	 * @internal The slide's shape tree (`p:cSld/p:spTree`), or `null` if absent.
	 * Exposed for cross-slide composition (`Presentation.importShape`).
	 */
	shapeTree(): Element | null {
		return this.#spTree()
	}

	/**
	 * @internal The smallest drawing id (`p:cNvPr/@id`) not already used on this
	 * slide. Exposed so `Presentation.importShape` can give a lifted shape (and its
	 * group children) collision-free ids on the host.
	 */
	nextShapeId(): number {
		return this.#nextShapeId()
	}

	/** Insert a shape after grpSpPr and before any trailing p:extLst on the tree; mark dirty. */
	#appendShape(spTree: Element, shape: Element): void {
		spTree.insertBefore(shape, firstChild(spTree, 'p:extLst'))
		this.part.markDirty()
	}

	/** The smallest drawing id (`p:cNvPr/@id`) not already used on the slide. */
	#nextShapeId(): number {
		const root = this.part.dom.documentElement
		let max = 1
		if (root) {
			for (const cNvPr of root.getElementsByTagNameNS(OOXML_NS.p, 'cNvPr')) {
				const id = intValue(attr(cNvPr, 'id'))
				if (id !== null && id > max) max = id
			}
		}
		return max + 1
	}

	#cSld(): Element | null {
		return cSldOf(this.part.dom.documentElement)
	}

	#spTree(): Element | null {
		return spTreeOf(this.part.dom.documentElement)
	}
}

function requireFinite(value: number, name: string): void {
	if (!Number.isFinite(value))
		throw new InvalidOptionError('coord/non-finite', `${name} must be a finite number of EMU, got ${value}`)
}

function requirePositive(value: number, name: string): void {
	requireFinite(value, name)
	if (value <= 0) throw new InvalidOptionError('coord/not-positive', `${name} must be positive, got ${value}`)
}

/** The box a built shape sits in, in EMU. */
interface BoxSpec {
	left: number
	top: number
	width: number
	height: number
}

/**
 * Create `qname` in the parent's document and append it.
 *
 * The builders below are uniformly parent-then-child, so this is the only shape
 * of element creation they need; `ownerDocumentOf` is what lets it take the
 * parent alone rather than threading a `Document` through every call.
 */
function appendEl(parent: Element, qname: string): Element {
	const child = createElement(ownerDocumentOf(parent), qname)
	parent.appendChild(child)
	return child
}

/**
 * Append the `p:spPr` a built `p:sp` and `p:pic` share: the spec's box as an
 * `a:xfrm`, then a rect `prstGeom` with the empty `a:avLst` the schema requires.
 */
function appendSpPr(parent: Element, spec: BoxSpec): Element {
	const spPr = appendEl(parent, 'p:spPr')
	const xfrm = appendEl(spPr, 'a:xfrm')
	const off = appendEl(xfrm, 'a:off')
	setAttr(off, 'x', String(spec.left))
	setAttr(off, 'y', String(spec.top))
	const ext = appendEl(xfrm, 'a:ext')
	setAttr(ext, 'cx', String(spec.width))
	setAttr(ext, 'cy', String(spec.height))
	const prstGeom = appendEl(spPr, 'a:prstGeom')
	setAttr(prstGeom, 'prst', 'rect')
	appendEl(prstGeom, 'a:avLst')
	return spPr
}

interface TextBoxSpec extends BoxSpec {
	id: number
	name: string
	text: string
}

/** Build a minimal, schema-valid text-box `p:sp` element (not yet attached). */
function buildTextBox(doc: Document, spec: TextBoxSpec): Element {
	const sp = createElement(doc, 'p:sp')

	const nvSpPr = appendEl(sp, 'p:nvSpPr')
	const cNvPr = appendEl(nvSpPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvSpPr = appendEl(nvSpPr, 'p:cNvSpPr')
	setAttr(cNvSpPr, 'txBox', '1')
	appendEl(nvSpPr, 'p:nvPr')

	appendSpPr(sp, spec)

	const txBody = appendEl(sp, 'p:txBody')
	appendEl(txBody, 'a:bodyPr')
	appendEl(txBody, 'a:lstStyle')
	const p = appendEl(txBody, 'a:p')
	if (spec.text !== '') {
		const r = appendEl(p, 'a:r')
		const t = appendEl(r, 'a:t')
		t.textContent = spec.text
		if (spec.text !== spec.text.trim()) setAttr(t, 'xml:space', 'preserve')
	}

	return sp
}

interface PictureSpec extends BoxSpec {
	id: number
	name: string
	relId: string
}

/** Build a minimal, schema-valid `p:pic` element (not yet attached). */
function buildPicture(doc: Document, spec: PictureSpec): Element {
	const pic = createElement(doc, 'p:pic')

	const nvPicPr = appendEl(pic, 'p:nvPicPr')
	const cNvPr = appendEl(nvPicPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvPicPr = appendEl(nvPicPr, 'p:cNvPicPr')
	const picLocks = appendEl(cNvPicPr, 'a:picLocks')
	setAttr(picLocks, 'noChangeAspect', '1')
	appendEl(nvPicPr, 'p:nvPr')

	const blipFill = appendEl(pic, 'p:blipFill')
	const blip = appendEl(blipFill, 'a:blip')
	setAttr(blip, 'r:embed', spec.relId)
	const stretch = appendEl(blipFill, 'a:stretch')
	appendEl(stretch, 'a:fillRect')

	appendSpPr(pic, spec)

	return pic
}
