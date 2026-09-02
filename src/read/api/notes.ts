/**
 * Read-model proxies for a slide's speaker-notes slide (`notesSlideN.xml`).
 *
 * A notes slide is a small, fixed surface: a shape tree of exactly three
 * placeholders — the slide thumbnail (`sldImg`), the notes body (`body`), and the
 * slide number (`sldNum`) — and never groups, pictures, connectors, or charts.
 * So it is modeled with a dedicated {@link NotesPlaceholder} rather than the full
 * slide shape hierarchy: each placeholder exposes its `type`/`name`/`id`, its
 * geometry (EMU, `null` when the part carries no `a:xfrm`), and — for the body and
 * the slide-number field — a navigable {@link TextFrame}.
 *
 * The writer authors all three placeholders but leaves the `sldImg`/`sldNum`
 * `p:spPr` empty (no `a:xfrm`), so on an *authored* deck geometry reads `null`; an
 * imported deck carries the notesMaster-derived geometry PowerPoint stamps. The
 * body's text and the `sldNum` field, by contrast, are authored and round-trip.
 *
 * Text frames are threaded with the notes theme context (notesMaster → `theme2.xml`,
 * see {@link resolveNotesColorContext}) so a notes run's own `schemeClr` resolves to
 * a literal hex, and with the notes part's own relationships so a notes hyperlink
 * resolves its url. The body frame is additionally given a placeholder context, so a
 * body run that sets no own size/face/bold resolves its *inherited* value from the
 * notesMaster's `p:notesStyle` — the notes analogue of a slide placeholder's
 * `Run.resolved*` chain.
 */
import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { firstChild, getElements, type Element } from '../oxml/dom.js'
import type { ThemeContext } from '../oxml/theme.js'
import { resolveNotesColorContext } from './theme-context.js'
import { TextFrame } from './text.js'
import { spTreeOf } from '../oxml/slide-dom.js'
import { Placeholder } from './chrome.js'
import { buildShapes, findShapeByIdDeep, type AnyShape, type ShapeHost } from './shapes.js'

/**
 * One placeholder shape (`p:sp`) of a notes slide. A notes slide holds a fixed set:
 * the slide thumbnail (`type="sldImg"`), the notes body (`type="body"`), and the
 * slide number (`type="sldNum"`). Geometry getters read the placeholder's *own*
 * `a:xfrm` and are `null` when it inherits geometry from the notesMaster (the case
 * on an authored deck, whose `sldImg`/`sldNum` carry an empty `p:spPr`).
 *
 * A notes placeholder IS a {@link Placeholder} — identity, geometry and the escape hatch are
 * the same twelve members read off the same `p:sp`, down to `p:ph` being found the same way —
 * so it extends one rather than repeating it. Two things are its own: the flattened
 * {@link text} convenience, and a {@link textFrame} that threads the notesMaster inheritance
 * context onto the body frame.
 */
export class NotesPlaceholder extends Placeholder {
	/**
	 * The placeholder's text as a navigable {@link TextFrame} (`p:txBody`), or `null`
	 * when it carries no text body (the `sldImg` thumbnail). The body placeholder's
	 * runs round-trip a slide's authored notes; the `sldNum` placeholder's frame
	 * holds the slide-number `a:fld`, whose value surfaces through `TextFrame.text`.
	 *
	 * Overrides {@link Placeholder.textFrame} for the inheritance context alone — see below.
	 */
	override get textFrame(): TextFrame | null {
		const txBody = firstChild(this.sp, 'p:txBody')
		if (!txBody) return null
		// The notes *body* placeholder's runs inherit their effective size/face/bold
		// (and colour) from the notesMaster's `p:notesStyle`, carried on the notes theme
		// context (`resolveNotesColorContext`). Give the body frame a placeholder context
		// so `Run.resolved*` walks that chain; the `sldNum` field frame needs none.
		const ctx = this.host.themeContext()
		const placeholder =
			this.type === 'body' ? { ph: { type: this.type, idx: this.idx ?? '0' }, flatten: ctx } : undefined
		return new TextFrame(txBody, this.host.part, ctx, placeholder, this.host.relationships)
	}

	/** The placeholder's flattened text (paragraphs joined by `\n`), or `''` when it has no text body. */
	get text(): string {
		return this.textFrame?.text ?? ''
	}
}

/**
 * A slide's speaker-notes slide (`notesSlideN.xml`) as a modeled object — the deep
 * companion to {@link import('./slide.js').Slide.notesText}. Exposes the notes
 * shape tree as {@link NotesPlaceholder}s ({@link body}/{@link slideImage}/
 * {@link slideNumber}) plus the body {@link textFrame} the older getters flatten.
 *
 * The theme context (notesMaster → `theme2.xml`) is resolved once and shared across
 * every placeholder's text frame.
 */
export class NotesSlide implements ShapeHost {
	#themeContext?: ThemeContext

	constructor(
		/** The package the notes slide belongs to, for reaching the notesMaster and its theme. */
		readonly opc: OpcPackage,
		/** The notes slide's OPC part (`/ppt/notesSlides/notesSlideN.xml`). */
		readonly part: Part
	) {}

	/** Partname of the notes slide part. */
	get partName(): string {
		return this.part.partName
	}

	/** The notes part's own relationships — a notes hyperlink resolves through these. */
	get relationships(): Relationships {
		return this.opc.relationshipsFor(this.partName)
	}

	/**
	 * The notes shape tree as full shapes, in document order.
	 *
	 * A notes slide holds nothing but its three placeholders, so {@link placeholders} is the
	 * view to read. This one exists because a notes slide is a {@link ShapeHost} — which is
	 * what lets a {@link NotesPlaceholder} be the same {@link AutoShape} view of its `p:sp`
	 * that a master or layout placeholder is, rather than a second reading of one.
	 */
	get shapes(): AnyShape[] {
		const spTree = this.#spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	/** The shape anywhere in the notes tree with the given drawing id, or `undefined`. */
	shapeByIdDeep(id: number): AnyShape | undefined {
		return findShapeByIdDeep(this.shapes, id)
	}

	/**
	 * The notes theme context, resolved through the notesMaster → `theme2.xml` chain
	 * and cached. Backs {@link NotesPlaceholder.textFrame}'s `Run.resolvedColor`.
	 */
	themeContext(): ThemeContext {
		return (this.#themeContext ??= resolveNotesColorContext(this.opc, this.partName))
	}

	/** The placeholders (`p:sp`) in the notes shape tree, in document order. */
	get placeholders(): NotesPlaceholder[] {
		const spTree = this.#spTree()
		if (!spTree) return []
		return getElements(spTree, 'p:sp').map((sp) => new NotesPlaceholder(sp, this))
	}

	/** The notes body placeholder (`p:ph type="body"`), or `null` when the notes slide has none. */
	get body(): NotesPlaceholder | null {
		return this.#byType('body')
	}

	/** The slide-thumbnail placeholder (`p:ph type="sldImg"`), or `null` when absent. */
	get slideImage(): NotesPlaceholder | null {
		return this.#byType('sldImg')
	}

	/** The slide-number placeholder (`p:ph type="sldNum"`, carrying the slide-number field), or `null`. */
	get slideNumber(): NotesPlaceholder | null {
		return this.#byType('sldNum')
	}

	/** The notes body as a navigable {@link TextFrame}, or `null` when there is no body text frame. */
	get textFrame(): TextFrame | null {
		return this.body?.textFrame ?? null
	}

	/** The notes body text flattened (paragraphs joined by `\n`), or `''` when the body is empty/absent. */
	get text(): string {
		return this.body?.text ?? ''
	}

	#byType(type: string): NotesPlaceholder | null {
		return this.placeholders.find((ph) => ph.type === type) ?? null
	}

	#spTree(): Element | null {
		return spTreeOf(this.part.dom.documentElement)
	}
}
