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
 * notesMaster's `p:notesStyle` (FIDELITY-BACKLOG F2) — the notes analogue of a slide
 * placeholder's `Run.resolved*` chain.
 */
import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { attr, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'
import type { FlattenContext } from '../oxml/theme.js'
import { resolveNotesColorContext } from './theme-context.js'
import { TextFrame } from './text.js'

/** One EMU coordinate from a placeholder's `p:spPr/a:xfrm` child, or `null` when absent. */
function xfrmEmu(sp: Element, container: 'a:off' | 'a:ext', axis: string): number | null {
	const spPr = firstChild(sp, 'p:spPr')
	const xfrm = spPr && firstChild(spPr, 'a:xfrm')
	const el = xfrm && firstChild(xfrm, container)
	return el ? intValue(attr(el, axis)) : null
}

/**
 * One placeholder shape (`p:sp`) of a notes slide. A notes slide holds a fixed set:
 * the slide thumbnail (`type="sldImg"`), the notes body (`type="body"`), and the
 * slide number (`type="sldNum"`). Geometry getters read the placeholder's *own*
 * `a:xfrm` and are `null` when it inherits geometry from the notesMaster (the case
 * on an authored deck, whose `sldImg`/`sldNum` carry an empty `p:spPr`).
 */
export class NotesPlaceholder {
	constructor(
		private readonly sp: Element,
		private readonly part: Part,
		/** The notes theme context (notesMaster → `theme2.xml`), threaded to {@link textFrame}. */
		private readonly themeContext: FlattenContext,
		/** The notes part's relationships, threaded to {@link textFrame} for hyperlink resolution. */
		private readonly relationships: Relationships
	) {}

	#ph(): Element | null {
		const nvSpPr = firstChild(this.sp, 'p:nvSpPr')
		const nvPr = nvSpPr && firstChild(nvSpPr, 'p:nvPr')
		return nvPr ? firstChild(nvPr, 'p:ph') : null
	}

	#cNvPr(): Element | null {
		const nvSpPr = firstChild(this.sp, 'p:nvSpPr')
		return nvSpPr ? firstChild(nvSpPr, 'p:cNvPr') : null
	}

	/** Placeholder type (`p:ph/@type`: `sldImg` | `body` | `sldNum`), or `null` when absent. */
	get type(): string | null {
		const ph = this.#ph()
		return ph ? attr(ph, 'type') : null
	}

	/** Placeholder index (`p:ph/@idx`), or `null` when unset. */
	get idx(): string | null {
		const ph = this.#ph()
		return ph ? attr(ph, 'idx') : null
	}

	/** Shape name (`p:cNvPr/@name`), or `''` when unnamed. */
	get name(): string {
		const cNvPr = this.#cNvPr()
		return (cNvPr && attr(cNvPr, 'name')) ?? ''
	}

	/** Drawing id (`p:cNvPr/@id`), or `null` when absent. */
	get id(): number | null {
		const cNvPr = this.#cNvPr()
		return cNvPr ? intValue(attr(cNvPr, 'id')) : null
	}

	/** Left edge in EMU (`a:off/@x`), or `null` when geometry is inherited (no own `a:xfrm`). */
	get left(): number | null {
		return xfrmEmu(this.sp, 'a:off', 'x')
	}

	/** Top edge in EMU (`a:off/@y`), or `null` when geometry is inherited. */
	get top(): number | null {
		return xfrmEmu(this.sp, 'a:off', 'y')
	}

	/** Width in EMU (`a:ext/@cx`), or `null` when geometry is inherited. */
	get width(): number | null {
		return xfrmEmu(this.sp, 'a:ext', 'cx')
	}

	/** Height in EMU (`a:ext/@cy`), or `null` when geometry is inherited. */
	get height(): number | null {
		return xfrmEmu(this.sp, 'a:ext', 'cy')
	}

	/**
	 * The placeholder's text as a navigable {@link TextFrame} (`p:txBody`), or `null`
	 * when it carries no text body (the `sldImg` thumbnail). The body placeholder's
	 * runs round-trip a slide's authored notes; the `sldNum` placeholder's frame
	 * holds the slide-number `a:fld`, whose value surfaces through `TextFrame.text`.
	 */
	get textFrame(): TextFrame | null {
		const txBody = firstChild(this.sp, 'p:txBody')
		if (!txBody) return null
		// The notes *body* placeholder's runs inherit their effective size/face/bold
		// (and colour) from the notesMaster's `p:notesStyle`, carried on the notes theme
		// context (`resolveNotesColorContext`). Give the body frame a placeholder context
		// so `Run.resolved*` walks that chain; the `sldNum` field frame needs none.
		const placeholder =
			this.type === 'body' ? { ph: { type: this.type, idx: this.idx ?? '0' }, flatten: this.themeContext } : undefined
		return new TextFrame(txBody, this.part, this.themeContext, placeholder, this.relationships)
	}

	/** The placeholder's flattened text (paragraphs joined by `\n`), or `''` when it has no text body. */
	get text(): string {
		return this.textFrame?.text ?? ''
	}

	/** The underlying `p:sp` element, for advanced reads. */
	get element_(): Element {
		return this.sp
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
export class NotesSlide {
	#themeContext?: FlattenContext

	constructor(
		private readonly opc: OpcPackage,
		/** The notes slide's OPC part (`/ppt/notesSlides/notesSlideN.xml`). */
		readonly part: Part
	) {}

	/** Partname of the notes slide part. */
	get partName(): string {
		return this.part.partName
	}

	/**
	 * The notes theme context, resolved through the notesMaster → `theme2.xml` chain
	 * and cached. Backs {@link NotesPlaceholder.textFrame}'s `Run.resolvedColor`.
	 */
	themeContext(): FlattenContext {
		return (this.#themeContext ??= resolveNotesColorContext(this.opc, this.partName))
	}

	/** The placeholders (`p:sp`) in the notes shape tree, in document order. */
	get placeholders(): NotesPlaceholder[] {
		const spTree = this.#spTree()
		if (!spTree) return []
		const ctx = this.themeContext()
		const rels = this.opc.relationshipsFor(this.partName)
		return getElements(spTree, 'p:sp').map((sp) => new NotesPlaceholder(sp, this.part, ctx, rels))
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
		const root = this.part.dom.documentElement
		const cSld = root && firstChild(root, 'p:cSld')
		return cSld ? firstChild(cSld, 'p:spTree') : null
	}
}
