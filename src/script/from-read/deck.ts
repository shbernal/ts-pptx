/**
 * `Presentation` → {@link DeckIr}. The entry point of the read half.
 *
 * Pure in the sense that matters: it reads a loaded package and returns data, touching no
 * filesystem and mutating nothing. Given the same bytes it returns the same IR, so a
 * round-trip check can compare two runs directly.
 *
 * The one judgement made at this level rather than per-shape is {@link SlideSource}. A
 * slide holding a construct the write API cannot express *at all* — a graphic frame with a
 * complete reader and no emitter, which today means an extended chart, a SmartArt diagram,
 * an OLE object or ink — cannot be honestly transcribed, because the output would silently
 * differ from the source. Such a slide is marked `carried` so the printer copies it instead.
 * That decision has to be made here, before any shape is mapped, since it is a property of
 * the slide as a whole.
 */
import type { Presentation } from '../../read/api/presentation.js'
import type { Slide } from '../../read/api/slide.js'
import { isGraphicFrame, isGroupShape, type AnyShape } from '../../read/api/shapes.js'
import { NoteCollector, scopeNotes, type NoteScope } from '../fidelity.js'
import type { AssetIr, AssetRef, BackgroundIr, CallIr, DeckIr, DeckPropsIr, SlideIr, SlideLayoutIr } from '../ir.js'
import { shapeCall, unwritableFramePayload, type AssetResolver } from './shape.js'
import { chromeToIr } from './chrome.js'
import { transitionToIr } from './transition.js'
import { compact, compactRequired, inches, literalColor } from './values.js'
import { STANDARD_LAYOUTS } from '../../units.js'
import { assetFilenameExtension } from '../../media/content-type.js'

/** Default slide size for a deck whose `presentation.xml` declares none: the 10in × 7.5in 4:3 layout. */
const DEFAULT_SLIDE_SIZE = {
	widthEmu: STANDARD_LAYOUTS.LAYOUT_4x3.widthEmu,
	heightEmu: STANDARD_LAYOUTS.LAYOUT_4x3.heightEmu,
}

/**
 * Collects media as shapes reference it, assigning each part a stable sequential name.
 *
 * Names are assigned in first-reference order rather than derived from the source part
 * name, so the output is deterministic for a given deck and does not leak whatever the
 * source package happened to call its files. A part referenced twice resolves to one asset.
 *
 * Numbering runs per *kind* (`image1.png`, `image2.png`, `audio1.wav`) rather than over one
 * shared counter. Not cosmetics: the asset name becomes a `const` identifier in the emitted
 * script and a filename in its asset directory, and a transition sound bound to `image7`
 * reads as a bug in the generated source every time anyone opens it.
 */
class Assets implements AssetResolver {
	readonly #byPartName = new Map<string, AssetRef>()
	readonly #assets: AssetIr[] = []
	readonly #counts = new Map<string, number>()

	constructor(private readonly pres: Presentation) {}

	assetFor(partName: string): AssetRef | null {
		const existing = this.#byPartName.get(partName)
		if (existing) return existing

		const part = this.pres.opc.part(partName)
		if (!part) return null

		const extension = assetFilenameExtension(part.contentType) ?? partName.split('.').pop() ?? 'bin'
		const kind = part.contentType.startsWith('audio/') ? 'audio' : 'image'
		const index = (this.#counts.get(kind) ?? 0) + 1
		this.#counts.set(kind, index)
		const name = `${kind}${index}.${extension}`
		const ref: AssetRef = { $asset: name }
		this.#byPartName.set(partName, ref)
		this.#assets.push({ name, contentType: part.contentType, bytes: part.bytes })
		return ref
	}

	contentTypeOf(partName: string): string | null {
		return this.pres.opc.part(partName)?.contentType ?? null
	}

	get assets(): AssetIr[] {
		return this.#assets
	}
}

/** Convert a loaded deck into its IR. */
export function readModelToIr(pres: Presentation): DeckIr {
	const collector = new NoteCollector()
	const assets = new Assets(pres)
	const size = pres.slideSize
	const deckScope = scopeNotes(collector, null)

	if (!size) {
		deckScope.note(
			'deck.slideSize',
			'approximated',
			'unsupported',
			`this deck declares no slide size, so the output uses the ${inches(DEFAULT_SLIDE_SIZE.widthEmu)}in × ${inches(DEFAULT_SLIDE_SIZE.heightEmu)}in default`
		)
	}

	const layouts = layoutGallery(pres)
	// Before the slides, so the chrome's deck-level notes lead the list the way the chrome
	// leads the deck. Assets are shared: a layout background image and a slide image that are
	// the same part resolve to one asset.
	const chrome = chromeToIr(pres, deckScope, assets)
	const slides = pres.slides.map((slide, index) => slideToIr(slide, index + 1, collector, assets, layouts))

	return {
		slideSize: size ? { widthEmu: size.widthEmu, heightEmu: size.heightEmu } : DEFAULT_SLIDE_SIZE,
		props: deckProps(pres, deckScope),
		chrome,
		slides,
		assets: assets.assets,
		fidelity: collector.notes,
	}
}

/**
 * Deck properties, reduced to the five `docProps` fields the write API sets. The other
 * seven the read model exposes have no setter — they survive anyway in a template-anchored
 * output, since the source deck itself becomes the template, so this is only a loss for a
 * standalone one.
 */
function deckProps(pres: Presentation, notes: NoteScope): DeckPropsIr {
	const core = pres.coreProperties
	const carried = ['keywords', 'description', 'category', 'contentStatus', 'lastModifiedBy'] as const
	if (carried.some((key) => core[key] !== undefined)) {
		notes.note(
			'deck.docProps',
			'dropped',
			'unwritable',
			'only title, author, company, subject and revision have write-API setters; keywords, description, category, content status and last-modified-by have none'
		)
	}
	return (
		compact({
			title: core.title,
			author: core.creator,
			subject: core.subject,
			revision: core.revision,
		}) ?? {}
	)
}

/**
 * The deck's layouts keyed by partname, each resolved to what a printer needs to bind to
 * it: its gallery position and whether its name identifies it on its own.
 *
 * Built once per deck rather than per slide — `layouts()` walks every master's layout list,
 * and the uniqueness question is about the gallery as a whole, so asking it per slide would
 * be both quadratic and unable to see layouts no slide happens to use.
 */
function layoutGallery(pres: Presentation): Map<string, SlideLayoutIr> {
	const handles = pres.layouts()
	const nameCounts = new Map<string, number>()
	for (const handle of handles) nameCounts.set(handle.name, (nameCounts.get(handle.name) ?? 0) + 1)

	return new Map(
		handles.map((handle, index) => [
			handle.partName,
			{ name: handle.name, index, nameIsUnique: (nameCounts.get(handle.name) ?? 0) === 1 },
		])
	)
}

function slideToIr(
	slide: Slide,
	number: number,
	collector: NoteCollector,
	assets: Assets,
	layouts: Map<string, SlideLayoutIr>
): SlideIr {
	const notes = scopeNotes(collector, number)
	const layoutPartName = slide.layout?.partName
	const base = {
		number,
		layout: (layoutPartName === undefined ? undefined : layouts.get(layoutPartName)) ?? null,
		hidden: slide.hidden,
		...(slide.name ? { name: slide.name } : {}),
		...backgroundOf(slide, notes),
		...notesOf(slide, notes),
	}

	const transition = transitionToIr(slide, notes, assets)
	const carried = hasUnwritableContent(slide)
	if (carried) {
		notes.note(
			'slide.carried',
			'flattened',
			'unwritable',
			'this slide holds a graphic frame the write API cannot author (an extended chart, a SmartArt diagram, an OLE object or ink); a printer that can copy the source slide does so instead of transcribing it, so it renders correctly but the emitted script does not describe its contents. The per-shape note beside this one says which construct forced the copy'
		)
	}

	// Mapped even for a carried slide: a printer with no source package to copy from prints
	// these and loses only the unwritable construct, rather than the whole slide. Each shape
	// records its own note, so that loss is declared either way.
	const calls: CallIr[] = []
	for (const shape of slide.shapes) {
		const call = shapeCall(shape, notes, assets)
		if (call) calls.push(call)
	}

	// Build animations are recorded once per slide rather than per shape, since they live in a
	// slide-scoped element with no shape to attribute them to.
	recordAnimationLoss(slide, notes)

	return {
		...base,
		...(transition ? { transition } : {}),
		source: carried ? 'carried' : 'authored',
		calls,
	}
}

/**
 * `true` when the slide holds something with no write-API expression at all, forcing a
 * carried slide. Descends into groups, since a group can host a graphic frame.
 *
 * Membership comes from {@link unwritableFramePayload}, which `graphicFrameCall` also uses to
 * pick its note — one enumeration, so the two cannot drift apart again. This test named
 * extended charts alone until then, which is why a slide holding SmartArt (or an OLE object,
 * or ink) was transcribed as `authored` and rendered with a hole where the graphic had been.
 */
function hasUnwritableContent(slide: Slide): boolean {
	const check = (shapes: AnyShape[]): boolean =>
		shapes.some((shape) => {
			if (isGroupShape(shape)) return check(shape.shapes)
			return isGraphicFrame(shape) && unwritableFramePayload(shape) !== null
		})
	return check(slide.shapes)
}

function backgroundOf(slide: Slide, notes: NoteScope): { background?: BackgroundIr } {
	const background = slide.background
	// An inherited background comes with the layout, so re-authoring it onto the slide would
	// pin a colour that should keep following the layout.
	if (!background || background.source !== 'slide') return {}

	switch (background.type) {
		case 'solid':
			// `compactRequired`: the guard has already established there is a colour, so the IR object
			// cannot come back empty, and `compact`'s "nothing survived" return does not apply here.
			return background.color
				? { background: compactRequired({ color: literalColor(background.color.effectiveHex) }) }
				: {}
		case 'none':
			return {}
		default:
			notes.note(
				'slide.background',
				'dropped',
				'unsupported',
				`a ${background.type} slide background is not expressible through the write API's background option, so the slide takes its layout's background`
			)
			return {}
	}
}

function notesOf(slide: Slide, notes: NoteScope): { notesText?: string } {
	const text = slide.notesText
	if (!text) return {}
	const frame = slide.notesTextFrame
	if (frame && frame.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.bold || run.italic))) {
		notes.note(
			'notes.formatting',
			'flattened',
			'unsupported',
			"speaker notes carry as plain text; per-run formatting and the notes slide's own placeholder geometry do not"
		)
	}
	return { notesText: text }
}

/**
 * Build animations, which live in a slide-scoped `p:timing` element rather than in any
 * shape's subtree.
 *
 * The transition next door used to be recorded here for the same reason and is now mapped
 * instead (`from-read/transition.ts`), which leaves this the one genuinely slide-scoped
 * loss: there is no structural reader for `p:timing` at all, only spid manipulation for the
 * deck-to-deck import path, so a build sequence is invisible to this converter.
 */
function recordAnimationLoss(slide: Slide, notes: NoteScope): void {
	if (slide.hasAnimations) {
		notes.note(
			'slide.animation',
			'dropped',
			'unread',
			'build animations have no structural reader (p:timing is only manipulated by spid for the deck-to-deck import path), so every shape on this slide lands static'
		)
	}
}
