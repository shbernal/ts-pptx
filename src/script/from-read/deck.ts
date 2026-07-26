/**
 * `Presentation` → {@link DeckIr}. The entry point of the read half.
 *
 * Pure in the sense that matters: it reads a loaded package and returns data, touching no
 * filesystem and mutating nothing. Given the same bytes it returns the same IR, so a
 * round-trip check can compare two runs directly.
 *
 * The one judgement made at this level rather than per-shape is {@link SlideSource}. A
 * slide holding a construct the write API cannot express *at all* — an extended chart is
 * the clear case, with a complete reader and no emitter — cannot be honestly transcribed,
 * because the output would silently differ from the source. Such a slide is marked
 * `carried` so the printer copies it instead. That decision has to be made here, before
 * any shape is mapped, since it is a property of the slide as a whole.
 */
import type { Presentation } from '../../read/api/presentation.js'
import type { Slide } from '../../read/api/slide.js'
import type { AnyShape } from '../../read/api/shapes.js'
import { isGraphicFrame, isGroupShape } from '../../read/api/shapes.js'
import { NoteCollector, scopeNotes, type NoteScope } from '../fidelity.js'
import type { AssetIr, AssetRef, BackgroundIr, CallIr, DeckIr, DeckPropsIr, SlideIr } from '../ir.js'
import type { AssetResolver } from './shape.js'
import { shapeCall } from './shape.js'
import { compact } from './values.js'

/** Default slide size (10" × 7.5") for a deck whose `presentation.xml` declares none. */
const DEFAULT_SLIDE_SIZE = { widthEmu: 9144000, heightEmu: 6858000 }

/** Extensions by content type, so an emitted asset filename is one a viewer recognises. */
const ASSET_EXTENSIONS: Record<string, string> = {
	'image/png': 'png',
	'image/jpeg': 'jpg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/webp': 'webp',
	'image/tiff': 'tif',
	'image/x-emf': 'emf',
	'image/x-wmf': 'wmf',
	'image/svg+xml': 'svg',
}

/**
 * Collects media as shapes reference it, assigning each part a stable sequential name.
 *
 * Names are assigned in first-reference order rather than derived from the source part
 * name, so the output is deterministic for a given deck and does not leak whatever the
 * source package happened to call its files. A part referenced twice resolves to one asset.
 */
class Assets implements AssetResolver {
	readonly #byPartName = new Map<string, AssetRef>()
	readonly #assets: AssetIr[] = []

	constructor(private readonly pres: Presentation) {}

	assetFor(partName: string): AssetRef | null {
		const existing = this.#byPartName.get(partName)
		if (existing) return existing

		const part = this.pres.opc.part(partName)
		if (!part) return null

		const extension = ASSET_EXTENSIONS[part.contentType] ?? partName.split('.').pop() ?? 'bin'
		const name = `image${this.#assets.length + 1}.${extension}`
		const ref: AssetRef = { $asset: name }
		this.#byPartName.set(partName, ref)
		this.#assets.push({ name, contentType: part.contentType, bytes: part.bytes })
		return ref
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
			`this deck declares no slide size, so the output uses the ${DEFAULT_SLIDE_SIZE.widthEmu / 914400}in × ${DEFAULT_SLIDE_SIZE.heightEmu / 914400}in default`
		)
	}

	const slides = pres.slides.map((slide, index) => slideToIr(slide, index + 1, collector, assets))

	return {
		slideSize: size ? { widthEmu: size.widthEmu, heightEmu: size.heightEmu } : DEFAULT_SLIDE_SIZE,
		props: deckProps(pres, deckScope),
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

function slideToIr(slide: Slide, number: number, collector: NoteCollector, assets: Assets): SlideIr {
	const notes = scopeNotes(collector, number)
	const base = {
		number,
		layoutName: slide.layout?.name ?? null,
		hidden: slide.hidden,
		...(slide.name ? { name: slide.name } : {}),
		...backgroundOf(slide, notes),
		...notesOf(slide, notes),
	}

	if (hasUnwritableContent(slide)) {
		notes.note(
			'slide.carried',
			'flattened',
			'unwritable',
			'this slide holds an extended chart (waterfall, funnel, box-and-whisker, …), which has a full reader but no write-API emitter; the slide is copied from the source deck instead of transcribed, so it renders correctly but the emitted script does not describe its contents'
		)
		return { ...base, source: 'carried', calls: [] }
	}

	const calls: CallIr[] = []
	for (const shape of slide.shapes) {
		const call = shapeCall(shape, notes, assets)
		if (call) calls.push(call)
	}

	// Transitions and build animations are recorded once per slide rather than per shape,
	// since both live in slide-scoped elements.
	recordTimingLosses(slide, notes)

	return { ...base, source: 'authored', calls }
}

/**
 * `true` when the slide holds something with no write-API expression at all, forcing a
 * carried slide. Descends into groups, since a group can host a graphic frame.
 */
function hasUnwritableContent(slide: Slide): boolean {
	const check = (shapes: AnyShape[]): boolean =>
		shapes.some((shape) => {
			if (isGroupShape(shape)) return check(shape.shapes)
			return isGraphicFrame(shape) && shape.hasChartEx
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
			return background.color
				? { background: compact({ color: background.color.effectiveHex.replace(/^#/, '').toUpperCase() }) }
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
 * Transitions and build animations, both of which live in slide-scoped elements rather than
 * in any shape's subtree.
 *
 * The transition case is the read path being *ahead* of the write path: `TransitionInfo`
 * decodes the `p14`/`p15` modern transitions by namespace, while the write vocabulary names
 * only 21. The animation case is the opposite — there is no structural reader at all, only
 * spid manipulation for the deck-to-deck import path, so a build sequence is invisible here.
 */
function recordTimingLosses(slide: Slide, notes: NoteScope): void {
	if (slide.transition) {
		notes.note(
			'slide.transition',
			'dropped',
			'unwritable',
			'a slide transition is read but appendSlides authors none on the destination slide, so the output advances with no effect'
		)
	}
	if (slide.hasAnimations) {
		notes.note(
			'slide.animation',
			'dropped',
			'unread',
			'build animations have no structural reader (p:timing is only manipulated by spid for the deck-to-deck import path), so every shape on this slide lands static'
		)
	}
}
