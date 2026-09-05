/**
 * ts-pptx: the construct-family seam
 *
 * A construct family is one value: everything the library knows about charts, or tables, or
 * speaker notes, reachable through a single name. What that value carries is what the write path
 * needs handed to it rather than imported — today the methods a family adds to a slide and the
 * child descriptors it recognises inside a slide master or a group.
 *
 * It has to be a value rather than a set of imports for the same reason the shape walk is handed
 * its renderers (`gen/slide/objects/shared.ts`) and the packager its part contributors
 * (`package/parts/shared.ts`): a named import inside a reachable function body is retained
 * unconditionally, and a class method body is never shaken at all, so a `SlideBuilder` that calls
 * `addChartDefinition` itself links every chart emitter into every program that writes a slide.
 * Those two seams say what a family emits; this one says what a family *is*, and a presentation is
 * composed with a list of them.
 *
 * Composed per presentation, never registered at import time. A module-level registry populated by
 * import side effects is how the same input came to produce different bytes once already (the
 * chart-part-id counter, `package/assemble.ts`), it would stop two presentations composed
 * differently in one process from each getting their own families, and it would make
 * `"sideEffects": false` a lie.
 */

import type { CHART_NAME } from '../enums.js'
import type {
	ChartMulti,
	ChartOpts,
	MeasureTextOptions,
	OptsChartData,
	OverflowBoxOptions,
	SlideMasterObject,
	TableLayoutResult,
	TableProps,
	TableRow,
	TableToSlidesProps,
	TextMeasurement,
	TextProps,
} from '../types/index.js'
import type { Slide } from '../types/slide.js'
import type { PresSlideInternal } from '../types/internal.js'
import type { RendererTable } from '../gen/slide/objects/shared.js'
import type { PartContributor } from '../package/parts/shared.js'
import type { ExtractedSlide } from '../read/api/presentation-types.js'
import type SlideBuilder from '../slide.js'
import type PresentationCore from '../presentation.js'
import type { FontMetricsRegistry } from '../measure/font-metrics.js'
import { UnsupportedFeatureError } from '../errors.js'

/**
 * Every construct family, by name. A name is what a diagnostic can say about a family that is not
 * composed, so it exists even for the families every tier carries.
 */
export type FamilyName =
	| 'animations'
	| 'chart'
	| 'comments'
	| 'connector'
	| 'group'
	| 'image'
	| 'measure'
	| 'media'
	| 'model3d'
	| 'notes'
	| 'ole'
	| 'shape'
	| 'table'
	| 'text'
	| 'zoom'

/** The `Slide` methods that come from a family rather than from the slide itself. */
export type SlideAuthorMethod =
	| 'addAnimation'
	| 'addChart'
	| 'addComment'
	| 'addConnector'
	| 'addGroup'
	| 'addImage'
	| 'addMedia'
	| 'addModel3d'
	| 'addNotes'
	| 'addOleObject'
	| 'addSectionZoom'
	| 'addShape'
	| 'addSlideZoom'
	| 'addSummaryZoom'
	| 'addTable'
	| 'addText'
	| 'groupObjects'

/**
 * One family method's author: the public signature, with the slide it authors onto in front and
 * the return dropped. Chaining is not the author's business — `SlideBuilder` returns the slide
 * from every bound method, so an author that returned one could only return the same one.
 */
type Author<M extends SlideAuthorMethod> = (slide: SlideBuilder, ...args: Parameters<Slide[M]>) => void

/**
 * The slide methods a list of families supplies between them, each optional because a program
 * composed without a family does not have its methods.
 *
 * Derived from {@link Slide} rather than restated, so an author cannot drift from the method it
 * backs. `addChart` is written out because it is the one overloaded member: `Parameters` reads the
 * last overload only, which would let a family that handles just the combo form type-check while
 * the public method promises both.
 */
export type SlideAuthors = { [M in Exclude<SlideAuthorMethod, 'addChart'>]?: Author<M> } & {
	addChart?: (
		slide: SlideBuilder,
		arg1: OptsChartData[] | ChartMulti[],
		arg2?: ChartOpts & { type?: CHART_NAME }
	) => void
}

/** Which family supplies each slide method, as data, so a program can name a family it never linked. */
export const SLIDE_METHOD_FAMILIES: Readonly<Record<SlideAuthorMethod, FamilyName>> = Object.freeze({
	addAnimation: 'animations',
	addChart: 'chart',
	addComment: 'comments',
	addConnector: 'connector',
	addGroup: 'group',
	addImage: 'image',
	addMedia: 'media',
	addModel3d: 'model3d',
	addNotes: 'notes',
	addOleObject: 'ole',
	addSectionZoom: 'zoom',
	addShape: 'shape',
	addSlideZoom: 'zoom',
	addSummaryZoom: 'zoom',
	addTable: 'table',
	addText: 'text',
	groupObjects: 'group',
})

/**
 * What a presentation-level author is given: the deck it is authoring onto, and the write-side
 * font metrics, which the presentation keeps private and a family cannot reach through the class.
 */
export interface PresentationAuthorContext {
	readonly pres: PresentationCore
	readonly fontMetrics: FontMetricsRegistry
}

/**
 * The presentation methods that come from a family rather than from the presentation itself.
 *
 * Written out with their signatures rather than derived from the class, because the class is what
 * they are being lifted out of. Each returns its own answer, so unlike a slide author there is
 * nothing generic to apply on the way back.
 */
export interface PresentationAuthors {
	measureText?: (
		ctx: PresentationAuthorContext,
		text: string | TextProps[],
		opts: MeasureTextOptions
	) => TextMeasurement
	overflowsBox?: (ctx: PresentationAuthorContext, text: string | TextProps[], opts: OverflowBoxOptions) => boolean
	tableLayout?: (ctx: PresentationAuthorContext, rows: TableRow[], opts: TableProps) => TableLayoutResult
	tableToSlides?: (ctx: PresentationAuthorContext, eleId: string, options: TableToSlidesProps) => void
}

/** Which family supplies each presentation method. */
export const PRESENTATION_METHOD_FAMILIES: Readonly<Record<keyof PresentationAuthors, FamilyName>> = Object.freeze({
	measureText: 'measure',
	overflowsBox: 'measure',
	tableLayout: 'measure',
	tableToSlides: 'table',
})

/**
 * The key-tagged child descriptors a slide master's `objects` and a group's `children` are written
 * with (`{ text: … }`, `{ rect: … }`, …). `placeholder` is not one: it is master-specific and needs
 * the object's index, so `createSlideMaster` handles it itself.
 */
export type ChildDescriptorKey = 'chart' | 'image' | 'line' | 'rect' | 'roundRect' | 'shape' | 'text'

/**
 * Which family supplies each child descriptor, as data — the counterpart to
 * {@link SLIDE_METHOD_FAMILIES}, and for the same reason: a descriptor key that reaches a
 * presentation composed without its family has to be able to name the family it is missing.
 *
 * Four of the seven keys are the shape family's, because `line`, `rect` and `roundRect` are the
 * shorthands `shape` spells out.
 */
export const CHILD_DESCRIPTOR_FAMILIES: Readonly<Record<ChildDescriptorKey, FamilyName>> = Object.freeze({
	chart: 'chart',
	image: 'image',
	line: 'shape',
	rect: 'shape',
	roundRect: 'shape',
	shape: 'shape',
	text: 'text',
})

/** The payload behind one descriptor key, taken from the descriptor union so it cannot drift. */
type ChildDescriptor<K extends ChildDescriptorKey> = Extract<SlideMasterObject, Record<K, unknown>>[K]

/**
 * The child descriptors a list of families recognises. A descriptor whose key is missing is one
 * nothing composed can author, and both walks warn — naming the family behind the key when
 * {@link CHILD_DESCRIPTOR_FAMILIES} has one, so a `{ chart: … }` in a master's `objects` on a
 * chartless presentation says what to compose rather than disappearing from the deck.
 */
export type ChildAuthors = {
	[K in ChildDescriptorKey]?: (target: PresSlideInternal, child: ChildDescriptor<K>) => void
}

/**
 * What a family contributes to an *extraction* -- the second emission path, which serializes a
 * slide for injection into a deck read from bytes rather than for a package of its own.
 *
 * One slot so far. Extraction re-emits a slide's shapes through the same renderer table a write
 * uses, so a family with only shapes to draw needs nothing here; charts are the exception because
 * an extracted chart carries its part XML and its workbook alongside the slide body.
 */
export interface SlideExtractors {
	readonly charts?: (slide: PresSlideInternal) => ExtractedSlide['charts']
}

/** One construct family. Every slot is optional; a family fills the ones it has something to say about. */
export interface ConstructFamily {
	/** This family's name, for the diagnostic a program without it raises. */
	readonly name: FamilyName
	/** Methods this family adds to a slide. */
	readonly authors?: SlideAuthors
	/** Child descriptors this family recognises inside a slide master or a group. */
	readonly children?: ChildAuthors
	/** The shape kinds this family emits XML for, keyed by `SlideObjectType`. */
	readonly renderers?: RendererTable
	/** The parts this family puts in the written package. */
	readonly parts?: PartContributor
	/** Methods this family adds to the presentation itself. */
	readonly presentationAuthors?: PresentationAuthors
	/** What this family adds to an extracted slide, beyond the shape XML its renderer emits. */
	readonly extract?: SlideExtractors
}

/** What a presentation was composed with: every family's contribution, flattened per seam. */
export interface Composition {
	readonly authors: SlideAuthors
	readonly children: ChildAuthors
	readonly renderers: RendererTable
	readonly partContributors: readonly PartContributor[]
	readonly presentationAuthors: PresentationAuthors
	readonly extract: SlideExtractors
}

/**
 * Flatten a family list into the tables the write path reads.
 *
 * A family listed twice is composed once. The list's order does not reach the output. Authors, child descriptors and renderers are all
 * addressed by name, so two families can only collide by claiming the same one, and none do; the
 * part contributors come out in list order and the packager sorts them by the rank each one
 * declares, which is the ordering that is byte-significant.
 */
export function composeFamilies(families: readonly ConstructFamily[]): Composition {
	const authors: SlideAuthors = {}
	const children: ChildAuthors = {}
	const renderers: RendererTable = {}
	const partContributors: PartContributor[] = []
	const presentationAuthors: PresentationAuthors = {}
	const extract: SlideExtractors = {}
	// By identity, because a family is a singleton value: listing one twice is a thing a caller will
	// do (`use: [tables]` on a tier that already has tables), and a part contributor collected twice
	// would write its parts twice -- a deck PowerPoint reports as corrupt, from a harmless spelling.
	for (const family of new Set(families)) {
		Object.assign(authors, family.authors)
		Object.assign(children, family.children)
		Object.assign(renderers, family.renderers)
		Object.assign(presentationAuthors, family.presentationAuthors)
		Object.assign(extract, family.extract)
		if (family.parts) partContributors.push(family.parts)
	}
	return { authors, children, renderers, partContributors, presentationAuthors, extract }
}

/**
 * The failure a method raises when its family was not composed.
 *
 * A method that is not composed still exists, and says which family it needs. Leaving it off
 * instead would report the same condition as `TypeError: slide.addChart is not a function`, which
 * names neither the family nor the fix; the family list this program was built with is not
 * something a missing property could point at.
 * @param call - how the method is called, e.g. `slide.addChart`
 * @param family - the family that supplies it
 */
export function familyMethodUnavailable(call: string, family: FamilyName): never {
	throw new UnsupportedFeatureError(
		'family/not-composed',
		`${call}() needs the "${family}" construct family, and this presentation was composed without it. ` +
			`Add the "${family}" family to the list the presentation is composed with.`
	)
}

/**
 * A presentation author with its context already supplied — what a method on the presentation
 * calls, having nothing left to pass but its own arguments.
 */
export type BoundPresentationAuthor<F> = F extends (ctx: PresentationAuthorContext, ...args: infer A) => infer R
	? (...args: A) => R
	: never

/** One presentation method's author, or the failure that names the family it came from. */
export function requirePresentationAuthor<M extends keyof PresentationAuthors>(
	authors: PresentationAuthors,
	method: M
): NonNullable<PresentationAuthors[M]> {
	const author = authors[method]
	if (!author) familyMethodUnavailable(`pptx.${method}`, PRESENTATION_METHOD_FAMILIES[method])
	return author
}
