/**
 * ts-pptx: composing a presentation from the construct families it actually needs
 *
 * The body behind each entry's `createPresentation`, and the types that describe what comes back.
 * It lives beside `entry-surface.ts` for the same reason that does: `index.ts`, `node.ts` and
 * `browser.ts` differ only in the runtime adapter they hand down, so the shared half is written
 * once and cannot drift between them.
 *
 * There is no second lean entry point, deliberately. A `pptx-ts/core` would need the same
 * three-condition treatment `.` gets, so three more source entries, three more `dist` files, three
 * more budget rows and a second place for the entries to diverge -- and it would buy nothing that
 * export-level shaking does not already give. `pptx-ts` is an intermediate module in a consumer's
 * graph: a program that imports only `createPresentation` drops the `TsPptx` class beside it and,
 * with it, every family that class names.
 */

import PresentationCore from './presentation.js'
import type { RuntimeAdapter } from './runtime/types.js'
import { CORE_CONSTRUCT_FAMILIES } from './families/core.js'
import type {
	BoundPresentationAuthor,
	ConstructFamily,
	PresentationAuthors,
	SlideAuthorMethod,
} from './families/shared.js'
import type { CHART_NAME } from './enums.js'
import type { AddSlideProps, ChartMulti, ChartOpts, OptsChartData } from './types/index.js'
import type { Slide } from './types/slide.js'

/** Turn a union into the intersection of its members. */
type UnionToIntersection<U> = (U extends unknown ? (k: U) => void : never) extends (k: infer I) => void ? I : never

/** One method's public signature, re-pointed at the slide type the caller actually holds. */
type Chained<F, S> = F extends (...args: infer A) => unknown ? (...args: A) => S : never

/**
 * The overloaded methods, re-pointed by hand.
 *
 * `Chained` reads one call signature, and TypeScript hands a conditional type the *last* overload,
 * so `addChart` would come out taking a combo array and nothing else. There is one such method, and
 * writing it out is cheaper than a generic that unrolls an arbitrary overload list.
 */
interface ChainedOverloads<S> {
	addChart: {
		(data: OptsChartData[], options: ChartOpts & { type: CHART_NAME }): S
		(charts: ChartMulti[], options?: ChartOpts): S
	}
}

/** The slide methods a family list supplies between them. */
export type SlideMethodsOf<Fs extends readonly ConstructFamily[]> = Extract<
	keyof UnionToIntersection<Extract<Fs[number], { authors: object }>['authors']>,
	SlideAuthorMethod
>

/** The presentation methods a family list supplies between them. */
export type PresentationMethodsOf<Fs extends readonly ConstructFamily[]> = Extract<
	keyof UnionToIntersection<Extract<Fs[number], { presentationAuthors: object }>['presentationAuthors']>,
	keyof PresentationAuthors
>

/**
 * A slide from a composed presentation: everything a slide is, plus exactly the authoring methods
 * `M` names.
 *
 * Parameterized by method name rather than by the family list, so what an editor shows a caller is
 * the list of calls this slide takes. The methods answer this same type, so a chain cannot widen
 * back to the full surface and offer a call the deck was not composed for.
 */
export type ComposedSlide<M extends SlideAuthorMethod> = Omit<Slide, SlideAuthorMethod> & {
	[K in M]: K extends keyof ChainedOverloads<unknown>
		? ChainedOverloads<ComposedSlide<M>>[K]
		: Chained<Slide[K], ComposedSlide<M>>
}

/** Every presentation method a family can supply, as the presentation exposes it. */
type PresentationFamilyMethods = {
	[K in keyof PresentationAuthors]-?: BoundPresentationAuthor<NonNullable<PresentationAuthors[K]>>
}

/**
 * A composed presentation: the whole deck-level API, with `addSlide` answering a slide that carries
 * only `M`, and only the presentation-level methods `P` names.
 */
export type ComposedPresentation<M extends SlideAuthorMethod, P extends keyof PresentationAuthors> = Omit<
	PresentationCore,
	'addSlide' | keyof PresentationAuthors
> &
	Pick<PresentationFamilyMethods, P> & { addSlide: (options?: AddSlideProps) => ComposedSlide<M> }

/** What `createPresentation` takes: the families this deck needs beyond the core tier. */
export interface ComposeOptions<Fs extends readonly ConstructFamily[]> {
	/** Construct families to compose in, from `pptx-ts/families`. Listing one twice is harmless. */
	use?: Fs
}

/** What a family list composes to once the core tier is in front of it. */
export type Composed<Fs extends readonly ConstructFamily[]> = ComposedPresentation<
	SlideMethodsOf<[...typeof CORE_CONSTRUCT_FAMILIES, ...Fs]>,
	PresentationMethodsOf<[...typeof CORE_CONSTRUCT_FAMILIES, ...Fs]>
>

/**
 * Build a presentation that carries the core tier plus `opts.use`, on the given runtime adapter.
 *
 * The cast is the one place this file is not type-safe, and it is unavoidable: the class is one
 * class whatever it was composed with, while the type says which of its methods are actually
 * there. The two are kept honest by the runtime, which raises `family/not-composed` naming the
 * family for any method the type has already refused.
 */
export function composePresentation<const Fs extends readonly ConstructFamily[]>(
	runtime: RuntimeAdapter,
	opts?: ComposeOptions<Fs>
): Composed<Fs> {
	return new PresentationCore(runtime, [...CORE_CONSTRUCT_FAMILIES, ...(opts?.use ?? [])]) as unknown as Composed<Fs>
}
