/**
 * Presentation-level types: write/export options, sections, the slide layout size, and
 * `PresentationProps`.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { WRITE_OUTPUT_TYPE } from '../enums.js'
import type { Slide } from './slide.js'
import type { ThemeProps } from './theme.js'

export interface WriteBaseProps {
	/**
	 * Whether to DEFLATE-compress the package (PowerPoint itself always compresses;
	 * set `false` only if export time matters more than file size)
	 * @default true
	 */
	compression?: boolean
	/**
	 * How to handle a media asset (image/audio/video) that fails to load during export.
	 * - `'throw'` (default): reject the export with an error naming the failing asset. A deck
	 *   that silently embeds a broken-image placeholder is a degenerate result, so failing
	 *   loudly is the safe default.
	 * - `'placeholder'`: substitute a broken-image placeholder, emit a `media/load-failed`
	 *   diagnostic (see `setDiagnosticHandler`), and
	 *   continue. Useful for best-effort/batch jobs where one missing asset should not abort
	 *   the whole deck.
	 * @default 'throw'
	 */
	onMediaError?: 'throw' | 'placeholder'
}
export interface WriteProps extends WriteBaseProps {
	/**
	 * Output type
	 * - values: 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	 * @default 'blob'
	 */
	outputType?: WRITE_OUTPUT_TYPE
}
export interface WriteFileProps extends WriteBaseProps {
	/**
	 * Export file name
	 * @default 'Presentation.pptx'
	 */
	fileName?: string
}
/**
 * One emitted OOXML package part: its full slash-path and already-encoded bytes. Returned by
 * {@link TsPptx.toParts} for callers that want the raw parts of the `.pptx` OPC package without
 * zipping (custom containers, streaming, part-level inspection). The bytes are byte-identical to
 * what `write()` would compress for the same part;
 * XML parts are UTF-8 (decode with `new TextDecoder().decode(part.data)`), media/font parts are
 * their raw binary.
 */
export interface PackagePart {
	/** Full OOXML package part path, e.g. `ppt/slides/slide1.xml`. */
	readonly path: string
	/** The part's already-encoded bytes: a fresh view per call, safe to keep or transfer. */
	readonly data: Uint8Array
}
/**
 * Options for {@link TsPptx.toParts}. A deliberately narrow subset of {@link WriteProps}:
 * `compression` and `outputType` are zip concerns and do not apply when parts are returned
 * unzipped, so only media-error handling is exposed.
 */
export interface PartsProps {
	/**
	 * How to handle a media asset (image/audio/video) that fails to load. Same semantics as
	 * {@link WriteBaseProps.onMediaError}.
	 * @default 'throw'
	 */
	onMediaError?: 'throw' | 'placeholder'
}
export interface SectionProps {
	/**
	 * Section title
	 */
	title: string
	/**
	 * Where to insert the section, counting from **1**: `order: 1` puts it first, `order: 2`
	 * second. Omit it to append.
	 *
	 * An order past the end appends, which is the only sensible reading of "put it at
	 * position 12" in a deck with four sections. Anything that is not a whole number of at
	 * least 1 -- `0`, a negative, a fraction -- warns and appends: it named no position, and
	 * silently treating it as one is how `order: 1` came to insert *second* and `order: 0` to
	 * append with no word said.
	 * - values: 1-n
	 */
	order?: number
}
export interface PresLayout {
	/**
	 * Layout Name
	 * @example 'LAYOUT_WIDE'
	 */
	name: string
	/**
	 * Slide width.
	 *
	 * **The unit depends on which direction this shape is travelling.** `defineLayout` reads it
	 * as INCHES -- that is what its own example passes -- while `pptx.presLayout` returns it in
	 * EMU. So `defineLayout({ ...pptx.presLayout, name: 'Copy' })`, the obvious way to derive a
	 * layout from the current one, states a width of nine million inches; `defineLayout` now
	 * clamps it to the 56in maximum and says so rather than emitting a `sldSz` PowerPoint
	 * refuses.
	 */
	width: number
	/** Slide height; the same two units as {@link width}, in the same two directions. */
	height: number
}
export type CustomPropertyValue = string | number | boolean | Date

export interface PresentationProps {
	author: string
	company: string
	layout: string
	masterSlide: Slide
	/**
	 * Presentation's layout
	 * read-only
	 */
	presLayout: PresLayout
	revision: string
	/**
	 * Slide number to assign to the first slide (affects the slide-number field displayed in placeholders).
	 * @default 1
	 */
	firstSlideNum: number
	/**
	 * Whether to enable right-to-left mode
	 * @default false
	 */
	rtlMode: boolean
	subject: string
	theme: ThemeProps
	title: string
}
