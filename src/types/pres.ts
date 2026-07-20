/**
 * Presentation-level types: write/export options, sections, the slide layout size, and
 * `PresentationProps`.
 *
 * Re-exported by `../core-interfaces.js`, which is the import site for the rest of `src/`.
 */
import type { WRITE_OUTPUT_TYPE } from '../core-enums.js'
import type { EmbeddedFont } from '../embedded-fonts.js'
import type { PresSlide, PresSlideInternal, SlideLayoutInternal } from './slide.js'
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
	 * - `'placeholder'`: substitute a broken-image placeholder, emit a `console.warn`, and
	 *   continue. Useful for best-effort/batch jobs where one missing asset should not abort
	 *   the whole deck.
	 * @default 'throw'
	 */
	onMediaError?: 'throw' | 'placeholder'
}
export interface WriteProps extends WriteBaseProps {
	/**
	 * Output type
	 * - values: 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array' | 'STREAM'
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
export interface SectionProps {
	/**
	 * Section title
	 */
	title: string
	/**
	 * Section order - uses to add section at any index
	 * - values: 1-n
	 */
	order?: number
}
export interface SectionInternalProps extends SectionProps {
	_type?: 'user' | 'default'
	_slides: PresSlideInternal[]
}
export interface PresLayout {
	_sizeW?: number
	_sizeH?: number

	/**
	 * Layout Name
	 * @example 'LAYOUT_WIDE'
	 */
	name: string
	width: number
	height: number
}
export type CustomPropertyValue = string | number | boolean | Date

export interface PresentationProps {
	author: string
	company: string
	layout: string
	masterSlide: PresSlide
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
// PRIVATE interface
export interface PresentationPropsInternal extends PresentationProps {
	masterSlide: PresSlideInternal
	sections: SectionInternalProps[]
	slideLayouts: SlideLayoutInternal[]
	slides: PresSlideInternal[]
	/** Author-side embedded fonts (see {@link PptxGenJS.embedFont}); empty when none. */
	embeddedFonts: EmbeddedFont[]
}
