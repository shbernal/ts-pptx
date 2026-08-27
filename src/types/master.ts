/**
 * Slide-master types: master object descriptors, bullet configuration and per-level text styles.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { CHART_NAME } from '../enums.js'
import type { ChartMulti, ChartOpts, OptsChartData } from './chart.js'
import type { BackgroundProps, Color, HAlign, Margin, PositionProps } from './core.js'
import type { CommonObjectDescriptor, PlaceholderProps } from './object.js'
import type { TextBaseProps } from './text.js'

export interface SlideNumberProps extends PositionProps, TextBaseProps {
	/**
	 * margin (inches) — text-frame internal margin; a value `>= 1` warns as a likely legacy points value
	 */
	margin?: Margin
}
export interface SlideMasterChartProps {
	type: CHART_NAME | ChartMulti[]
	data: OptsChartData[]
	options?: ChartOpts
	opts?: ChartOpts
}
/**
 * One object on a slide master, as a key-tagged descriptor.
 *
 * The six descriptors it shares with a group child are {@link CommonObjectDescriptor}; the two
 * below are the master's own — a chart, which a group cannot hold, and a placeholder, which is
 * the whole point of a master.
 */
export type SlideMasterObject =
	| CommonObjectDescriptor
	| { chart: SlideMasterChartProps }
	| {
			placeholder: {
				options: PlaceholderProps
				/**
				 * Text to be shown in placeholder (shown until user focuses textbox or adds text)
				 * - Leave blank to have powerpoint show default phrase (ex: "Click to add title")
				 */
				text?: string
			}
	  }
/**
 * Bullet configuration for one slide-master text-style level (`a:buChar`/`a:buAutoNum`/`a:buNone`).
 * A focused subset of the run-level bullet model — picture bullets and per-glyph sizing are not
 * supported in master defaults (those require slide rels). Set `bullet: false` to suppress the
 * level's bullet entirely.
 */
export interface MasterBulletProps {
	/** `'bullet'` emits a character bullet (`a:buChar`); `'number'` emits an auto-number (`a:buAutoNum`). @default 'bullet' */
	type?: 'bullet' | 'number'
	/** Bullet character (unicode code point hex), e.g. `'2022'` for •. Used when `type` is `'bullet'`. */
	characterCode?: string
	/** Glyph font typeface for the bullet character (`a:buFont`), e.g. `'Arial'` or `'Wingdings'`. */
	fontFace?: string
	/** Auto-number scheme (`a:buAutoNum@type`) when `type` is `'number'`, e.g. `'arabicPeriod'`. @default 'arabicPeriod' */
	numberType?: string
	/** Starting value for auto-numbered bullets (`a:buAutoNum@startAt`). @default 1 */
	numberStartAt?: number
}
/**
 * Styling for one paragraph level (`a:lvlNpPr`) of a slide-master text style. Every field is
 * optional; an unset field keeps PowerPoint's built-in default for that level. Configure via
 * {@link SlideMasterProps.textStyles}.
 */
export interface MasterTextStyleLevel {
	/** Font size in points (`a:defRPr@sz`), e.g. `24`. */
	fontSize?: number
	/** Font face (`a:defRPr/a:latin@typeface`). Unset keeps the theme font (`+mn-lt`/`+mj-lt`). */
	fontFace?: string
	/** Text color — hex (`'C00000'`) or theme slot (`'tx1'`) — emitted as `a:defRPr/a:solidFill`. */
	color?: Color
	/** Bold text (`a:defRPr@b`). */
	bold?: boolean
	/** Italic text (`a:defRPr@i`). */
	italic?: boolean
	/** Horizontal alignment (`a:lvlNpPr@algn`). */
	align?: HAlign
	/** Left margin in inches (`a:lvlNpPr@marL`); for a hanging bullet this is the text indent. */
	marginLeft?: number
	/** First-line indent in inches (`a:lvlNpPr@indent`); negative produces a hanging bullet. */
	indent?: number
	/** Bullet config: `false` emits `a:buNone`; an object configures a char/number bullet. Unset keeps the level default. */
	bullet?: boolean | MasterBulletProps
}
/**
 * Per-level slide-master text styles, written to `slideMaster1.xml`'s `<p:txStyles>`.
 * Because a deck has a single shared slide master, these styles are **deck-wide**: when set on more
 * than one `defineSlideMaster()` call the last value for each group (`title`/`body`/`other`) wins.
 */
export interface MasterTextStyleProps {
	/** Title placeholder style (`p:titleStyle`, single level). */
	title?: MasterTextStyleLevel
	/** Body placeholder per-level styles (`p:bodyStyle`, levels 1–9; index 0 is `lvl1`). */
	body?: MasterTextStyleLevel[]
	/** Other/default placeholder per-level styles (`p:otherStyle`, levels 1–9). */
	other?: MasterTextStyleLevel[]
}
export interface SlideMasterProps {
	/**
	 * Unique name for this master
	 */
	title: string
	background?: BackgroundProps
	margin?: Margin
	slideNumber?: SlideNumberProps
	objects?: SlideMasterObject[]
	/**
	 * Per-level master text styles (title / body / other) written to the shared slide master's
	 * `<p:txStyles>`. Configure nested bullet character, font size, color, alignment, and indent
	 * for each of the nine list levels. Deck-wide (see {@link MasterTextStyleProps}).
	 * @example textStyles: { body: [{ fontSize: 24, color: 'C00000', bullet: { characterCode: '25AA' } }] }
	 */
	textStyles?: MasterTextStyleProps
}
