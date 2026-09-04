/**
 * Text types: the `TextBaseProps` paragraph/run contract shared by text, table cells and charts;
 * `TextPropsOptions`/`TextProps`; measurement and layout results; notes and comments.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { SHAPE_NAME } from '../enums.js'
import type {
	Color,
	DataOrPathProps,
	GeometryPoint,
	HAlign,
	HexColor,
	Margin,
	PositionProps,
	TextVertType,
	VAlign,
} from './core.js'
import type { TextShapeType } from '../ooxml/st-enums.js'
import type { ShapeAdjustValue } from './shape.js'
import type { ObjectNameProps } from './object.js'
import type { FillOption, HyperlinkProps, ShadowProps, ShapeLineProps } from './style.js'

/**
 * A drawn bullet: the object arm of {@link TextPropsOptions.bullet}.
 *
 * Named rather than left inline so a caller can build one as a variable and so the two
 * relationship ids `addText()` stamps on it have somewhere off the public surface to live
 * (`TextBulletPropsInternal`).
 */
export interface TextBulletProps {
	/**
	 * Bullet type
	 * @default bullet
	 */
	type?: 'bullet' | 'number'
	/**
	 * Bullet character code (unicode)
	 * @example '25BA' // 'BLACK RIGHT-POINTING POINTER' (U+25BA)
	 */
	characterCode?: string
	/**
	 * Bullet glyph font typeface (`<a:buFont/>`), e.g. for symbol-font bullets
	 * @example 'Wingdings' // render `characterCode` using the Wingdings font
	 */
	fontFace?: string
	/**
	 * Bullet glyph size as a percentage of the run's text size (25–400)
	 * @default 100
	 * @example 80 // bullet glyph is 80% of the text size
	 */
	size?: number
	/**
	 * Indentation (space between bullet and text) (points)
	 * @default 27 // DEF_BULLET_MARGIN
	 * @example 10 // Indents text 10 points from bullet
	 */
	indent?: number
	/**
	 * Number type
	 * @example 'romanLcParenR' // roman numerals lower-case with paranthesis right
	 */
	numberType?:
		| 'alphaLcParenBoth'
		| 'alphaLcParenR'
		| 'alphaLcPeriod'
		| 'alphaUcParenBoth'
		| 'alphaUcParenR'
		| 'alphaUcPeriod'
		| 'arabicParenBoth'
		| 'arabicParenR'
		| 'arabicPeriod'
		| 'arabicPlain'
		| 'romanLcParenBoth'
		| 'romanLcParenR'
		| 'romanLcPeriod'
		| 'romanUcParenBoth'
		| 'romanUcParenR'
		| 'romanUcPeriod'
	/**
	 * Number bullets start at
	 * @default 1
	 * @example 10 // numbered bullets start with 10
	 */
	numberStartAt?: number
	/**
	 * Image to use as the bullet glyph ("picture bullet", `<a:buBlip>`)
	 * - supply an image `path` (filesystem/URL) or base64 `data` (same forms as `addImage()`)
	 * - raster formats (PNG/JPG/GIF) and SVG are supported; use `size` to scale relative to the text height
	 * - SVG bullets embed a PNG preview plus the SVG (the same dual-rel handling as `addImage()`)
	 * - takes precedence over `type`/`characterCode` when set
	 * @example image: { path: 'images/star.png' }
	 * @example image: { data: 'image/png;base64,iVBOR...' }
	 * @example image: { path: 'images/star.svg' }
	 */
	image?: { path?: string; data?: string }
	/**
	 * Bullet glyph color (separate from the text run color)
	 * @example 'FF0000' // red bullet
	 */
	color?: HexColor
}

// used by: chart, slide, table, text
export interface TextBaseProps {
	/**
	 * Horizontal alignment
	 *
	 * When a shape's text is an array of runs, set this on **every run** of a paragraph.
	 * {@link breakLine} is not the only paragraph boundary: two adjacent runs whose `align`
	 * differs start a new one, so stating it on the opening run alone splits the paragraph in
	 * two. That is the mirror of {@link bullet}, {@link paraMarginLeft} and {@link paraIndent},
	 * which are read from the opening run only. Neither mistake changes the run count — the
	 * only symptom is a differently shaped paragraph, with no error.
	 * @default 'left'
	 */
	align?: HAlign
	/**
	 * Bold style
	 * @default false
	 */
	bold?: boolean
	/**
	 * Add a line-break
	 * @default false
	 */
	breakLine?: boolean
	/**
	 * Preset text warp / WordArt shape (`<a:bodyPr><a:prstTxWarp prst="..">`), which
	 * bends the text along a preset path (arch, circle, wave, …) — the whole
	 * `ST_TextShapeType` set.
	 * @example 'textArchUp' // bend text along an upward arch (e.g. a label following a ring/arc)
	 * @example 'textCircle'
	 */
	textWarp?: TextShapeType
	/**
	 * Add standard or custom bullet
	 * - use `true` for standard bullet
	 * - pass object options for custom bullet
	 * - `false` (and omitting the option) is the *explicit off*: it writes `<a:buNone/>`
	 *   plus `marL="0" indent="0"`, which overrides whatever bullet the layout's or
	 *   master's list style sets for this level
	 * - `'inherit'` states nothing at all — no bullet child and no margin attributes — so
	 *   the list style keeps reaching the paragraph. This is the one state omission
	 *   cannot spell, because omitting the option means `false` here for compatibility
	 *
	 * The margins are only a *default* of the bullet state: {@link paraMarginLeft} and
	 * {@link paraIndent} state `@marL`/`@indent` independently in any of the three, including
	 * an inherited margin under a drawn bullet.
	 *
	 * When a shape's text is an array of runs, state this on the **opening run only**. A run
	 * that draws a bullet starts a new paragraph by itself, so repeating it — the placement
	 * that {@link align} requires — turns one three-run paragraph into three one-run
	 * paragraphs. A paragraph takes its properties from its first run.
	 * @default false
	 */
	bullet?: boolean | 'inherit' | TextBulletProps
	/**
	 * Text capitalization (`a:rPr/@cap`, `ST_TextCapsType`)
	 * - `'all'` = ALL CAPS
	 * - `'small'` = Small Caps
	 * - `'none'` = the *explicit off*: it writes `cap="none"`, which overrides a capitalization
	 *   the run would otherwise inherit. Omitting the option states nothing and lets that
	 *   inheritance stand
	 * - PowerPoint: Font > Effects > All Caps / Small Caps
	 * @default (unset) inherit
	 */
	caps?: 'none' | 'small' | 'all'
	/**
	 * Text color
	 * - `HexColor` or `ThemeColor`
	 * - MS-PPT > Format Shape > Text Options > Text Fill & Outline > Text Fill > Color
	 * @example 'FF0000' // hex color (red)
	 * @example SchemeColor.text1 // Theme color (Text1)
	 */
	color?: Color
	/**
	 * Font face name
	 *
	 * Applied to the Latin (`<a:latin>`) and complex-script (`<a:cs>`) font slots, matching
	 * how PowerPoint writes a font picked from the UI. The East Asian slot (`<a:ea>`) is left
	 * to inherit from the theme unless `fontFaceEA` is set — forcing a Latin-only face into the
	 * East Asian slot duplicates/ghosts text in Office 365.
	 * @example 'Arial' // Arial font
	 */
	fontFace?: string
	/**
	 * East Asian font face name (`<a:ea>` slot), used to render CJK (Chinese/Japanese/Korean) glyphs
	 *
	 * Set this when the East Asian font differs from `fontFace`. When omitted, `<a:ea>` inherits the
	 * theme East Asian font, which is what PowerPoint does for Latin fonts.
	 * @example '微軟正黑體' // render East Asian glyphs with Microsoft JhengHei
	 */
	fontFaceEA?: string
	/**
	 * Font size
	 * @example 12 // Font size 12
	 */
	fontSize?: number
	/**
	 * Text highlight color (hex format)
	 * @example 'FFFF00' // yellow
	 */
	highlight?: HexColor
	/**
	 * italic style
	 * @default false
	 */
	italic?: boolean
	/**
	 * language
	 * - ISO 639-1 standard language code
	 * @default 'en-US' // english US
	 * @example 'fr-CA' // french Canadian
	 */
	lang?: string
	/**
	 * First-line indent (points) — `a:pPr/@indent`, the offset of the paragraph's FIRST line
	 * from {@link paraMarginLeft}. Negative hangs the first line to the left of the body, which
	 * is what a bulleted paragraph does; positive indents it to the right, the "first line
	 * indented" prose form. PowerPoint: Paragraph > Indentation > Special.
	 * - `'inherit'` writes no `@indent` at all, so the paragraph takes whatever its list style
	 *   (`a:lstStyle` → placeholder → layout → master) sets — the same third state
	 *   {@link bullet} spells, on the attribute beside it
	 * - omitting the option keeps the bullet-derived default: a drawn bullet hangs the first
	 *   line by its margin, `bullet: false` writes `indent="0"`, `bullet: 'inherit'` writes
	 *   nothing
	 * - a value here overrides all of those, in every bullet state
	 * @example -18 // hang the first line 18pt left of the body text
	 * @example 18 // indent the first line 18pt, prose style
	 * @example 'inherit' // keep the list style's indent, even on a bulleted paragraph
	 * @remarks Read from the **opening run** of a paragraph, like {@link bullet} and unlike
	 * {@link align}; a value on a continuation run is ignored.
	 */
	paraIndent?: number | 'inherit'
	/**
	 * Left margin of the paragraph (points) — `a:pPr/@marL`, where the paragraph's body text
	 * starts. This is the paragraph's own margin, not the text frame's internal padding
	 * (`margin`) and not the discrete outline level (`indentLevel`, which writes `a:p/@lvl`).
	 * PowerPoint: Paragraph > Indentation > Before text.
	 * - `'inherit'` writes no `@marL` at all, so the paragraph takes whatever its list style
	 *   (`a:lstStyle` → placeholder → layout → master) sets — the third state that omission
	 *   cannot spell, since omitting it writes the bullet-derived default
	 * - omitting the option keeps that default: a drawn bullet writes its own margin (see
	 *   `bullet.indent`), `bullet: false` writes `marL="0"`, `bullet: 'inherit'` writes nothing
	 * - a value here overrides all of those, in every bullet state
	 * @example 36 // body text starts 36pt (0.5in) from the frame's text edge
	 * @example 'inherit' // keep the list style's margin, even on a bulleted paragraph
	 * @remarks Read from the **opening run** of a paragraph, like {@link bullet} and unlike
	 * {@link align}; a value on a continuation run is ignored.
	 */
	paraMarginLeft?: number | 'inherit'
	/**
	 * Add a soft line-break (shift+enter) before line text content
	 * @default false
	 */
	softBreakBefore?: boolean
	/**
	 * tab stops
	 * - PowerPoint: Paragraph > Tabs > Tab stop position
	 * @example [{ position:1 }, { position:3 }] // Set first tab stop to 1 inch, set second tab stop to 3 inches
	 */
	tabStops?: Array<{ position: number; alignment?: 'l' | 'r' | 'ctr' | 'dec' }>
	/**
	 * text direction
	 * `horz` = horizontal
	 * `vert` = rotate 90^
	 * `vert270` = rotate 270^
	 * `wordArtVert` = stacked
	 * @default 'horz'
	 */
	textDirection?: 'horz' | 'vert' | 'vert270' | 'wordArtVert'
	/**
	 * Transparency (percent)
	 * - MS-PPT > Format Shape > Text Options > Text Fill & Outline > Text Fill > Transparency
	 * - range: 0-100
	 * @default 0
	 */
	transparency?: number
	/**
	 * underline properties
	 * - PowerPoint: Font > Color & Underline > Underline Style/Underline Color
	 * - `style` is the full `ST_TextUnderlineType` enumeration (ECMA-376 §20.1.10.81)
	 * - `'none'` is the *explicit off*: it writes `a:rPr/@u="none"`, which overrides an
	 *   underline the run would otherwise inherit from its list style, placeholder, layout or
	 *   master. Omitting the option instead states nothing and lets that inheritance stand —
	 *   the two are different facts, not two spellings of one.
	 * @default (unset) inherit
	 */
	underline?: {
		style?:
			| 'dash'
			| 'dashHeavy'
			| 'dashLong'
			| 'dashLongHeavy'
			| 'dbl'
			| 'dotDash'
			| 'dotDashHeavy'
			| 'dotDotDash'
			| 'dotDotDashHeavy'
			| 'dotted'
			| 'dottedHeavy'
			| 'heavy'
			| 'none'
			| 'sng'
			| 'wavy'
			| 'wavyDbl'
			| 'wavyHeavy'
			| 'words'
		color?: Color
	}
	/**
	 * vertical alignment
	 * @default 'top'
	 */
	valign?: VAlign
}

export interface TextGlowProps {
	/**
	 * Border color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * opacity (0.0 - 1.0)
	 * @example 0.5
	 * 50% opaque
	 */
	opacity?: number
	/**
	 * size (points)
	 */
	size: number
}

export interface TextFitShrinkProps {
	/**
	 * Shrink text on overflow (`<a:normAutofit>`)
	 */
	type: 'shrink'
	/**
	 * Font scale as a percent (0-100), mapped to `<a:normAutofit fontScale="..">`.
	 *
	 * PowerPoint normally calculates this dynamically when text overflows; set it
	 * explicitly to bake the scale into the generated file.
	 * @example 85 // render text at 85% of its nominal size
	 * @default undefined // attribute omitted (PowerPoint defaults to 100%)
	 */
	fontScale?: number
	/**
	 * Line-space reduction as a percent (0-100), mapped to `<a:normAutofit lnSpcReduction="..">`.
	 * @example 20 // reduce line spacing by 20%
	 * @default undefined // attribute omitted (PowerPoint defaults to 0%)
	 */
	lnSpcReduction?: number
}

export interface TextPropsOptions extends PositionProps, DataOrPathProps, TextBaseProps, ObjectNameProps {
	/**
	 * Preset-geometry start/end angles for text frames (`shape` + these build `<a:prstGeom>`
	 * adjustment guides through the shared emitter — `genXmlPresetGeom` reads them whether
	 * the object is a shape or a styled text frame).
	 */
	angleRange?: [number, number]
	/** Preset-geometry block-arc thickness for text frames. */
	arcThicknessRatio?: number
	/** Custom geometry points when a text frame uses a custom geometry shape. */
	points?: GeometryPoint[]
	/** Preset-geometry adjustment guides for text frames. */
	shapeAdjust?: ShapeAdjustValue | ShapeAdjustValue[]

	baseline?: number
	/**
	 * Character spacing
	 */
	charSpacing?: number
	/**
	 * Number of text columns in the text body
	 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Columns > "Number"
	 * - range: 1-16
	 * @default 1
	 * @example 2 // flow text into two columns
	 */
	columns?: number
	/**
	 * Spacing between text columns (points)
	 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Columns > "Spacing"
	 * - only applies when `columns` > 1
	 * @default 0
	 * @example 10 // 10pt gap between columns
	 */
	columnSpacing?: number
	/**
	 * Text fit options
	 *
	 * MS-PPT > Format Shape > Shape Options > Text Box > "[unlabeled group]": [3 options below]
	 * - 'none' = Do not Autofit
	 * - 'shrink' = Shrink text on overflow
	 * - 'resize' = Resize shape to fit text
	 *
	 * **Measured fit:** if you register the box's font with
	 * {@link TsPptx.registerFontMetrics}, both `'shrink'` and `'resize'` are
	 * **measured at export time**, so the text renders correctly in headless renderers
	 * and on plain file-open (no edit/resize needed):
	 * - `'shrink'` computes the largest `fontScale` at which the wrapped text fits and
	 *   bakes `<a:normAutofit fontScale=…/>`.
	 * - `'resize'` computes the height the text needs and bakes it into the shape's
	 *   `a:ext/@cy` (adjusting `a:off/@y` per vertical anchor), the marker being
	 *   `<a:spAutoFit/>`.
	 * Without registered metrics they fall back to the bare flag (`<a:normAutofit/>` /
	 * `<a:spAutoFit/>`, which only PowerPoint recomputes on edit) and warn once.
	 *
	 * **Note** Bare `'shrink'`/`'resize'` (no metrics) only take effect after editing
	 * text / resizing the shape; PowerPoint calculates the result then. The object form
	 * of `'shrink'` always bakes the explicit values you pass.
	 * @example 'shrink' // measured when metrics are registered; else bare <a:normAutofit/>
	 * @example 'resize' // measured when metrics are registered; else bare <a:spAutoFit/>
	 * @example { type: 'shrink', fontScale: 85, lnSpcReduction: 20 } // pre-shrink with explicit values
	 * @default "none"
	 */
	fit?: 'none' | 'shrink' | 'resize' | TextFitShrinkProps
	/**
	 * Shape fill, or a bare {@link Color} as shorthand for a solid fill ({@link FillOption}).
	 * @example 'FF0000' // hex color (red), shorthand for { color:'FF0000' }
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:SchemeColor.accent1 } // theme color Accent1
	 */
	fill?: FillOption
	/**
	 * Flip shape horizontally?
	 * @default false
	 */
	flipH?: boolean
	/**
	 * Flip shape vertical?
	 * @default false
	 */
	flipV?: boolean
	glow?: TextGlowProps
	hyperlink?: HyperlinkProps
	/**
	 * Outline level of the paragraph (`a:p/@lvl`) — which of the list style's nine levels the
	 * paragraph takes its bullet, indent and typography from.
	 * - range: 0-8, whole numbers only (`ST_TextIndentLevelType`); anything else is reported
	 *   under `text/invalid-indent-level` and ignored
	 * - `0` is the default and writes no attribute
	 * - this is the discrete level, not a measurement: `paraMarginLeft` and `paraIndent` are the
	 *   points-valued controls
	 * @default 0
	 */
	indentLevel?: number
	isTextBox?: boolean
	line?: ShapeLineProps
	/**
	 * Line spacing (pt)
	 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Exactly"
	 * @example 28 // 28pt
	 */
	lineSpacing?: number
	/**
	 * line spacing multiple (percent)
	 * - range: 0.0-9.99
	 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Multiple"
	 * @example 1.5 // 1.5X line spacing
	 */
	lineSpacingMultiple?: number
	/**
	 * Margin (inches) — the text-frame internal margin/padding
	 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Left/Right/Top/Bottom margin (shown in inches)
	 * - array order is `[top, right, bottom, left]`
	 * - a value `>= 1` is honored as inches but warns once (it is likely a legacy points value; divide by 72)
	 * @default (unset) PowerPoint's "Normal" internal margin [0.05", 0.1", 0.05", 0.1"]
	 * @example 0 // Top/Right/Bottom/Left margin 0
	 * @example 0.1 // Top/Right/Bottom/Left margin 0.1 inch
	 * @example [0.05, 0.1, 0.05, 0.1] // top 0.05", right 0.1", bottom 0.05", left 0.1"
	 */
	margin?: Margin
	outline?: { color: Color; size: number }
	paraSpaceAfter?: number
	paraSpaceBefore?: number
	/**
	 * Placeholder type
	 * - when the value matches a placeholder defined on the slide layout/master, this text
	 *   inherits that placeholder's position and formatting
	 * - otherwise the text shape is promoted to a standalone placeholder of this type, emitting
	 *   a real `<p:ph type="...">`. Use `placeholder: 'title'` to give a slide an accessible
	 *   title (PowerPoint's accessibility checker otherwise reports "Missing Slide Title")
	 * - values: 'title' | 'body' | et. al.
	 *
	 * **The placeholder supplies, it does not impose.** Every option this bag states wins; the
	 * layout placeholder fills in the ones it leaves out. So `{ placeholder: 'body' }` alone
	 * takes the placeholder's frame, anchor, margins, bullet and text style, and
	 * `{ placeholder: 'body', valign: 'top' }` takes all of that except the anchor. It matches
	 * PowerPoint, where a slide placeholder's XML carries only the properties it overrides and
	 * inherits the rest.
	 *
	 * The rule is about what the *caller* wrote, not about what the option bag holds by the time
	 * inheritance runs: a value this library defaults on the way past does not count as the
	 * caller stating one, and does not beat the layout.
	 * @example 'title'
	 * @see https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
	 */
	placeholder?: string
	/**
	 * Rounded rectangle radius (only for ShapeType.roundRect)
	 * - values: 0.0 to 1.0
	 * @default 0
	 */
	rectRadius?: number
	/**
	 * Rotation (degrees)
	 * - range: -360 to 360
	 * @default 0
	 * @example 180 // rotate 180 degrees
	 */
	rotate?: number
	/**
	 * Whether to enable right-to-left mode
	 * @default false
	 */
	rtlMode?: boolean
	/**
	 * Shadow options. **Which shadow depends on which bag it is on**, and the two are different
	 * effects rather than two spellings of one:
	 *
	 * - on the options passed to `addText` — the *shape's* bag — it is the shape's drop shadow
	 *   (`p:spPr/a:effectLst`), PowerPoint's Shape Effects ▸ Shadow;
	 * - on a run's own `options` inside the text array, it is the *text* shadow
	 *   (`a:rPr/a:effectLst`), PowerPoint's Text Effects ▸ Shadow.
	 *
	 * A run therefore does **not** inherit the shape's shadow: PowerPoint's two gestures are
	 * independent, and applying both is what darkens a shadow twice. State it in both places to
	 * get both, exactly as it takes two actions there.
	 * @example addText('hi', { shadow }) // the box has a shadow; its glyphs do not
	 * @example addText([{ text: 'hi', options: { shadow } }]) // the glyphs do; the box does not
	 */
	shadow?: ShadowProps
	shape?: SHAPE_NAME
	/**
	 * Strikethrough (`a:rPr/@strike`, `ST_TextStrikeType` — ECMA-376 §20.1.10.78)
	 * - `true` is `'sngStrike'`; `'dblStrike'` is the double rule
	 * - `'noStrike'` is the *explicit off*: it writes `strike="noStrike"`, which overrides a
	 *   strikethrough the run would otherwise inherit from its list style, placeholder, layout
	 *   or master
	 * - `false`, like leaving the option out, writes no attribute at all and so states
	 *   *nothing* — the run keeps whatever it inherits. This matches `bold`/`italic`, whose
	 *   falsy arm is likewise an omission; reach for `'noStrike'` when the intent is "not
	 *   struck" rather than "unspecified"
	 * @default (unset) inherit
	 */
	strike?: boolean | 'noStrike' | 'sngStrike' | 'dblStrike'
	subscript?: boolean
	superscript?: boolean
	/**
	 * Vertical alignment
	 * @default middle
	 */
	valign?: VAlign
	/**
	 * Advanced/legacy escape hatch for the full `ST_TextVerticalType` range (e.g. `eaVert`,
	 * `mongolianVert`, `wordArtVert`). Prefer {@link TextBaseProps.textDirection} for the common
	 * cases; both map to `a:bodyPr@vert`.
	 */
	vert?: TextVertType
	/**
	 * Text wrap
	 * @default true
	 */
	wrap?: boolean
}
export interface TextProps {
	text?: string | number
	options?: TextPropsOptions
	/**
	 * Raw OMML (Office MathML) for a native, editable PowerPoint equation. By default it is emitted
	 * as its own centered display-math paragraph (`<a14:m><m:oMathPara><m:oMath>…`) and `text` is
	 * ignored; set `inline` to instead flow it as an equation run mid-paragraph. Accepts either the
	 * inner OMML (children of `<m:oMath>`), a full `<m:oMath>…</m:oMath>`, or a full
	 * `<m:oMathPara>…</m:oMathPara>`; the `m:` prefix is resolved by the wrapper, so the markup does
	 * not need its own namespace declarations.
	 * This is the raw-OMML entry point. To author from LaTeX or MathML, convert with
	 * `latexToOmml()` / `mathmlToOmml()` from the `pptx-ts/math` subpath.
	 * @example { math: '<m:r><m:t>x^2+1=y</m:t></m:r>' } // raw OMML
	 * @example import { latexToOmml } from 'pptx-ts/math'; ({ math: latexToOmml('x^2+1=y') })
	 */
	math?: string
	/**
	 * Emit this item's `math` equation *inline* — an `<a14:m><m:oMath>` run flowing mid-paragraph
	 * between the surrounding plain text runs — instead of as its own centered display-math
	 * paragraph. Has no effect unless `math` is set. Pair with the bare-`<m:oMath>` inline form of
	 * the equation: `latexToOmml(tex, { display: false })` or `mathmlToOmml(mathml)`.
	 * @example addText([{ text: 'where ' }, { math: latexToOmml('x^2+1=y', { display: false }), inline: true }, { text: ' holds' }])
	 */
	inline?: boolean
}

/**
 * Options for layout-time text measurement ({@link TsPptx.measureText}).
 * Inches for width, points for type/spacing — the consumer-facing units. The
 * measured face must have metrics registered via {@link TsPptx.registerFontMetrics}
 * (a named face without exact metrics uses a conservative heuristic; an unnamed
 * theme-default face is unmeasurable).
 */
export interface MeasureTextOptions {
	/** Available text width in inches (the box width minus L/R inset, unless `insetIn` is given). */
	wIn: number
	/** Font size in points. */
	fontSize: number
	/** Font family name, as used in `fontFace`. Required for an exact measure; an unnamed face is unmeasurable. */
	fontFace?: string
	bold?: boolean
	italic?: boolean
	/** Character spacing in points. */
	charSpacing?: number
	/** Exact line spacing in points (overrides `lineSpacingMultiple`). */
	lineSpacing?: number
	/** Line spacing as a multiple of single (e.g. `1.5`). */
	lineSpacingMultiple?: number
	/** Space before each paragraph, in points. */
	paraSpaceBefore?: number
	/** Space after each paragraph, in points. */
	paraSpaceAfter?: number
	/** L/R text inset in inches; when set, subtracted from `wIn` on both sides (pass a raw box width). */
	insetIn?: number
}

/**
 * Result of {@link TsPptx.measureText}. Heights err **tall** (conservative) —
 * they match the value the export-time autofit bake uses, so the laid-out height is
 * ≥ what PowerPoint/LibreOffice render. Use it to grow a container; for an overflow
 * check it may slightly over-report (good for a warning, not a hard gate).
 */
export interface TextMeasurement {
	/** Laid-out height in inches at the given `fontSize` (conservative/tall). */
	heightIn: number
	/** Number of wrapped lines (conservative — the model wraps marginally early). */
	lineCount: number
	/**
	 * Width in inches of the widest laid-out line (conservative — the model wraps
	 * marginally early, so this errs slightly wide). With an unconstrained `wIn` it
	 * is the natural single-line width; constrained, it is the widest wrapped line.
	 * A box set to this width will not re-wrap the text.
	 */
	widestLineIn: number
	/** `false` only for an unnamed theme-default face that could not be measured. */
	measurable: boolean
	/**
	 * The **named** faces laid out with the conservative average-advance heuristic
	 * instead of their real metrics — i.e. faces with no registered metrics. Empty
	 * when every run was measured exactly (and for an unmeasurable result, which
	 * guessed nothing). The numbers stay conservative (they err tall/wide), but a
	 * non-empty list means they are an approximation, not a measurement; register
	 * the face via {@link TsPptx.registerFontMetrics} for an exact result.
	 */
	approximatedFaces: string[]
	/**
	 * Code points in the measured text — sorted, deduplicated — whose run resolved to a
	 * **registered** face that has no glyph for them. Empty for fully covered text.
	 *
	 * Unlike everything else here, a non-empty list means the numbers are **not reliably
	 * conservative**. PowerPoint renders such a code point from a substituted face and
	 * measures it in that face's advances; this model has no fallback and charges the
	 * registered font's `.notdef` advance, which is unrelated to the glyph that paints. A
	 * `.notdef` wider than the real glyph over-reports (harmless); a narrower one
	 * under-reports the width and can drop a line, so `heightIn` may come back **short**
	 * and the text overflow. Register a face that covers them (or set an explicit
	 * `fontFace` on those runs) before trusting the height.
	 *
	 * Faces with no registered metrics are not audited — they measure through the
	 * cmap-less heuristic and are reported in {@link approximatedFaces} instead.
	 * `String.fromCodePoint(...uncoveredCodepoints)` renders the list.
	 */
	uncoveredCodepoints: number[]
	/** True if the text fits a box of inner height `hIn` (inches) at full size. */
	fitsBox: (hIn: number) => boolean
	/** The `fontScale` (percent) that fits inner height `hIn`; `100` if it already fits, never below the shrink floor. */
	shrinkScaleFor: (hIn: number) => number
}

/** Options for {@link TsPptx.overflowsBox}: a measure plus the box inner height to test against. */
export interface OverflowBoxOptions extends MeasureTextOptions {
	/** Box inner height in inches to test for overflow. */
	hIn: number
}

/**
 * One cell's computed rectangle from {@link TsPptx.tableLayout}. All values are
 * inches; `x`/`y` are absolute (offset from the table's `x`/`y`). For a merged cell,
 * `row`/`col` are the top-left origin and `wIn`/`hIn` cover the whole span; the
 * cells it covers are not emitted separately.
 */
export interface TableCellLayout {
	/** Zero-based grid row of the cell's top-left origin. */
	row: number
	/** Zero-based grid column of the cell's top-left origin. */
	col: number
	/** Number of rows the cell spans (1 if not merged). */
	rowSpan: number
	/** Number of columns the cell spans (1 if not merged). */
	colSpan: number
	/** Left edge in inches (absolute). */
	xIn: number
	/** Top edge in inches (absolute). */
	yIn: number
	/** Outer cell width in inches (sum of spanned column widths). */
	wIn: number
	/** Outer cell height in inches (sum of spanned row heights). */
	hIn: number
	/**
	 * `true` when `hIn`/`yIn` are pinned by an explicit `rowH` (array or scalar) or
	 * table `h`; `false` when the row is auto-height and the value is a conservative
	 * (tall) estimate from the same text model as {@link TsPptx.measureText}.
	 */
	heightExact: boolean
}

/**
 * Result of {@link TsPptx.tableLayout}: per-cell geometry plus overall table
 * bounds, for placing images/shapes over a table without rendering it. Geometry is
 * for a single, un-paginated table laid out at `opts.x`/`y`/`w`; `autoPage` paging
 * is not modeled. Widths are exact; auto-height row heights are conservative
 * estimates (see {@link TableCellLayout.heightExact}).
 */
export interface TableLayoutResult {
	/** One entry per non-merged origin cell, in row-major order. */
	cells: TableCellLayout[]
	/** Overall table width in inches (sum of column widths). */
	widthIn: number
	/** Overall table height in inches (sum of row heights; may include estimates). */
	heightIn: number
	/** `false` if any row height was estimated (the total errs tall, like `measureText`). */
	heightExact: boolean
}

/**
 * Per-run options for a speaker-notes text run.
 * A focused subset of `TextPropsOptions`: inline formatting plus an (external URL) hyperlink.
 * Notes hyperlinks support `url` only; `slide` targets are not yet supported.
 */
export type NotesTextOptions = Pick<
	TextPropsOptions,
	'hyperlink' | 'bold' | 'italic' | 'underline' | 'color' | 'fontSize' | 'fontFace'
>

/** A single speaker-notes text run: text plus optional inline formatting / hyperlink. */
export interface NotesProps {
	text: string
	options?: NotesTextOptions
}

/**
 * A review comment attached to a slide (legacy ISO/IEC 29500 §13 comment).
 */
export interface CommentProps {
	/** Author display name (required). Comments sharing the same `author`+`initials` are grouped under one author entry. */
	author: string
	/** Author initials shown in the comment marker. Defaults to letters derived from `author`. */
	initials?: string
	/** Comment body text (required). */
	text: string
	/** Comment marker X position in inches. @default 0.5 */
	x?: number
	/** Comment marker Y position in inches. @default 0.5 */
	y?: number
	/** Authored date/time as a `Date` or ISO-8601 string. Omitted from the XML when not provided. */
	date?: Date | string
}

/** Internal normalized comment stored on a slide (`_comments`). x/y are inches; `date` is ISO-8601 when present. */
export interface SlideComment {
	author: string
	initials: string
	text: string
	x: number
	y: number
	date?: string
}

/** Resolved presentation-level comment author, emitted to `commentAuthors.xml`. */
export interface ResolvedCommentAuthor {
	id: number
	name: string
	initials: string
	lastIdx: number
	clrIdx: number
}

/** Factory for a single inline text run. Prevents `as never` casts when building mixed-style run arrays. */
export function textRun(text: string | number, options?: TextPropsOptions): TextProps {
	return options !== undefined ? { text, options } : { text }
}

/** Wraps a run array so TypeScript accepts it as `TextProps[]` without a cast. */
export function textRuns(runs: TextProps[]): TextProps[] {
	return runs
}
