/**
 * Table types: `TableProps`/`TableCell(+Props)`, table styles, `tableToSlides` options and layout results.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { SHAPE_NAME, SlideObjectType, TableStyle } from '../enums.js'
import type { DataOrPathProps, HexColor, Margin, PositionProps } from './core.js'
import type { ObjectNameProps } from './object.js'
import type { ShapeProps } from './shape.js'
import type { BorderProps, HyperlinkProps, ShapeFillProps } from './style.js'
import type { TextBaseProps, TextFitShrinkProps, TextProps, TextPropsOptions } from './text.js'

/**
 * A `<table>` element, from whatever DOM implementation the caller has.
 *
 * Deliberately *not* `lib.dom`'s `Element`. `lib.dom`'s types are built for a browser, and a
 * non-browser DOM's element type — happy-dom's, jsdom's — declares its own members and does
 * not satisfy them, so demanding `Element` would reject from TypeScript exactly the
 * implementations `ts-pptx/html` exists to accept. Requiring the three members the conversion
 * actually calls keeps a mistyped argument an error while letting every real DOM through.
 */
export interface TableToSlidesElement {
	getAttribute: (name: string) => string | null
	querySelector: (selectors: string) => unknown
	querySelectorAll: (selectors: string) => unknown
}

/**
 * The one thing `tableToSlides` asks a document for. Structural for the same reason
 * {@link TableToSlidesElement} is.
 */
export interface TableToSlidesDocument {
	getElementById: (elementId: string) => TableToSlidesElement | null
}

export interface TableToSlidesProps extends TableProps {
	_arrObjTabHeadRows?: TableRow[]
	// _masterSlide?: SlideLayout

	/**
	 * Add an image to slide(s) created during autopaging
	 * - `image` prop requires either `path` or `data`
	 * - see `DataOrPathProps` for details on `image` props
	 * - see `PositionProps` for details on `options` props
	 */
	addImage?: { image: DataOrPathProps; options: PositionProps }
	/**
	 * Add a shape to slide(s) created during autopaging
	 */
	addShape?: { shapeName: SHAPE_NAME; options: ShapeProps }
	/**
	 * Add a table to slide(s) created during autopaging
	 */
	addTable?: { rows: TableRow[]; options: TableProps }
	/**
	 * Add a text object to slide(s) created during autopaging
	 */
	addText?: { text: TextProps[]; options: TextPropsOptions }
	/**
	 * Whether to enable auto-paging
	 * - auto-paging creates new slides as content overflows a slide
	 * @default true
	 */
	autoPage?: boolean
	/**
	 * Auto-paging character weight
	 * - adjusts how many characters are used before lines wrap
	 * - range: -1.0 to 1.0
	 * @default 0.0
	 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
	 */
	autoPageCharWeight?: number
	/**
	 * Auto-paging line weight
	 * - adjusts how many lines are used before slides wrap
	 * - range: -1.0 to 1.0
	 * @default 0.0
	 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
	 */
	autoPageLineWeight?: number
	/**
	 * Whether to repeat head row(s) on new tables created by autopaging
	 * @default false
	 */
	autoPageRepeatHeader?: boolean
	/**
	 * The `y` location to use on subsequent slides created by autopaging
	 * @default (top margin of Slide)
	 */
	autoPageSlideStartY?: number
	/**
	 * Column widths (inches)
	 */
	colW?: number | number[]
	/**
	 * Document to resolve a table *id* against.
	 * - only consulted when the table is identified by string id; passing the element itself
	 *   needs no document, since one is reached through the element's `ownerDocument`
	 * - defaults to the global `document`, which is what a browser caller gets
	 * - supply this to convert a table outside a browser (any DOM implementation will do)
	 */
	document?: TableToSlidesDocument
	/**
	 * Title of the slide master to use for the auto-paged slides (the `title` passed to
	 * `defineSlideMaster`). Matches the `masterTitle` option of `addSlide`.
	 * - define a master slide to have your auto-paged slides carry a corporate design, etc.
	 */
	masterTitle?: string
	/**
	 * Slide margin
	 * - this margin will be across all slides created by auto-paging
	 */
	slideMargin?: Margin
}
/**
 * A table cell's corner-to-corner rules (`a:tcPr/a:lnTlToBr` / `a:lnBlToTr`) — the
 * "Diagonal Down Border" / "Diagonal Up Border" of PowerPoint's border gallery.
 *
 * Deliberately its own object rather than two more entries on {@link TableCellProps.border}'s
 * tuple: widening that tuple to six would break every existing caller for a construct almost
 * no deck uses, and the two diagonals are not edges — they do not participate in the
 * perimeter, and they are not inherited by a merged region's covered cells.
 *
 * @example diagonal: { tlToBr: { type:'solid', color:'C00000', width:1 } } // strike a cell out
 * @example diagonal: { tlToBr: { type:'solid' }, blToTr: { type:'solid' } } // an X across the cell
 */
export interface TableCellDiagonals {
	/** Top-left to bottom-right (`╲`, `a:lnTlToBr`). */
	tlToBr?: BorderProps
	/** Bottom-left to top-right (`╱`, `a:lnBlToTr`). */
	blToTr?: BorderProps
}
/** `ST_BevelPresetType` — the bevel profile applied to a cell's edges. */
export type BevelPresetType =
	| 'relaxedInset'
	| 'circle'
	| 'slope'
	| 'cross'
	| 'angle'
	| 'softRound'
	| 'convex'
	| 'coolSlant'
	| 'divot'
	| 'riblet'
	| 'hardEdge'
	| 'artDeco'
/** `ST_PresetMaterialType` — how a 3-D surface responds to light. */
export type PresetMaterialType =
	| 'legacyMatte'
	| 'legacyPlastic'
	| 'legacyMetal'
	| 'legacyWireframe'
	| 'matte'
	| 'plastic'
	| 'metal'
	| 'warmMatte'
	| 'translucentPowder'
	| 'powder'
	| 'dkEdge'
	| 'softEdge'
	| 'clear'
	| 'flat'
	| 'softmetal'
/** `ST_LightRigType` — the preset lighting scene. */
export type LightRigType =
	| 'legacyFlat1'
	| 'legacyFlat2'
	| 'legacyFlat3'
	| 'legacyFlat4'
	| 'legacyNormal1'
	| 'legacyNormal2'
	| 'legacyNormal3'
	| 'legacyNormal4'
	| 'legacyHarsh1'
	| 'legacyHarsh2'
	| 'legacyHarsh3'
	| 'legacyHarsh4'
	| 'threePt'
	| 'balanced'
	| 'soft'
	| 'harsh'
	| 'flood'
	| 'contrasting'
	| 'morning'
	| 'sunrise'
	| 'sunset'
	| 'chilly'
	| 'freezing'
	| 'flat'
	| 'twoPt'
	| 'glow'
	| 'brightRoom'
/** `ST_LightRigDirection` — where the light rig sits relative to the scene. */
export type LightRigDirection = 'tl' | 't' | 'tr' | 'l' | 'r' | 'bl' | 'b' | 'br'
/**
 * A cell's 3-D bevel (`a:tcPr/a:cell3D`).
 *
 * A minimal surface over `CT_Cell3D`, which is a required `a:bevel` plus an optional
 * `a:lightRig`. PowerPoint's table UI exposes no control for it — it reaches a deck through a
 * theme's table style or through another producer — so this exists to author and reproduce
 * one, not as a general 3-D model.
 */
export interface TableCell3DProps {
	/**
	 * The bevel's preset shape (`a:bevel/@prst`).
	 * @default 'circle'
	 */
	preset?: BevelPresetType
	/**
	 * Bevel width in points (`a:bevel/@w`; the attribute is EMU).
	 * @default 6 // 76200 EMU
	 */
	width?: number
	/**
	 * Bevel height in points (`a:bevel/@h`; the attribute is EMU).
	 * @default 6 // 76200 EMU
	 */
	height?: number
	/**
	 * Surface material (`a:cell3D/@prstMaterial`), which decides how the bevel takes light.
	 * @default 'plastic'
	 */
	material?: PresetMaterialType
	/**
	 * Scene lighting (`a:cell3D/a:lightRig`). Both fields are required by the schema, so
	 * either give both or omit the whole object and let the renderer light the bevel.
	 * @example lightRig: { rig: 'threePt', dir: 't' }
	 */
	lightRig?: { rig: LightRigType; dir: LightRigDirection }
}
export interface TableCellProps extends TextBaseProps {
	/**
	 * Auto-paging character weight
	 * - adjusts how many characters are used before lines wrap
	 * - range: -1.0 to 1.0
	 * @default 0.0
	 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
	 */
	autoPageCharWeight?: number
	/**
	 * Auto-paging line weight
	 * - adjusts how many lines are used before slides wrap
	 * - range: -1.0 to 1.0
	 * @default 0.0
	 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
	 */
	autoPageLineWeight?: number
	/**
	 * The cell's four edge borders (`a:tcPr/a:lnT|lnR|lnB|lnL`).
	 * - a single `BorderProps` is broadcast to all four sides
	 * - an array is read in **TRBL** order (`[top, right, bottom, left]`)
	 *
	 * Overrides the table-level {@link TableProps.border} default entirely — the two do not
	 * merge per side. For the two corner-to-corner rules see {@link diagonal}; for the
	 * table's outside edge see {@link TableProps.outerBorder}.
	 */
	border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
	/**
	 * The cell's corner-to-corner rules — PowerPoint's "Diagonal Down/Up Border".
	 *
	 * Kept off {@link border}'s tuple on purpose: see {@link TableCellDiagonals}. A diagonal
	 * on a merged cell is one stroke across the whole region, so only the span origin carries
	 * it — the covered cells inherit the origin's *edges* but never its diagonals.
	 * @example diagonal: { tlToBr: { type:'solid', color:'C00000' } } // strike the cell out
	 */
	diagonal?: TableCellDiagonals
	/**
	 * Centre the cell's whole text **block** horizontally within the cell
	 * (`a:tcPr/@anchorCtr`), independent of each paragraph's own `align`.
	 *
	 * The two do different things and compose: `align` decides where each line sits inside
	 * the text block, `anchorCtr` decides where that block sits inside the cell. With
	 * left-aligned text in a wide cell, `anchorCtr: true` centres the ragged-right column of
	 * lines as a unit while keeping their left edges flush with each other.
	 *
	 * `false` is the schema default, so it emits nothing — the same reasoning as
	 * `horzOverflow: 'clip'`.
	 * @default false
	 */
	anchorCtr?: boolean
	/**
	 * A 3-D bevel on the cell (`a:tcPr/a:cell3D`). Niche: PowerPoint's table UI has no
	 * control for it, so it reaches a deck from a theme or another producer, and this exists
	 * to author and reproduce one. See {@link TableCell3DProps}.
	 */
	cell3D?: TableCell3DProps
	/**
	 * Cell colspan
	 */
	colspan?: number
	/**
	 * Cell fill — a solid color, or a picture that fills the cell (`a:blipFill` in
	 * the cell's `a:tcPr`, stretched to the cell box). A picture fill embeds the
	 * image as slide media; identical sources are embedded once. Raster only —
	 * an SVG source warns and is ignored, matching shape fills.
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:SchemeColor.accent1 } // theme color Accent1
	 * @example { type:'image', image:{ path:'logo.png' } } // picture fill
	 * @example { image:{ data:'image/png;base64,…' } } // picture fill, `type` inferred
	 */
	fill?: ShapeFillProps
	/**
	 * What happens to a **single glyph** wider than the cell's text width (`a:tcPr/@horzOverflow`).
	 * - `'clip'` (PowerPoint's default) cuts the glyph off at the cell edge
	 * - `'overflow'` lets it draw past the edge, over whatever is beside it
	 *
	 * This is **not** a text-wrap switch, despite the name. Lines always wrap to the column
	 * width; only an individual character too wide to fit on a line by itself is affected, so
	 * it matters for oversized display type, wide CJK/emoji glyphs, and icon fonts in a narrow
	 * column. PowerPoint has no per-cell no-wrap at all — `wrap="none"` on a cell's `a:bodyPr`
	 * is discarded on the next save (probe:
	 * `test/read/fixtures/authoring/probe-table-cell-wrap.ps1`).
	 * `'clip'` is written as asked but is the schema default, so PowerPoint drops the
	 * attribute again the first time it saves the deck — the visible result is the same
	 * either way, and leaving this unset emits nothing at all.
	 * @default (unset — PowerPoint clips)
	 * @example 'overflow' // a 60pt glyph in a 0.5in column draws whole instead of being cut
	 */
	horzOverflow?: 'clip' | 'overflow'
	hyperlink?: HyperlinkProps
	/**
	 * Cell margin (inches)
	 * @default 0
	 */
	margin?: Margin
	/**
	 * Cell rowspan
	 */
	rowspan?: number
	/**
	 * Shrink cell text to fit when it would overflow the cell's fixed height.
	 * - `'shrink'` measures the wrapped text and bakes a **reduced literal font size**
	 *   onto the cell's runs so the text fits — PowerPoint does not support text
	 *   autofit (`normAutofit`) inside table cells, so there is no font-scale flag to
	 *   set; the size itself is lowered, which both PowerPoint and LibreOffice render
	 *   identically with no edit/resize.
	 * - Requires the cell font registered via {@link TsPptx.registerFontMetrics};
	 *   without metrics it is a no-op (the cell keeps its authored size) and warns once.
	 * - Only triggers when the cell's row has a **fixed** height that the text exceeds.
	 *   With auto-height rows (no `rowH`/`h`), the row simply grows, so nothing shrinks.
	 * - Only `'shrink'` is acted on for cells. `'resize'` and the object form are ignored
	 *   here: a table row already auto-grows to fit its tallest cell (the cell equivalent
	 *   of `spAutoFit`), so there is nothing to bake. (The wider union is shared with
	 *   {@link TextPropsOptions.fit} so table-level `fit` can cascade to cells.)
	 * @example 'shrink' // measured when the cell font is registered; else no-op
	 */
	fit?: 'none' | 'shrink' | 'resize' | TextFitShrinkProps
}
/**
 * Styling for one region of a custom table style (maps to a `CT_TablePartStyle`).
 * A region (e.g. the header row or banded rows) is shown only when the matching
 * `TableProps` flag is set — `firstRow` needs `hasHeader`, `band1H`/`band2H` need
 * `hasBandedRows`, and so on.
 * @see TableStyleProps
 */
export interface TableStyleRegionProps {
	/**
	 * Solid cell fill color (hex).
	 * - `HexColor` only; theme references are not supported in custom styles
	 * @example '1A2B3C'
	 */
	fill?: HexColor
	/**
	 * Text color (hex).
	 * @example 'FFFFFF'
	 */
	color?: HexColor
	/** Bold text. */
	bold?: boolean
	/** Italic text. */
	italic?: boolean
	/**
	 * Cell border(s).
	 * - single value is applied to all four sides plus the interior grid lines
	 * - array of values in TRBL order styles only the four outer sides
	 */
	border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
}
/**
 * A reusable custom table style written to `ppt/tableStyles.xml`.
 * Pass to `pptx.defineTableStyle()`, which registers it and returns a GUID to use
 * as `TableProps.tableStyle`. Unlike the fixed built-in `TableStyle` set, a custom
 * style can use arbitrary brand colors, is editable in PowerPoint's Table Styles
 * gallery, and bands correctly across any row/column count (including auto-paged tables).
 *
 * **A region's `fill` and `bold` apply; its `border` and `color` do not.** The library
 * stamps a default `border` (`{type:'none'}` on all four sides) and `color` (`'000000'`)
 * onto every cell as *direct* formatting, and direct formatting outranks a style region in
 * PowerPoint — so a style's borders and text colour are overridden before they can render.
 * Set those on the table instead (`border`) or on `headerRow` / the cells (`color`), as
 * below. See `docs/tables.md` → "The defaults tier".
 * @example
 * const brand = pptx.defineTableStyle({
 *   name: 'Brand Banded',
 *   firstRow: { fill:'1A2B3C', bold:true },   // fill + bold come from the style
 *   band1H:   { fill:'EAF1F8' },
 *   band2H:   { fill:'FFFFFF' },
 * })
 * slide.addTable(rows, {
 *   tableStyle: brand, hasHeader:true, hasBandedRows:true,
 *   border: { type:'solid', color:'D9D9D9', width:0.5 },  // not wholeTbl.border
 *   headerRow: { color:'FFFFFF' },                        // not firstRow.color
 * })
 */
export interface TableStyleProps {
	/** Display name shown in PowerPoint's Table Styles gallery. */
	name: string
	/** Base styling applied to every cell. */
	wholeTbl?: TableStyleRegionProps
	/** Header (first) row — activated by `TableProps.hasHeader`. */
	firstRow?: TableStyleRegionProps
	/** Footer (last) row — activated by `TableProps.hasFooter`. */
	lastRow?: TableStyleRegionProps
	/** First column — activated by `TableProps.hasFirstColumn`. */
	firstCol?: TableStyleRegionProps
	/** Last column — activated by `TableProps.hasLastColumn`. */
	lastCol?: TableStyleRegionProps
	/** Odd horizontal band — activated by `TableProps.hasBandedRows`. */
	band1H?: TableStyleRegionProps
	/** Even horizontal band — activated by `TableProps.hasBandedRows`. */
	band2H?: TableStyleRegionProps
	/** Odd vertical band — activated by `TableProps.hasBandedColumns`. */
	band1V?: TableStyleRegionProps
	/** Even vertical band — activated by `TableProps.hasBandedColumns`. */
	band2V?: TableStyleRegionProps
}
export interface TableProps extends PositionProps, TextBaseProps, ObjectNameProps {
	_arrObjTabHeadRows?: TableRow[]

	/**
	 * Name of a table/content placeholder defined on the slide layout/master to bind this table to.
	 * - when it matches a layout/master placeholder, the table's `<p:graphicFrame>` emits that
	 *   placeholder's `<p:ph>` (idx/type) so PowerPoint treats the table as filling the placeholder
	 *   (e.g. a "Title and Content" content placeholder)
	 * - the table also inherits the placeholder's position/size for any of x/y/w/h left unset
	 * @example 'body' // bind to the layout placeholder named 'body'
	 */
	placeholder?: string

	/**
	 * Whether to enable auto-paging
	 * - auto-paging creates new slides as content overflows a slide
	 * @default false
	 */
	autoPage?: boolean
	/**
	 * Auto-paging character weight
	 * - adjusts how many characters are used before lines wrap
	 * - range: -1.0 to 1.0
	 * @default 0.0
	 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
	 */
	autoPageCharWeight?: number
	/**
	 * Auto-paging line weight
	 * - adjusts how many lines are used before slides wrap
	 * - range: -1.0 to 1.0
	 * @default 0.0
	 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
	 */
	autoPageLineWeight?: number
	/**
	 * Whether table header row(s) should be repeated on each new slide creating by autoPage.
	 * Use `autoPageHeaderRows` to designate how many rows comprise the table header (1+).
	 * @default false
	 */
	autoPageRepeatHeader?: boolean
	/**
	 * Number of rows that comprise table headers
	 * - required when `autoPageRepeatHeader` is set to true.
	 * @example 2 - repeats the first two table rows on each new slide created
	 * @default 1
	 */
	autoPageHeaderRows?: number
	/**
	 * The `y` location to use on subsequent slides created by autopaging
	 * @default (top margin of Slide)
	 */
	autoPageSlideStartY?: number
	/**
	 * Whether populated placeholders on the source slide (e.g. a title set via
	 * `addText(text, { placeholder })`) are copied onto each overflow slide created by autoPage.
	 * - new slides otherwise inherit only the layout's empty placeholders, so a title set on the
	 *   first slide would not appear on continuation slides.
	 * @default false
	 */
	autoPagePlaceholder?: boolean
	/**
	 * Default border for **every cell** in the table — not the table's outer perimeter.
	 *
	 * A single `BorderProps` is broadcast to all four sides of each cell; an array is read
	 * in TRBL order and applied to each cell's own four sides. So
	 * `border: [{type:'solid'}, {type:'none'}, {type:'solid'}, {type:'none'}]` gives *every*
	 * cell a top and bottom rule — a full set of horizontal grid lines, not a rule at the
	 * top and bottom of the table. A cell's own `options.border` overrides this entirely.
	 *
	 * For the perimeter, use {@link outerBorder}, which applies only to the cells on the
	 * table's outside edge. The two compose: `border` draws the interior grid and
	 * `outerBorder` overrides the edges it reaches.
	 * @example border: { type:'solid', color:'D9D9D9', width:0.5 } // full grid
	 * @example border: [{type:'solid'}, {type:'none'}, {type:'solid'}, {type:'none'}] // horizontal rules only
	 */
	border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
	/**
	 * Border for the table's outer perimeter only — the top edge of the first row, the
	 * bottom edge of the last row, the left edge of the first column and the right edge of
	 * the last column.
	 *
	 * A single `BorderProps` is used for all four perimeter sides; an array is read in TRBL
	 * order, so `outerBorder: [box, box, box, box]` boxes the table and
	 * `outerBorder: [rule, undefined, rule, undefined]` rules only above and below it.
	 * An `undefined`/omitted entry leaves that side to whatever {@link border} (or the cell's
	 * own `options.border`) already put there.
	 *
	 * Applied **after** every other border source, so it wins on the sides it touches and
	 * leaves the interior alone — the "outline the table, no interior grid" case is
	 * `outerBorder` with no `border` at all. Merged cells are handled: a span's covered cells
	 * inherit the origin's borders, so a colspan straddling the last column still gets the
	 * right-hand rule.
	 * @example outerBorder: { type:'solid', color:'1A2B3C', width:1 } // box the table
	 * @example outerBorder: [{type:'solid'}, undefined, {type:'solid'}, undefined] // rules above and below
	 */
	outerBorder?: BorderProps | [BorderProps?, BorderProps?, BorderProps?, BorderProps?]
	/**
	 * Width of table columns (inches)
	 * - single value is applied to every column equally based upon `w`
	 * - array of values in applied to each column in order
	 * @default columns of equal width based upon `w`
	 */
	colW?: number | number[]
	/**
	 * Shrink columns proportionally so a too-wide table fits the slide.
	 * - `'shrink'`: if the total column width exceeds the space available between
	 *   the table's `x` and the right slide margin, scale every column down by the
	 *   same factor so the whole table fits. Columns that already fit are untouched.
	 * - applies to an explicit `colW` array (the common "too many columns" case) and
	 *   to a `w` wider than the slide; never grows columns and never enforces a
	 *   minimum width, so a very high column count can still become thin.
	 * - opt-in: explicit widths are otherwise emitted as-is and may run off the slide.
	 * @default undefined (no scaling)
	 */
	fitColumns?: 'shrink'
	/**
	 * Table-level text-fit policy, cascaded to every cell that does not set its own
	 * {@link TableCellProps.fit}.
	 * - `'shrink'`: each cell whose row has a **fixed** height that its wrapped text
	 *   would exceed gets a reduced literal font size baked onto its runs so the text
	 *   fits (see {@link TableCellProps.fit} for the full semantics and the font-metrics
	 *   requirement). Cells that set their own `fit` win over this table default.
	 * - `'none'` / `'resize'` / the object form are accepted for parity with
	 *   {@link TextPropsOptions.fit} but only `'shrink'` is acted on for table cells.
	 * @default undefined (cells keep their authored size)
	 * @example 'shrink'
	 */
	fit?: 'none' | 'shrink' | 'resize' | TextFitShrinkProps
	/**
	 * Mark the first row as a header row.
	 * Emits `firstRow="1"` on `<a:tblPr>`, activating the first-row style region of
	 * the table style and satisfying the PowerPoint accessibility checker's "table header" rule.
	 * @default false
	 */
	hasHeader?: boolean
	/**
	 * Mark the last row as a footer row.
	 * Emits `lastRow="1"` on `<a:tblPr>`, activating the last-row style region.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasFooter?: boolean
	/**
	 * Enable alternating row (band) shading.
	 * Emits `bandRow="1"` on `<a:tblPr>`, activating band1H/band2H style regions.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasBandedRows?: boolean
	/**
	 * Enable alternating column (band) shading.
	 * Emits `bandCol="1"` on `<a:tblPr>`, activating band1V/band2V style regions.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasBandedColumns?: boolean
	/**
	 * Apply special styling to the first column.
	 * Emits `firstCol="1"` on `<a:tblPr>`, activating the firstCol style region.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasFirstColumn?: boolean
	/**
	 * Apply special styling to the last column.
	 * Emits `lastCol="1"` on `<a:tblPr>`, activating the lastCol style region.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasLastColumn?: boolean
	/**
	 * Lay the table out right-to-left.
	 * Emits `rtl="1"` on `<a:tblPr>`, which mirrors the column order so the first
	 * column renders on the right — the correct layout for RTL scripts (Arabic, Hebrew).
	 * This controls only the table/column direction; per-cell text direction is set
	 * with each cell's `rtlMode` option.
	 * @default false
	 */
	rtl?: boolean
	/**
	 * Table style to apply, either a built-in `TableStyle` member or the GUID
	 * returned by `pptx.defineTableStyle()` for a custom style.
	 * Emits `<a:tableStyleId>` inside `<a:tblPr>` with the corresponding GUID.
	 * Style flags (`hasHeader`, `hasFooter`, `hasBandedRows`, etc.) select which
	 * regions of the chosen style are activated; they have no visible effect without
	 * a `tableStyle` set.
	 *
	 * @example tableStyle: pptx.TableStyle.MEDIUM_STYLE_2_ACCENT_1 // built-in
	 * @example const brand = pptx.defineTableStyle({ name:'Brand', firstRow:{ fill:'1A2B3C', color:'FFFFFF', bold:true } }); tableStyle: brand
	 */
	tableStyle?: TableStyle | string
	/**
	 * Inline styling for the header (first) row, applied as direct per-cell formatting.
	 *
	 * Convenience shortcut for styling a header distinctly from the body **without** first
	 * registering a custom style via `pptx.defineTableStyle({ firstRow })`. Each property is
	 * merged onto every cell of row 0 (`fill`, `color`, `bold`, `align`, `border`, etc.).
	 *
	 * Precedence (highest wins), matching how PowerPoint resolves styling — direct cell
	 * formatting overrides a table-style region:
	 * 1. explicit per-cell `options` on a row-0 cell
	 * 2. this `headerRow`
	 * 3. the `firstRow` region of any `tableStyle`
	 * 4. `wholeTbl` / defaults
	 *
	 * Setting `headerRow` also implies `hasHeader: true` (emits `firstRow="1"` for the
	 * accessibility "table header" marker) unless `hasHeader` is explicitly set to `false`.
	 * @example headerRow: { fill: { color:'1A2B3C' }, color:'FFFFFF', bold:true, align:'center' }
	 */
	headerRow?: TableCellProps
	/**
	 * Per-column default cell styling, applied as direct per-cell formatting.
	 *
	 * `columns[i]` is merged onto every cell that starts in column `i` (`fill`, `color`,
	 * `bold`, `align`, `valign`, `border`, `margin`, …), so a wide **colored** matrix —
	 * per-column fills, maturity-gradient headers — does not require hand-writing a fill
	 * onto every cell. Entries may be sparse (`columns[2]` alone styles only column 2);
	 * an `undefined`/omitted entry leaves that column untouched. The whole option is
	 * optional and degrades cleanly to today's text-on-white when not given.
	 *
	 * Precedence (highest wins), matching how PowerPoint resolves styling — direct cell
	 * formatting overrides a table-style region:
	 * 1. explicit per-cell `options`
	 * 2. `headerRow` (row 0 only)
	 * 3. this `columns[colIdx]`
	 * 4. `wholeTbl` / `tableStyle` / defaults
	 *
	 * The merge is property-level, so a header cell keeps `headerRow`'s typography **and**
	 * takes its column's `fill` when they set different properties. For a graduated header
	 * band, put shared header typography (bold/white/centered, **no fill**) in `headerRow`
	 * and let each `columns[i].fill` supply that column's fill.
	 *
	 * Column index counts each cell's `colspan` (default 1) within a row, so merged cells
	 * map to the correct column; it does not track `rowspan`s inherited from earlier rows.
	 *
	 * There is deliberately no built-in "group bracket" annotation. Label a span of columns
	 * by composing existing primitives: `addShape('rightBrace', …)` (or `'bracePair'`) plus
	 * `addText`, positioned from the table's `x` and `colW` (a span's `x`/`w` is the running
	 * sum of the preceding/covered column widths).
	 *
	 * @example columns: [{}, { fill: { color:'E8F0FE' } }, { fill: { color:'C6DAFC' } }] // per-column solid fills
	 * @example // gradient header: shared typography on headerRow, graduated fills per column
	 * headerRow: { color:'FFFFFF', bold:true, align:'center' },
	 * columns: [{}, { fill:{color:'BBD3FB'} }, { fill:{color:'89AEF6'} }, { fill:{color:'4B7BE5'} }]
	 */
	columns?: TableCellProps[]
	/**
	 * Default cell background — **stamped onto every cell**, not painted behind the table.
	 *
	 * Each cell that does not set its own `options.fill` (or take one from `headerRow` /
	 * `columns[i]`) gets this as its own `a:tcPr` fill. A picture fill is embedded once and
	 * shared by every cell that inherits it.
	 *
	 * Visually this is usually what you want, but it is not what a source deck contains, and
	 * it has one real consequence: because every cell ends up explicitly filled, a cell that
	 * later gets its own fill has nothing to fall back to. For a genuine table background —
	 * one `a:tblPr` fill that cells sit on top of — use {@link tableFill}.
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:SchemeColor.accent1 } // theme color Accent1
	 * @example { type:'image', image:{ path:'watermark.png' } } // picture fill
	 */
	fill?: ShapeFillProps
	/**
	 * The table's own background (`a:tblPr` fill) — one fill behind the whole table that
	 * every cell sits on top of.
	 *
	 * The counterpart to {@link fill}, and the difference is where the paint lands rather
	 * than how it looks: `fill` copies a colour onto every cell, while this writes a single
	 * fill on the table itself. That matters when a cell is meant to be transparent — with
	 * `tableFill`, an unfilled cell shows the table background through; with `fill` there is
	 * no such thing as an unfilled cell. It is also what a deck read back from PowerPoint
	 * actually carries, so a replica reproduces the source's structure rather than an
	 * equivalent-looking flattening of it.
	 *
	 * Takes the same `ShapeFillProps` a cell does — solid, gradient, pattern or picture.
	 * @example { color:'F2F2F2' } // a light background behind the whole table
	 * @example { type:'image', image:{ path:'watermark.png' } } // one watermark, not one per cell
	 */
	tableFill?: ShapeFillProps
	/**
	 * Cell margin (inches)
	 * - affects all table cells, is superceded by cell options
	 */
	margin?: Margin
	/**
	 * Height of table rows (inches)
	 * - single value is applied to every row equally based upon `h`
	 * - array of values in applied to each row in order
	 * @default rows of equal height based upon `h`
	 */
	rowH?: number | number[]
	/**
	 * DEV-ONLY diagnostic flag: when `true`, logs a verbose trace of the auto-paging
	 * calculations to the console. Intended for debugging table layout only; leave unset
	 * in production. Inherited by {@link TableToSlidesProps}.
	 * @default false
	 */
	verbose?: boolean
}
export interface TableCell {
	_type?: SlideObjectType.tablecell
	/** lines in this cell (autoPage) */
	_lines?: TableCell[][]
	/** `text` prop but guaranteed to hold "TableCell[]" */
	_tableCells?: TableCell[]
	/** height in EMU */
	_lineHeight?: number
	_hmerge?: boolean
	_vmerge?: boolean
	_rowContinue?: number
	/** origin cell of a colspan/rowspan span, set on the dummy `_hmerge`/`_vmerge` cells so they can
	 * inherit the origin's border/fill and render the merged region's outer edges */
	_spanOrigin?: TableCell

	/**
	 * Cell content: a plain string, or an array of `TableCell` runs for mixed formatting.
	 * (A `number` is still coerced to a string at runtime for plain-JS callers, but is no
	 * longer part of the type — pass `String(n)` from TypeScript.)
	 */
	text?: string | TableCell[]
	options?: TableCellProps
}
export interface TableRowSlide {
	rows: TableRow[]
	/**
	 * Per-row height (inches) aligned 1:1 with `rows`, derived from the original `rowH` array.
	 * Auto-paging splits/reorders rows across slides and inserts repeated headers, so the caller's
	 * `rowH[i]` (keyed by *original* row index) can no longer be applied by physical row index on
	 * each generated slide. This carries each output row's resolved height so a configured height
	 * follows its source row instead of being re-applied to whatever lands at that index.
	 * Entries are `undefined` where no explicit height was configured (auto-distributed height).
	 */
	rowH?: Array<number | undefined>
}
export type TableRow = TableCell[]
