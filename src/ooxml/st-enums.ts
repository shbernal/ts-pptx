/**
 * `ST_` simple-type enumerations from ECMA-376, declared once as the single source of both the
 * TypeScript union a caller is checked against and the runtime list a value is validated against.
 *
 * **The problem this solves.** Each of these used to exist twice: a hand-written string union in
 * `src/types/` (what the compiler enforces) and a hand-written array in `src/gen/` or `src/read/`
 * (what the validator checks at run time). Nothing tied the two together, so adding a member to
 * one and not the other produced either a value TypeScript accepts and the validator drops, or a
 * value the validator passes and the type rejects. Two of them had already diverged in *order*,
 * which is harmless, but nothing would have caught a divergence in *content*.
 *
 * **The shape.** Every entry is a `readonly` tuple declared `as const`, with its union derived via
 * `(typeof X)[number]`. Widening the tuple widens the union in the same edit — they cannot drift.
 * This is the idiom already used for `EMBEDDED_FONT_SLOTS` and `EXTRA_SHAPE_PRESETS`; this module
 * brings the stragglers onto it.
 *
 * **Why this module is neither `gen/` nor `read/`.** `src/types/` needs the unions for the public
 * write API, `src/gen/` needs the tuples to vet a value before emitting it, and `src/read/` needs
 * them to vet a value before writing it into a loaded part. It is a fact about the schema, so it
 * belongs to none of the three.
 *
 * Values are listed in **schema declaration order**, which is the order the diagnostics print
 * when they tell a caller what is legal.
 *
 * **The chart option bag reads seven of these directly**: `BarDirection`, `BarGrouping`,
 * `Bar3DShape`, `LegendPosition`, `DisplayBlanksAs`, `DataLabelPosition` and `LineDataSymbol`.
 * They were the half of this consolidation that went unfinished for a while. `src/types/chart.ts`
 * hand-wrote four of those unions beside the option that takes them and typed the other three as
 * bare `string`, so `barDir` carried no compile-time check at all, and `dataLabelPosition`'s
 * hand-written union was missing `inBase` — a value the validator accepts on a stacked bar, so
 * the two spellings had already diverged in content. Both are gone: for every option here, the
 * type a caller is checked against and the tuple the validator checks are one declaration.
 */

/**
 * `ST_BevelPresetType` — the bevel profile applied to a 3-D edge.
 * @see ECMA-376 Part 1 §20.1.10.9
 */
export const BEVEL_PRESETS = [
	'relaxedInset',
	'circle',
	'slope',
	'cross',
	'angle',
	'softRound',
	'convex',
	'coolSlant',
	'divot',
	'riblet',
	'hardEdge',
	'artDeco',
] as const
export type BevelPresetType = (typeof BEVEL_PRESETS)[number]

/**
 * `ST_PresetMaterialType` — how a 3-D surface responds to light.
 * @see ECMA-376 Part 1 §20.1.10.50
 */
export const PRESET_MATERIALS = [
	'legacyMatte',
	'legacyPlastic',
	'legacyMetal',
	'legacyWireframe',
	'matte',
	'plastic',
	'metal',
	'warmMatte',
	'translucentPowder',
	'powder',
	'dkEdge',
	'softEdge',
	'clear',
	'flat',
	'softmetal',
] as const
export type PresetMaterialType = (typeof PRESET_MATERIALS)[number]

/**
 * `ST_LightRigType` — the preset lighting scene.
 * @see ECMA-376 Part 1 §20.1.10.35
 */
export const LIGHT_RIGS = [
	'legacyFlat1',
	'legacyFlat2',
	'legacyFlat3',
	'legacyFlat4',
	'legacyNormal1',
	'legacyNormal2',
	'legacyNormal3',
	'legacyNormal4',
	'legacyHarsh1',
	'legacyHarsh2',
	'legacyHarsh3',
	'legacyHarsh4',
	'threePt',
	'balanced',
	'soft',
	'harsh',
	'flood',
	'contrasting',
	'morning',
	'sunrise',
	'sunset',
	'chilly',
	'freezing',
	'flat',
	'twoPt',
	'glow',
	'brightRoom',
] as const
export type LightRigType = (typeof LIGHT_RIGS)[number]

/**
 * `ST_LightRigDirection` — where the light rig sits relative to the scene.
 * @see ECMA-376 Part 1 §20.1.10.33
 */
export const LIGHT_RIG_DIRECTIONS = ['tl', 't', 'tr', 'l', 'r', 'bl', 'b', 'br'] as const
export type LightRigDirection = (typeof LIGHT_RIG_DIRECTIONS)[number]

/**
 * `ST_TextHorzOverflowType` — `a:bodyPr/@horzOverflow` and `a:tcPr/@horzOverflow`.
 * @see ECMA-376 Part 1 §20.1.10.68
 */
export const TEXT_HORZ_OVERFLOW = ['clip', 'overflow'] as const

/**
 * `ST_TextAnchoringType` — `a:bodyPr/@anchor` and `a:tcPr/@anchor`.
 * @see ECMA-376 Part 1 §20.1.10.60
 */
export const TEXT_ANCHORS = ['t', 'ctr', 'b', 'just', 'dist'] as const

/**
 * `ST_TextVerticalType` — `a:bodyPr/@vert` and `a:tcPr/@vert`.
 * @see ECMA-376 Part 1 §20.1.10.83
 */
export const TEXT_VERTICAL = [
	'horz',
	'vert',
	'vert270',
	'wordArtVert',
	'eaVert',
	'mongolianVert',
	'wordArtVertRtl',
] as const
export type TextVerticalType = (typeof TEXT_VERTICAL)[number]

/**
 * `ST_LineEndType` — the arrowhead at either end of a stroke (`a:headEnd/@type`,
 * `a:tailEnd/@type`).
 * @see ECMA-376 Part 1 §20.1.10.33
 */
export const LINE_END_TYPES = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'] as const
export type LineEndType = (typeof LINE_END_TYPES)[number]

/**
 * `ST_TextShapeType` — the preset path text is bent along (`a:prstTxWarp/@prst`), which
 * PowerPoint calls WordArt.
 *
 * `textNoShape` is the identity member: it is legal, and it warps nothing. The library does not
 * narrow this set — every member is a shape PowerPoint draws.
 * @see ECMA-376 Part 1 §20.1.10.76
 */
export const TEXT_SHAPE_TYPES = [
	'textNoShape',
	'textPlain',
	'textStop',
	'textTriangle',
	'textTriangleInverted',
	'textChevron',
	'textChevronInverted',
	'textRingInside',
	'textRingOutside',
	'textArchUp',
	'textArchDown',
	'textCircle',
	'textButton',
	'textArchUpPour',
	'textArchDownPour',
	'textCirclePour',
	'textButtonPour',
	'textCurveUp',
	'textCurveDown',
	'textCanUp',
	'textCanDown',
	'textWave1',
	'textWave2',
	'textDoubleWave1',
	'textWave4',
	'textInflate',
	'textDeflate',
	'textInflateBottom',
	'textDeflateBottom',
	'textInflateTop',
	'textDeflateTop',
	'textDeflateInflate',
	'textDeflateInflateDeflate',
	'textFadeRight',
	'textFadeLeft',
	'textFadeUp',
	'textFadeDown',
	'textSlantUp',
	'textSlantDown',
	'textCascadeUp',
	'textCascadeDown',
] as const
export type TextShapeType = (typeof TEXT_SHAPE_TYPES)[number]

/**
 * `ST_PresetLineDashVal` — the dash pattern of a line or border (`a:prstDash/@val`).
 *
 * The `sys*` members are the "system" dashes whose exact dot/gap ratio the renderer picks; the
 * rest are defined by the spec. Both `BorderProps.dashType` and `ShapeLineProps.dashType` accept
 * this whole set, so a caller who knows one knows the other.
 * @see ECMA-376 Part 1 §20.1.10.49
 */
export const PRESET_LINE_DASHES = [
	'solid',
	'dot',
	'dash',
	'lgDash',
	'dashDot',
	'lgDashDot',
	'lgDashDotDot',
	'sysDash',
	'sysDot',
	'sysDashDot',
	'sysDashDotDot',
] as const
export type PresetLineDashVal = (typeof PRESET_LINE_DASHES)[number]

/**
 * `ST_BarDir` — whether a bar/column plot runs horizontally or vertically (`c:barDir/@val`).
 * @see ECMA-376 Part 1 §21.2.3.3
 */
export const BAR_DIRECTIONS = ['bar', 'col'] as const
export type BarDirection = (typeof BAR_DIRECTIONS)[number]

/**
 * `ST_BarGrouping` — how a bar/column plot stacks its series (`c:grouping/@val` inside
 * `c:barChart`/`c:bar3DChart`).
 * @see ECMA-376 Part 1 §21.2.3.4
 */
export const BAR_GROUPINGS = ['percentStacked', 'clustered', 'standard', 'stacked'] as const
export type BarGrouping = (typeof BAR_GROUPINGS)[number]

/**
 * `ST_Grouping` — the same decision for the plots that have no clustered form
 * (`c:areaChart`, `c:lineChart`).
 * @see ECMA-376 Part 1 §21.2.3.17
 */
export const GROUPINGS = ['percentStacked', 'standard', 'stacked'] as const
export type Grouping = (typeof GROUPINGS)[number]

/**
 * `ST_Shape` — the solid a 3-D bar is drawn as (`c:shape/@val`).
 * @see ECMA-376 Part 1 §21.2.3.35
 */
export const BAR_3D_SHAPES = ['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'] as const
export type Bar3DShape = (typeof BAR_3D_SHAPES)[number]

/**
 * `ST_RadarStyle` — how a radar plot is drawn (`c:radarStyle/@val`).
 * @see ECMA-376 Part 1 §21.2.3.33
 */
export const RADAR_STYLES = ['standard', 'marker', 'filled'] as const
export type RadarStyle = (typeof RADAR_STYLES)[number]

/**
 * This library's own names for two of the three radar styles, mapped to the wire spelling each
 * one means. They are the PowerPoint UI's words ("Radar", "Radar with Markers"), which is why
 * they were chosen; `filled` needed no rename and is already the schema's own spelling, which is
 * what made the other two read as typos rather than as a second vocabulary.
 *
 * Both spellings are accepted on the public option and normalized to the wire one at definition
 * time. Before that, `radarStyle: 'marker'` — the schema's word, and the obvious guess for anyone
 * reading a chart part — failed the enum check and silently fell back to `standard`, so a caller
 * asking for markers got a plain radar.
 */
export const RADAR_STYLE_ALIASES = { radar: 'standard', markers: 'marker' } as const
export type RadarStyleAlias = keyof typeof RADAR_STYLE_ALIASES
/**
 * The alias spellings alone, for the definer's enum check. Derived from the map rather than
 * written out again — the whole point of this module is that the list a value is validated
 * against and the union it is typed against cannot drift apart.
 */
export const RADAR_STYLE_ALIAS_NAMES = Object.keys(RADAR_STYLE_ALIASES) as readonly RadarStyleAlias[]

/**
 * `ST_LegendPos` — where the legend sits relative to the plot area (`c:legendPos/@val`).
 * @see ECMA-376 Part 1 §21.2.3.24
 */
export const LEGEND_POSITIONS = ['b', 'tr', 'l', 'r', 't'] as const
export type LegendPosition = (typeof LEGEND_POSITIONS)[number]

/**
 * `ST_DispBlanksAs` — what a plot draws where a value is missing (`c:dispBlanksAs/@val`).
 * @see ECMA-376 Part 1 §21.2.3.10
 */
export const DISPLAY_BLANKS_AS = ['span', 'gap', 'zero'] as const
export type DisplayBlanksAs = (typeof DISPLAY_BLANKS_AS)[number]

/**
 * `ST_DLblPos` — where a data label sits relative to its point (`c:dLblPos/@val`).
 *
 * The whole set is legal on the attribute, but only a subset of it is legal *per plot type* —
 * PowerPoint offers a pie chart `bestFit`/`ctr`/`inEnd`/`outEnd` and nothing else. Those subsets
 * are a fact about the chart types rather than about the attribute, so they live beside the
 * chart definer that applies them; this is the value space they are drawn from.
 * @see ECMA-376 Part 1 §21.2.3.13
 */
export const DATA_LABEL_POSITIONS = ['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'] as const
export type DataLabelPosition = (typeof DATA_LABEL_POSITIONS)[number]

/**
 * `ST_MarkerStyle`, **narrowed** to what the library authors (`c:symbol/@val`).
 *
 * The schema also has `plus`, `star`, `x`, `picture` and `auto`. Neither PowerPoint 2013 nor
 * PowerPoint Online draws the first three, `picture` needs a `c:spPr` blip this API has no
 * spelling for, and `auto` is what leaving the option off already means — so the accepted set is
 * deliberately smaller than the type. Widen it only with evidence that the member renders.
 * @see ECMA-376 Part 1 §21.2.3.27
 */
export const LINE_DATA_SYMBOLS = ['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'] as const
export type LineDataSymbol = (typeof LINE_DATA_SYMBOLS)[number]

/**
 * The twelve theme colour slots (`a:clrScheme` children) a `p:clrMap` token can point at.
 *
 * Not an `ST_` type — these are element names inside `a:CT_ColorScheme`, which the schema
 * declares as a fixed sequence rather than an enumeration. The effect is the same: a closed,
 * ordered set of twelve names, in the order the scheme declares them.
 * @see ECMA-376 Part 1 §20.1.6.2
 */
export const THEME_COLOR_SLOTS = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const
export type ThemeColorSlot = (typeof THEME_COLOR_SLOTS)[number]

/**
 * The twelve colour-map tokens (`p:clrMap` attributes), in schema attribute order.
 *
 * A slide's `schemeClr val="tx1"` names a *token*, not a slot: the master's `p:clrMap` is the
 * indirection that sends it to a {@link ThemeColorSlot}. Which is why the two lists look alike
 * and are not the same list — `bg1`/`tx1` are tokens, `lt1`/`dk1` are slots.
 * @see ECMA-376 Part 1 §19.3.1.6
 */
export const COLOR_MAP_TOKENS = [
	'bg1',
	'tx1',
	'bg2',
	'tx2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const
export type ColorMapToken = (typeof COLOR_MAP_TOKENS)[number]

/**
 * The identity colour map: the `p:clrMap` this library emits on every master and notes master,
 * and the one the script printer recognises so it can skip printing a map that is the default.
 *
 * "Identity" up to the light/dark naming: each token points at the slot of the same name, with
 * `bg`/`tx` reading as light/dark. A deck that maps them anywhere else — swapping the light and
 * dark slots is the usual case — resolves every `schemeClr` in the deck differently.
 *
 * Typed `Record<ColorMapToken, ThemeColorSlot>` so a token added to {@link COLOR_MAP_TOKENS}
 * without a mapping, or a slot misspelled, is a compile error rather than a theme that opens
 * fine and paints one script in the wrong colour. Declared in the same order as the tokens,
 * because it is emitted by iterating this object and the attribute order is part of the bytes.
 */
export const DEFAULT_COLOR_MAP: Readonly<Record<ColorMapToken, ThemeColorSlot>> = {
	bg1: 'lt1',
	tx1: 'dk1',
	bg2: 'lt2',
	tx2: 'dk2',
	accent1: 'accent1',
	accent2: 'accent2',
	accent3: 'accent3',
	accent4: 'accent4',
	accent5: 'accent5',
	accent6: 'accent6',
	hlink: 'hlink',
	folHlink: 'folHlink',
}
