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
export type TextHorzOverflowType = (typeof TEXT_HORZ_OVERFLOW)[number]

/**
 * `ST_TextAnchoringType` — `a:bodyPr/@anchor` and `a:tcPr/@anchor`.
 * @see ECMA-376 Part 1 §20.1.10.60
 */
export const TEXT_ANCHORS = ['t', 'ctr', 'b', 'just', 'dist'] as const
export type TextAnchoringType = (typeof TEXT_ANCHORS)[number]

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
