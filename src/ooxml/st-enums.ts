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
