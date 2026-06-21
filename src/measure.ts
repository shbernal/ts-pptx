/**
 * `pptxgenjs/measure` — the calibrated text-measurement engine as a standalone,
 * public surface, so a consumer can lay out its own geometry (grow a card, reflow
 * a grid, detect overflow) **before export** without a `PptxGenJS` instance.
 *
 * This is the same wrap model + solvers the export-time autofit bake uses (see
 * `docs/measured-text-fit.md`), so a layout-time prediction never disagrees with
 * the baked result. For the ergonomic, inches-based path use
 * `pptx.measureText()` / `pptx.overflowsBox()` on a presentation instance; reach
 * for these primitives when you need to build your own resolver/registry and
 * measure standalone.
 *
 * `opentype.js` (used by `parseFontMetrics`) is imported lazily, so importing this
 * module does not pull it into the bundle until a font is actually parsed.
 */

// Pure wrap model + solvers (points in, points out) and their types/constants.
export {
	measureLayout,
	measureHeightPt,
	solveShrink,
	solveResize,
	SINGLE_LINE_PITCH,
	FONT_SCALE_STEP_PCT,
	MIN_FONT_SCALE_PCT,
	WIDTH_SAFETY_FACTOR,
	HEIGHT_SAFETY_FACTOR,
	type LayoutResult,
	type FitRun,
	type FitParagraph,
	type FitBox,
	type MetricsResolver,
	type ShrinkResult,
	type ShrinkOutcome,
	type ResizeOutcome,
} from './text-fit.js'

// Font-metrics provider + registry, so a consumer can build its own resolver and
// measure without threading a whole PptxGenJS instance through its layout context.
export {
	parseFontMetrics,
	getHeuristicFontMetrics,
	FontMetricsRegistry,
	type FontMetrics,
} from './font-metrics.js'

// Higher-level helpers shared with the export pass: the public-props → FitParagraph
// converter, the registry-backed resolver factory, and the inches-based measure used
// by `pptx.measureText()` (callable here with any FontMetricsRegistry).
export { buildFitParagraphs, makeRegistryResolver, measureText } from './measure-fit.js'
export type { MeasureTextOptions, TextMeasurement, OverflowBoxOptions } from './core-interfaces.js'
