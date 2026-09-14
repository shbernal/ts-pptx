/**
 * The fixed ratios between English Metric Units, points and inches.
 *
 * `units.ts` publishes them, but it also holds the unit converters, which warn and throw, so it
 * imports `diagnostics.ts` and `errors.ts` at run time. A schema fact in `ooxml/` that only needs
 * a ratio (`body-insets.ts`) took both along with it. This module imports nothing, and `units.ts`
 * re-exports it, so the public names are unchanged.
 */

/** EMU in one inch. */
export const EMU_PER_INCH = 914400
/** EMU in one point. */
export const EMU_PER_POINT = 12700
/** Points in one inch. */
export const POINTS_PER_INCH = 72
