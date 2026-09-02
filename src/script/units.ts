/**
 * ts-pptx: the script converter's unit conversions, on the seam between its two halves.
 *
 * `docs/architecture.md` states the seam: "`from-read/` knows OOXML and the read model,
 * `print/` knows only strings, and neither can see the other." Both halves nonetheless have
 * to turn an EMU slide size into inches — `ir.ts` keeps `slideSize` in EMU deliberately,
 * because "converting here would introduce a rounding decision at the wrong layer" — and the
 * printer was reaching across the seam into `from-read/values.ts` to do it.
 *
 * So the conversion lives here, beside `ir.ts`, which is the module both halves already
 * import. Nothing in here knows about the read model or about strings.
 */

import { EMU_PER_INCH } from '../units.js'

/**
 * Decimal places used for the inch-typed options.
 *
 * The proven minimum for an EMU-exact round-trip, not a formatting preference: `appendSlides`
 * compares slide sizes exactly and throws when they differ, so a generated script's
 * `defineLayout` has to land back on the source EMU.
 */
export const INCH_DECIMALS = 6

/**
 * Geometry as inches, for the options that reject a raw-EMU string. Rounded to
 * {@link INCH_DECIMALS}, which is EMU-exact on the way back.
 */
export function inches(emuValue: number): number {
	return Number((emuValue / EMU_PER_INCH).toFixed(INCH_DECIMALS))
}
