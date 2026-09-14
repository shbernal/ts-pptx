/**
 * `a:ln/@cap`, and the write API's `cap` on the other side of it.
 *
 * `createLineCap` in `gen/drawingml/line.ts` and its inverse `CAP_TOKENS` in
 * `script/from-read/shape-paint.ts` stated the same three pairs in opposite directions. One table
 * here, both directions derived from it, as `ooxml/text-anchor.ts` does for `valign`.
 *
 * This module keeps `ooxml/`'s no-runtime-imports property: both references below are type-only.
 */

import type { LineCap } from '../types/chart.js'
import type { LINE_CAPS } from './st-enums.js'

/** An `ST_LineCap` token. */
export type LineCapToken = (typeof LINE_CAPS)[number]

/**
 * Every `cap` the write API names and the token it writes. `satisfies` over `Record<LineCap, …>`
 * makes a member added to `LineCap` without a token a compile error, and `ST_LineCap` has exactly
 * these three members, so the inverse covers every token too.
 */
const TOKEN_BY_LINE_CAP = {
	flat: 'flat',
	square: 'sq',
	round: 'rnd',
} as const satisfies Record<LineCap, LineCapToken>

/** The `cap` token a `LineCap` writes, or `undefined` for a value the write API does not name. */
export function lineCapToken(cap: string): LineCapToken | undefined {
	return Object.hasOwn(TOKEN_BY_LINE_CAP, cap) ? TOKEN_BY_LINE_CAP[cap as LineCap] : undefined
}

/** A `cap` token back to the `LineCap` that authors it. */
export const LINE_CAP_BY_TOKEN: Readonly<Record<string, LineCap>> = Object.fromEntries(
	Object.entries(TOKEN_BY_LINE_CAP).map(([cap, token]) => [token, cap as LineCap])
)
