/**
 * `a:pPr/@algn` and `a:lvlNpPr/@algn`, and the write API's `align` on the other side of it.
 *
 * Three places stated the same correspondence: the paragraph emitter's switch in
 * `gen/drawingml/text-run.ts`, `masterAlignAttr` in `gen/slide/master.ts`, and the script
 * converter's inverse in `script/from-read/text.ts`, which cannot import `gen/`. One table here,
 * both directions derived from it, as `ooxml/text-anchor.ts` does for `valign`.
 *
 * This module keeps `ooxml/`'s no-runtime-imports property: {@link TEXT_ALIGN_TYPES} comes from a
 * module with none, and the {@link HAlign} reference is type-only, so it is erased.
 */

import type { HAlign } from '../types/core.js'
import type { TEXT_ALIGN_TYPES } from './st-enums.js'

/** An `ST_TextAlignType` token. */
export type TextAlignToken = (typeof TEXT_ALIGN_TYPES)[number]

/**
 * Every `align` the write API names and the token it writes. `satisfies` over `Record<HAlign, …>`
 * makes a member added to `HAlign` without a token a compile error. `justLow`, `dist` and
 * `thaiDist` have no `align` spelling.
 */
const ALGN_BY_ALIGN = {
	left: 'l',
	center: 'ctr',
	right: 'r',
	justify: 'just',
} as const satisfies Record<HAlign, TextAlignToken>

/** The `algn` token an `align` value writes, or `undefined` for a value the write API does not name. */
export function textAlignToken(align: string): TextAlignToken | undefined {
	return Object.hasOwn(ALGN_BY_ALIGN, align) ? ALGN_BY_ALIGN[align as HAlign] : undefined
}

/** An `algn` token back to the `align` a generated script writes. */
export const HALIGN_BY_TEXT_ALIGN: Readonly<Record<string, HAlign>> = Object.fromEntries(
	Object.entries(ALGN_BY_ALIGN).map(([align, token]) => [token, align as HAlign])
)
