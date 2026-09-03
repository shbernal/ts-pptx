/**
 * `a:bodyPr/@anchor` and `a:tcPr/@anchor`, and the write API's `valign` on the other side of it.
 *
 * Two tables stated the same correspondence in opposite directions and could not see each other:
 * the emitters' spelling table in `gen/drawingml/text-body.ts`, and its hand-written inverse in
 * `script/from-read/values.ts`, which maps a read deck's anchor back to the `valign` a generated
 * script writes. `script/` cannot import `gen/` — it has no such import anywhere, deliberately —
 * so the inverse could not be derived where it stood, and a token added to one table would have
 * meant a script that reads an anchor it cannot write back.
 *
 * The correspondence itself is a fact about the format, so it lives here, as one table with both
 * directions derived from it.
 *
 * This module keeps `ooxml/`'s no-runtime-imports property: {@link TEXT_ANCHORS} is the schema's
 * own list from a module with none either, and the {@link TextAnchor} reference is type-only, so
 * it is erased. That reference is still a real check — it is what ties the enum the write API
 * exposes to the tokens named here.
 */

import type { TextAnchor } from '../enums.js'
import { TEXT_ANCHORS } from './st-enums.js'

/**
 * The `ST_TextAnchoringType` tokens this library writes, checked against both the schema's list
 * and the public {@link TextAnchor} enum.
 *
 * `Extract` is that check twice over: a token here that is not in `ST_TextAnchoringType`, or not
 * in the enum, is dropped from the type and the table below stops compiling. Compile-time on
 * purpose — every value in the table is a literal, so a runtime check could never fail, and a
 * check that cannot fail is worse than none.
 *
 * The library writes three of the schema's five: `just` and `dist` are anchoring modes no
 * option surfaces.
 */
export type TextAnchorToken = Extract<`${TextAnchor}`, (typeof TEXT_ANCHORS)[number]>

/**
 * Every `valign` the write API accepts, the anchor token it means, and the other spellings the
 * definers between them have always taken.
 *
 * The three definers read `valign` three ways — `startsWith('b'|'m'|'t')`, a chain of
 * `.replace('top','t')…` that let any other string through *verbatim* into `anchor=`, and a
 * longer chain that also took `btm` and `center`. This table is the union of what they named;
 * anything else is a warning and an omitted attribute rather than an invalid one.
 *
 * `valign` leads each row because it is the direction with a preferred answer: several spellings
 * mean one anchor, and only one of them is what a generated script should write back.
 */
const VALIGN_TABLE = [
	{ valign: 'top', anchor: 't', aliases: ['t'] },
	{ valign: 'middle', anchor: 'ctr', aliases: ['c', 'ctr', 'center', 'm'] },
	{ valign: 'bottom', anchor: 'b', aliases: ['b', 'btm'] },
] as const satisfies ReadonlyArray<{ valign: string; anchor: TextAnchorToken; aliases: readonly string[] }>

/** Every spelling of `valign` the definers accept, mapped to the one anchor token it means. */
export const TEXT_ANCHOR_BY_VALIGN: Readonly<Record<string, TextAnchorToken>> = Object.fromEntries(
	VALIGN_TABLE.flatMap((row) => [row.valign, ...row.aliases].map((spelling) => [spelling, row.anchor]))
)

/**
 * An anchor token back to the `valign` a generated script writes — the preferred spelling of the
 * row, never one of its aliases.
 */
export const VALIGN_BY_ANCHOR: Readonly<Record<string, string>> = Object.fromEntries(
	VALIGN_TABLE.map((row) => [row.anchor, row.valign])
)
