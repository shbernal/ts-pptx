/**
 * ts-pptx: the bullet elements, `a:buChar` and `a:buAutoNum`
 *
 * Two emitters write bullets: the slide text runs (`text-run.ts`) and the slide master's text
 * styles (`gen/slide/master.ts`). They built both elements themselves and disagreed. The slide path
 * checked a `characterCode` before using it, while the master interpolated it into the attribute
 * raw, so a code holding a quote wrote a master that does not parse. These builders are the only
 * place either element is built. What stays with each caller is its default font, which is a real
 * difference between the two rather than drift.
 */

import { BulletType } from '../../enums.js'
import { warn } from '../../diagnostics.js'
import { clampRangedInput } from '../../units-internal.js'
import { voidEl } from '../oxml/el.js'

/** A bullet `characterCode`: four hex digits naming a code point. */
const CHARACTER_CODE = /^[0-9A-Fa-f]{4}$/

/** `ST_TextBulletStartAtNum`'s range. */
const START_AT_MIN = 1
const START_AT_MAX = 32767

/**
 * `<a:buChar>`: the glyph a caller's `characterCode` names, or `glyph` when the caller named none.
 *
 * A template rather than `voidEl`, because `char` carries a pre-escaped numeric character reference
 * (`&#x2022;`, see `BulletType`) and the builder would escape its `&` a second time, rendering the
 * reference as literal text. That is why the code is checked before it gets near the attribute: one
 * that is not four hex digits warns and becomes {@link BulletType.DEFAULT}.
 * @param characterCode - the caller's `characterCode`, if they gave one
 * @param glyph - the glyph when they gave none; a library constant, already safe in an attribute
 * @param label - the option the code came from, opening the warning
 */
export function buCharEl(characterCode: string | undefined, glyph: string, label: string): string {
	let char = glyph
	if (characterCode) {
		if (CHARACTER_CODE.test(characterCode)) {
			char = `&#x${characterCode};`
		} else {
			warn(
				'bullet/invalid-character-code',
				`${label} \`characterCode\` should be a 4-digit unicode character (ex: 22AB), got ${JSON.stringify(characterCode)}; using the default bullet`
			)
			char = BulletType.DEFAULT
		}
	}
	return `<a:buChar char="${char}"/>`
}

/**
 * `<a:buAutoNum>` for an auto-number scheme, starting at `startAt` when one is stated.
 *
 * `startAt` is an `ST_TextBulletStartAtNum`, a whole number from 1 to 32767: a fractional one rounds,
 * one outside the range clamps with a warning, and one that is not a number throws.
 * @param type - the `ST_TextAutonumberScheme` token
 * @param startAt - the number the list starts at; omitted from the element when `undefined`
 * @param label - the option the value came from, opening the warning
 */
export function buAutoNumEl(type: string, startAt: number | undefined, label: string): string {
	const start =
		startAt === undefined
			? null
			: clampRangedInput(
					Math.round(startAt),
					START_AT_MIN,
					START_AT_MAX,
					'bullet/start-at-out-of-range',
					`${label} \`numberStartAt\``,
					'bullet/start-at-not-a-number'
				)
	return voidEl('a:buAutoNum', { type, startAt: start })
}
