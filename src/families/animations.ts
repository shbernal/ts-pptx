/**
 * ts-pptx: the animations construct family
 *
 * Preset build animations (`addAnimation`). Authoring one is only a push onto the slide's list --
 * what the family is really worth is the timing tree that list turns into, and that is still
 * emitted from the slide writer rather than from here.
 *
 * **That was measured, not assumed.** Moving the tree here would need a fourth seam: per-slide XML
 * outside a shape, handed to the slide writer the way the packager is handed part contributors,
 * and ordered, because `p:transition` and `p:timing` are position-significant children of `p:sld`.
 * What it would buy is `gen/anim/animation.ts` plus `gen/anim/transition.ts`, which minify to 6.6
 * kB and gzip to 2.3 kB -- about 3.6% of the composed tier's 63 kB entry chunk. And it would not
 * even buy all of that: `gen/anim/timing.ts` stays reachable either way, because looping *media*
 * goes through the same tree. A new architectural seam carrying a byte-significant ordering rule
 * is not worth 2.3 kB. Re-measure before reopening it; that number is what would have to change.
 */

import type { ConstructFamily } from './shared.js'

export const animationsFamily = {
	name: 'animations',
	authors: {
		addAnimation(slide, options) {
			slide._animations.push(options)
		},
	},
} satisfies ConstructFamily
