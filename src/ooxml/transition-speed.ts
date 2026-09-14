/**
 * PowerPoint's coarse `p:transition/@spd` bucket for an exact duration.
 *
 * The write path (`gen/anim/transition.ts`) and the read model's transition setter
 * (`read/api/transition.ts`) each carried a byte-identical copy. A duration that lands in a
 * different bucket on the two sides would make a transition set through the read model differ
 * from the same transition authored fresh, so the boundaries live here once.
 *
 * This module keeps `ooxml/`'s no-runtime-imports property: the reference below is type-only.
 */

import type { TRANSITION_SPEEDS } from './st-enums.js'

/** An `ST_TransitionSpeed` token. */
export type TransitionSpeedToken = (typeof TRANSITION_SPEEDS)[number]

/**
 * Map an exact duration (ms) to the speed bucket it falls in.
 * @param durationMs - the transition's duration, in milliseconds
 */
export function transitionSpeedForDuration(durationMs: number): TransitionSpeedToken {
	if (durationMs <= 500) return 'fast'
	if (durationMs <= 1000) return 'med'
	return 'slow'
}
