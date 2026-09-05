/**
 * ts-pptx: the construct families every presentation carries
 *
 * The core tier. A composed presentation gets these whether it asks for them or not, so
 * `createPresentation()` with no families still authors text, shapes, images, groups and speaker
 * notes -- a deck, rather than a slide-shaped hole.
 *
 * The line is a judgement rather than a measurement: these five are what "a slide" means to most
 * callers, they are cheap, and a surface that made someone compose `text` before it could write a
 * word would be a worse default than one that costs a few kilobytes more. A family is core when
 * leaving it out would make the result something other than a deck; everything else is asked for.
 */

import { groupFamily } from './group.js'
import { imageFamily } from './image.js'
import { notesFamily } from './notes.js'
import { shapeFamily } from './shape.js'
import { textFamily } from './text.js'
import type { ConstructFamily } from './shared.js'

/** Every family a composed presentation carries without being asked. */
export const CORE_CONSTRUCT_FAMILIES = [
	textFamily,
	shapeFamily,
	imageFamily,
	groupFamily,
	notesFamily,
] as const satisfies readonly ConstructFamily[]
