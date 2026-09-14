/**
 * ts-pptx: the text construct family
 *
 * Text boxes, and the shape-with-text descriptors a slide master or a group is written with. The
 * cheapest family and the one every tier carries, but it is a family like any other: a list with an
 * exception in it is a list nobody trusts.
 */

import { SlideObjectType } from '../enums.js'
import { InvalidOptionError } from '../errors.js'
import { renderTextObject } from '../gen/slide/objects/text.js'
import { addTextDefinition } from '../gen/define/text.js'
import type { ConstructFamily } from './shared.js'

/**
 * Refuse a `text` that is none of the forms the signature takes. A single run object,
 * `{ text: 'x' }`, is the one a caller reaches for, and it failed inside the definer with a
 * `TypeError`.
 * @param text - the caller's `text`
 * @param call - the method or descriptor the text was authored through, opening the message
 */
function requireTextForm(text: unknown, call: string): void {
	if (text == null || typeof text === 'string' || typeof text === 'number' || Array.isArray(text)) return
	throw new InvalidOptionError(
		'text/invalid-text',
		`${call}: \`text\` takes a string, a number or an array of runs, got ${typeof text}. Wrap a single run in an array: \`[{ text: 'x' }]\`.`
	)
}

export const textFamily = {
	name: 'text',
	authors: {
		addText(slide, text, options) {
			requireTextForm(text, 'addText')
			// The bare-string form is wrapped into a one-run list carrying NO options of its own -- the
			// same shape a caller writing `[{ text: 'x' }]` hands in, and for the same reason: a bare
			// string authors no *run*, so its run states nothing and inherits everything.
			//
			// It used to hand `options` to both the shape and the run, which made one object play both
			// roles. Every key on it was then read twice, once as a shape property and once as a run
			// property, and `shadow` is a key that means something different in each place: the shape's
			// `<a:effectLst>` and the run's are two separate gestures in PowerPoint (Shape Effects vs
			// Text Effects, `shadow-shape-vs-text.pptx`), and `addText('hi', { shadow })` emitted both.
			// A run inherits what a run inherits -- `RUN_INHERITABLE_OPTIONS` and the paragraph list
			// beside it in `gen/drawingml/text-run.ts` -- and nothing else.
			const runs = typeof text === 'string' || typeof text === 'number' ? [{ text }] : text
			addTextDefinition(slide, runs, options || {}, false)
		},
	},
	children: {
		text(target, child) {
			requireTextForm(child.text, 'a `text` descriptor')
			addTextDefinition(
				target,
				Array.isArray(child.text) ? child.text : [{ text: child.text }],
				child.options || {},
				false
			)
		},
	},
	// A text box and a layout placeholder are one emitter with two entry kinds: the placeholder
	// differences it does draw come off `ctx.placeholder`, not off the object's `_type`.
	renderers: {
		[SlideObjectType.placeholder]: renderTextObject,
		[SlideObjectType.text]: renderTextObject,
	},
} satisfies ConstructFamily
