/**
 * An `mc:AlternateContent` envelope: a Choice for consumers that understand one namespace, and an
 * optional Fallback for everyone else.
 *
 * Six emitters wrapped their markup this way (a transition's exact duration, a text shape holding
 * an equation, a chartEx frame, an OLE object, a 3D model, a zoom), and each restated the rule that
 * makes the envelope work: the prefix named in `Requires` has to be declared on the Choice itself,
 * or a consumer reading `Requires` cannot resolve it. Stated here once.
 */

import { OOXML_NS } from '../../ooxml/namespaces.js'
import { el, raw } from './el.js'

/**
 * Build the envelope.
 * @param opts.requires - the namespace the Choice needs: its prefix and URI, declared on the Choice
 * @param opts.choice - the already-serialized markup for a consumer that understands it
 * @param opts.fallback - the already-serialized markup for one that does not, when there is any
 */
export function alternateContentEl(opts: {
	requires: { prefix: string; uri: string }
	choice: string
	fallback?: string
}): string {
	const { prefix, uri } = opts.requires
	return el('mc:AlternateContent', { 'xmlns:mc': OOXML_NS.mc }, [
		raw(el('mc:Choice', { [`xmlns:${prefix}`]: uri, Requires: prefix }, raw(opts.choice))),
		...(opts.fallback === undefined ? [] : [raw(el('mc:Fallback', null, raw(opts.fallback)))]),
	])
}
