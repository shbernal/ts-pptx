/**
 * `a:hlinkClick` — the one reading of a click hyperlink, wherever it hangs.
 *
 * DrawingML puts the same element in two places: on a run's `a:rPr`, which links a span of text,
 * and on a shape's `p:cNvPr`, which links the whole shape. They are the same element with the same
 * attributes and the same relationship resolution, so they are read here once. The run's reader
 * used to hold the only copy, and the shape's reader did not exist -- a shape-level link read back
 * as nothing at all, so a replica lost it with nothing to say so.
 */

import { attr, type Element, firstChild } from '../oxml/dom.js'
import type { Relationships } from '../opc/relationships.js'

/**
 * A click hyperlink (`a:hlinkClick`), on a run or on a shape. A URL link carries an external
 * {@link url}; a slide jump carries the internal {@link targetPartName} (the linked slide's part)
 * alongside its `hlinksldjump` {@link action}. `tooltip` and `relId` are surfaced when present.
 */
export interface Hyperlink {
	/** External URL target (its `@r:id` resolves to a `TargetMode="External"` rel), or `null` for an internal/action-only link. */
	url: string | null
	/** Absolute partname of an internal target (e.g. the slide a jump points at), or `null`. */
	targetPartName: string | null
	/** Navigation action token (`@action`, e.g. `ppaction://hlinksldjump`), or `null` when absent/empty. */
	action: string | null
	/** Tooltip text (`@tooltip`), or `null` when absent/empty. */
	tooltip: string | null
	/** The relationship id (`@r:id`) backing the link, or `null` when the link is action-only. */
	relId: string | null
}

/**
 * Read the `a:hlinkClick` child of `parent`, or `null` when there is none.
 *
 * Without `rels` only the raw `@r:id`/`@action`/`@tooltip` are reported and the target stays
 * `null`: an id means nothing outside the part that declares it, and guessing at one would be
 * worse than saying nothing. That is the case a SmartArt drawing's cached text is read in.
 * @param parent - the element that may carry the link (`a:rPr` or `p:cNvPr`)
 * @param rels - the owning part's relationships, or `null` when they are not available
 */
export function readHyperlink(parent: Element | null, rels: Relationships | null): Hyperlink | null {
	const hlink = parent && firstChild(parent, 'a:hlinkClick')
	if (!hlink) return null
	const relId = attr(hlink, 'r:id') || null
	const action = attr(hlink, 'action') || null
	const tooltip = attr(hlink, 'tooltip') || null
	let url: string | null = null
	let targetPartName: string | null = null
	if (relId && rels) {
		const rel = rels.get(relId)
		if (rel?.targetMode === 'External') url = rel.target
		else if (rel) targetPartName = rels.resolveTarget(relId)
	}
	return { url, targetPartName, action, tooltip, relId }
}
