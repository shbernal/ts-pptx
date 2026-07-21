/**
 * PptxGenJS: Hyperlink Relationship Registration
 *
 * Walks the text / table-cell object tree from `addText()` / `addTable()` and registers a
 * `hyperlink`-type slide rel for every hyperlink found, stamping the resolved `_rId` back onto
 * each hyperlink so serialization can emit `r:id`. Shared by the shape, text and table layers.
 */
import { SlideObjectType } from '../../core-enums.js'
import type { ObjectOptions, TableCell, TextProps, TextPropsOptions } from '../../core-interfaces.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { getNewRelId } from '../../gen-utils.js'

type HyperlinkTextObject = (TextProps | SlideObject | TableCell) & {
	options?: TextPropsOptions | ObjectOptions
	text?: string | number | TextProps[] | TableCell[]
}

/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {PresSlideInternal} target - slide object that any hyperlinks will be be added to
 * @param {number | string | TextProps | TextProps[] | TableCell[][]} text - text to parse
 */
export function createHyperlinkRels(
	target: PresSlideInternal,
	text: number | string | SlideObject | TextProps | TextProps[] | TableCell[] | TableCell[][],
	options?: TextPropsOptions[]
): void {
	let textObjs: Array<HyperlinkTextObject | TableCell[]> = []

	// Only text objects can have hyperlinks, bail when text param is plain text
	if (typeof text === 'string' || typeof text === 'number') return
	// IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
	else if (Array.isArray(text)) textObjs = text
	else if (typeof text === 'object') textObjs = [text]

	textObjs.forEach((text: HyperlinkTextObject | TableCell[], idx: number) => {
		// NOTE: `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
		if (Array.isArray(text)) {
			const cellOpts: TextPropsOptions[] = []
			text.forEach((tablecell) => {
				if (tablecell.options) {
					cellOpts.push(tablecell.options)
				}
			})
			createHyperlinkRels(target, text, cellOpts)
			return
		}

		// IMPORTANT: `options` are lost due to recursion/copy!
		if (options && options[idx] && options[idx].hyperlink) text.options = { ...text.options, ...options[idx] }
		if (Array.isArray(text.text)) {
			createHyperlinkRels(target, text.text, options && options[idx] ? [options[idx]] : undefined)
		} else if (
			text &&
			typeof text === 'object' &&
			text.options &&
			text.options.hyperlink &&
			!text.options.hyperlink._rId
		) {
			const hyperlink = text.options.hyperlink
			if (typeof hyperlink !== 'object') {
				console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ")
			} else if (!hyperlink.url && !hyperlink.slide && !hyperlink.action) {
				console.log("ERROR: 'hyperlink requires either: `url`, `slide`, or `action`'")
			} else if (hyperlink.action && !hyperlink.url && !hyperlink.slide) {
				// Navigation action button: the `ppaction://hlinkshowjump` action is self-contained,
				// so there is no relationship to register (emitter writes `r:id=""`).
			} else {
				const relId = getNewRelId(target)

				target._rels.push({
					type: SlideObjectType.hyperlink,
					data: hyperlink.slide ? 'slide' : 'dummy',
					rId: relId,
					// `Target` is stored RAW; every emitter escapes it. See the note on `SlideRel.Target`.
					Target: hyperlink.url ? hyperlink.url : String(hyperlink.slide),
				})

				hyperlink._rId = relId
			}
		} else if (
			text &&
			typeof text === 'object' &&
			text.options &&
			text.options.hyperlink &&
			text.options.hyperlink._rId
		) {
			const hyperlink = text.options.hyperlink
			const hyperlinkRelId = hyperlink._rId
			// NOTE: auto-paging will create new slides, but skip above as _rId exists, BUT this is a new slide, so add rels!
			if (hyperlinkRelId && !target._rels.some((rel) => rel.rId === hyperlinkRelId)) {
				target._rels.push({
					type: SlideObjectType.hyperlink,
					data: hyperlink.slide ? 'slide' : 'dummy',
					rId: hyperlinkRelId,
					// `Target` is stored RAW; every emitter escapes it. See the note on `SlideRel.Target`.
					Target: hyperlink.url ? hyperlink.url : String(hyperlink.slide),
				})
			}
		}
	})
}
