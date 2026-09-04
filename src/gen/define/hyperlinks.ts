/**
 * ts-pptx: Hyperlink Relationship Registration
 *
 * Walks the text / table-cell object tree from `addText()` / `addTable()` and registers a
 * `hyperlink`-type slide rel for every hyperlink found, stamping the resolved `_rId` back onto
 * each hyperlink so serialization can emit `r:id`. Shared by the shape, text and table layers.
 */
import { SlideObjectType } from '../../enums.js'
import type { HyperlinkProps, ObjectOptions, TextProps, TextPropsOptions } from '../../types/index.js'
import type {
	HyperlinkPropsInternal,
	PresSlideInternal,
	SlideObject,
	SlideRel,
	TableCellInternal,
} from '../../types/internal.js'
import { getNewRelId } from '../utils.js'

/**
 * The relationship record one hyperlink serializes to.
 *
 * `data` distinguishes the two targets a hyperlink rel can have — `'slide'` for an internal
 * slide-to-slide link, `'dummy'` for an external URL — and `Target` is stored RAW, because
 * every emitter escapes it on the way out (see the note on `SlideRel.Target`).
 *
 * Four sites built this record by hand: the two below, `addImage`'s own hyperlink, and the
 * notes-part rels, which allocate from their own reserved id space and so cannot share the
 * registration below but can share this.
 * @param rId - the relationship id already allocated for this hyperlink
 * @param hyperlink - the caller's hyperlink; only `url`/`slide` are read
 */
export function hyperlinkRel(rId: number, hyperlink: HyperlinkProps): SlideRel {
	return {
		type: SlideObjectType.hyperlink,
		data: hyperlink.slide ? 'slide' : 'dummy',
		rId,
		Target: hyperlink.url ? hyperlink.url : String(hyperlink.slide),
	}
}

/**
 * Mint a fresh rel id for `hyperlink`, register it on `target`, and stamp the id back onto
 * the hyperlink so the emitter can write `r:id`.
 *
 * The id comes from {@link getNewRelId}, which skips every id already held on the slide.
 * `addImage` used to increment the image's own id instead, which is how an SVG picture — a
 * pair that already consumes two ids — ended up sharing the second of them with its
 * hyperlink and emitting a duplicate `Relationship Id`.
 * @returns the allocated relationship id
 */
export function registerHyperlinkRel(target: PresSlideInternal, hyperlink: HyperlinkPropsInternal): number {
	const relId = getNewRelId(target)
	target._rels.push(hyperlinkRel(relId, hyperlink))
	hyperlink._rId = relId
	return relId
}

type HyperlinkTextObject = (TextProps | SlideObject | TableCellInternal) & {
	options?: TextPropsOptions | ObjectOptions
	text?: string | number | TextProps[] | TableCellInternal[]
}

/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {PresSlideInternal} target - slide object that any hyperlinks will be be added to
 * @param {number | string | TextProps | TextProps[] | TableCellInternal[][]} text - text to parse
 */
export function createHyperlinkRels(
	target: PresSlideInternal,
	text: number | string | SlideObject | TextProps | TextProps[] | TableCellInternal[] | TableCellInternal[][],
	options?: TextPropsOptions[]
): void {
	let textObjs: Array<HyperlinkTextObject | TableCellInternal[]> = []

	// Only text objects can have hyperlinks, bail when text param is plain text
	if (typeof text === 'string' || typeof text === 'number') return
	// IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
	else if (Array.isArray(text)) textObjs = text
	else if (typeof text === 'object') textObjs = [text]

	textObjs.forEach((text: HyperlinkTextObject | TableCellInternal[], idx: number) => {
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
			!(text.options.hyperlink as HyperlinkPropsInternal)._rId
		) {
			const hyperlink: HyperlinkPropsInternal = text.options.hyperlink
			// Only a `url` or a `slide` needs a relationship. Two other shapes reach here and mint
			// nothing, for opposite reasons:
			//   - A navigation action button (`action` alone) is legitimately rel-free — the
			//     `ppaction://hlinkshowjump` action is self-contained, so the emitter writes `r:id=""`.
			//   - A malformed hyperlink (not an object, or no target at all) is not reportable here.
			//     The run emitter rejects the same input with `hyperlink/not-an-object` /
			//     `hyperlink/missing-target` when the deck is written; a console line ahead of that
			//     throw was one fault reported twice, the first time unroutably.
			if (typeof hyperlink === 'object' && (hyperlink.url || hyperlink.slide)) {
				registerHyperlinkRel(target, hyperlink)
			}
		} else if (
			text &&
			typeof text === 'object' &&
			text.options &&
			text.options.hyperlink &&
			(text.options.hyperlink as HyperlinkPropsInternal)._rId
		) {
			const hyperlink: HyperlinkPropsInternal = text.options.hyperlink
			const hyperlinkRelId = hyperlink._rId
			// NOTE: auto-paging will create new slides, but skip above as _rId exists, BUT this is a new slide, so add rels!
			if (hyperlinkRelId && !target._rels.some((rel) => rel.rId === hyperlinkRelId)) {
				target._rels.push(hyperlinkRel(hyperlinkRelId, hyperlink))
			}
		}
	})
}
