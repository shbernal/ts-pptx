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
import { getNewRelId, heldRelIds } from '../utils.js'
import { InvalidOptionError } from '../../errors.js'

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
		// `url` decides both fields, as it decides the emitted `<a:hlinkClick>`. `validateHyperlink`
		// refuses a link stating both, which once came out slide-typed with the URL as its target.
		data: hyperlink.url ? 'dummy' : 'slide',
		rId,
		Target: hyperlink.url ? hyperlink.url : String(hyperlink.slide),
	}
}

/**
 * Refuse a hyperlink the writer cannot express, the same way wherever one is authored.
 *
 * A hyperlink names exactly one destination: a `url`, a `slide`, or a slide-show `action`. The image
 * definer and the run emitter each checked their own subset, and the shape path checked nothing. An
 * image refused an action-only link that a shape accepted, and nothing refused `url` together with
 * `slide`: that wrote two `<a:hlinkClick>` into one `<p:cNvPr>`, which allows one, over a slide-typed
 * relationship whose target was built from the URL.
 * @param hyperlink - the caller's `hyperlink` option
 * @param call - the method or option the link was authored through, opening the message
 */
export function validateHyperlink(hyperlink: unknown, call: string): asserts hyperlink is HyperlinkPropsInternal {
	if (typeof hyperlink !== 'object' || hyperlink === null)
		throw new InvalidOptionError(
			'hyperlink/not-an-object',
			`${call}: \`hyperlink\` option should be an object. Ex: \`hyperlink:{url:'https://github.com'}\``
		)
	const { url, slide, action } = hyperlink as HyperlinkProps
	if (url && slide)
		throw new InvalidOptionError(
			'hyperlink/conflicting-targets',
			`${call}: \`hyperlink\` takes one destination but states both \`url\` and \`slide\`.`
		)
	if (!url && !slide && !action)
		throw new InvalidOptionError(
			'hyperlink/missing-target',
			`${call}: \`hyperlink\` requires either \`url\`, \`slide\`, or \`action\``
		)
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
 * @param call - the method the objects were authored through, for a refused hyperlink's message
 */
export function createHyperlinkRels(
	target: PresSlideInternal,
	text: number | string | SlideObject | TextProps | TextProps[] | TableCellInternal[] | TableCellInternal[][],
	call = 'addText'
): void {
	let textObjs: Array<HyperlinkTextObject | TableCellInternal[]> = []

	// Only text objects can have hyperlinks, bail when text param is plain text
	if (typeof text === 'string' || typeof text === 'number') return
	// IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
	else if (Array.isArray(text)) textObjs = text
	else if (typeof text === 'object') textObjs = [text]

	textObjs.forEach((text: HyperlinkTextObject | TableCellInternal[]) => {
		// A table row: walk its cells.
		if (Array.isArray(text)) {
			createHyperlinkRels(target, text, call)
			return
		}

		// Refused when it is authored rather than when it is written, and by the same rules as an
		// image's and a run's.
		if (text.options?.hyperlink) validateHyperlink(text.options.hyperlink, call)
		if (Array.isArray(text.text)) {
			// A shape carrying a hyperlink AND a run list needs BOTH walked: its own `hlinkClick` goes
			// on `p:cNvPr`, its runs' on their `a:rPr`, and they are two elements resolving two ids.
			// Recursing without registering this object's own left the shape's `r:id` pointing at a
			// relationship nothing had minted -- invisible while `addText`'s bare-string form handed
			// one options object to both the shape and its run, because the run then minted the id the
			// shape went on to read off the same object.
			//
			// A table cell's own link is registered on the cell below and not copied onto its runs:
			// the run emitter already carries it to every run. Merging the cell's options over the
			// first run's, which this walk once did, replaced that run's own colour and formatting.
			createHyperlinkRels(target, text.text, call)
		}

		const options = text.options
		const hyperlink: HyperlinkPropsInternal | undefined = options?.hyperlink
		if (!options || !hyperlink) return

		if (!hyperlink._rId) {
			// Only a `url` or a `slide` needs a relationship. A navigation action button (`action`
			// alone) is legitimately rel-free: the `ppaction://hlinkshowjump` action is self-contained,
			// so the emitter writes `r:id=""`. A malformed hyperlink never gets here; it was refused above.
			if (hyperlink.url || hyperlink.slide) registerHyperlinkRel(target, hyperlink)
			return
		}

		// A hyperlink already carrying an id was registered on some slide. Auto-paging re-registers
		// a repeated header row's hyperlink on each overflow slide under that id, and a caller can
		// hand one hyperlink object to two slides. Reusing the id is only sound when this slide has
		// it free: when it holds the id for a picture, a chart or another link, the rel would be
		// declared twice or the run would follow someone else's target. Then this slide gets its
		// own copy of the hyperlink and a fresh id, and the other slide keeps the original.
		const hyperlinkRelId = hyperlink._rId
		const wanted = hyperlinkRel(hyperlinkRelId, hyperlink)
		const registeredHere = target._rels.some(
			(rel) =>
				rel.rId === hyperlinkRelId &&
				rel.type === wanted.type &&
				rel.data === wanted.data &&
				rel.Target === wanted.Target
		)
		if (registeredHere) return
		if (!heldRelIds(target).has(hyperlinkRelId)) {
			target._rels.push(wanted)
		} else {
			const own: HyperlinkPropsInternal = { ...hyperlink }
			delete own._rId
			options.hyperlink = own
			registerHyperlinkRel(target, own)
		}
	})
}
