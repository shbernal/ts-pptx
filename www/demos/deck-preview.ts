/**
 * The demos page's pipeline, with no DOM and no Vue in it.
 *
 * `DeckPreview.vue` is markup plus a few assignments; everything that could be wrong is
 * here, because `tsc` reads this file and does not read the SFC. The same split is what
 * lets `test/regression/www/deck-preview.test.js` cover the parts that are pure.
 *
 * The deck is written by the **workspace** library and read back by the **published** one
 * that `pptx-html` depends on. See `www/README.md` for why those are deliberately two
 * copies.
 */
import { eachNode, importDeck, renderDeck } from 'pptx-html'
import { build, compose, showcase } from 'ts-pptx-demos-showcases/quarterly-review'

/**
 * One entry from `pptx-html`'s fidelity ledger, flattened for display.
 *
 * The four fields are upstream's own vocabulary (`construct` × `disposition` × `cause`,
 * plus its prose) and are deliberately not re-labelled. A second vocabulary for the same
 * concept is how a preview starts describing losses in terms its source never agreed to.
 */
export interface FidelityRow {
	slide: number
	construct: string
	disposition: string
	cause: string
	detail: string
}

/** One slide, cut out of the rendered document so the page can show it on its own. */
export interface SlideView {
	/** 1-based, as the renderer numbers it. */
	number: number
	/** The slide's `<section>`, ready for a shadow root. See {@link splitDeck} for what changed. */
	markup: string
	/** The speaker notes the renderer printed beside the slide, one entry per paragraph. */
	notes: string[]
}

export interface DeckPreview {
	slides: SlideView[]
	/** The renderer's own stylesheet. Every slide needs it beside it. */
	styles: string
	/** Width over height, from the slides' `viewBox`. */
	aspectRatio: number
	/** `renderDeck`'s own shortfalls — missing bytes, an unplaceable node. Normally empty. */
	warnings: string[]
	fidelity: FidelityRow[]
}

/** The deck this page previews, for the page's own headings and file name. */
export const DECK = showcase

/** Assemble the showcase deck and return the package bytes. Nothing is written. */
export async function buildDeckBytes(): Promise<Uint8Array> {
	const pptx = await compose()
	return await pptx.toBytes()
}

/**
 * Build the deck straight to the visitor's downloads.
 *
 * Deliberately *not* `buildDeckBytes` plus a hand-rolled anchor: `writeFile` on the
 * browser runtime is the object-URL `<a download>` path, and routing the button through
 * it is what keeps that path exercised by something other than a human with a tab open.
 */
export async function downloadDeck(): Promise<void> {
	await build(DECK.fileName)
}

/**
 * Read a package back and render it as an HTML document.
 *
 * `renderDeck`'s `bytes` source is wired to the import's own asset index, so a deck with
 * pictures in it renders them rather than silently painting empty frames. The quarterly
 * review happens to draw every shape it shows, but a preview that only works for decks
 * without media would be a trap for the next deck added here.
 */
export async function previewDeck(
	bytes: Uint8Array,
	parse: (html: string) => Document = (html) => new DOMParser().parseFromString(html, 'text/html')
): Promise<DeckPreview> {
	const { render, assets } = await importDeck(bytes)
	const { html, warnings } = await renderDeck(render, {
		bytes: (name: string) => assets.bytesFor({ $asset: name }),
	})

	return {
		...splitDeck(parse(html), groupIds(render.slides)),
		warnings,
		fidelity: render.slides.flatMap((slide) =>
			slide.fidelity.map((note) => ({
				slide: slide.number,
				construct: note.construct,
				disposition: note.disposition,
				cause: note.cause,
				detail: note.detail,
			}))
		),
	}
}

/**
 * What a slide's shadow root adds after the renderer's stylesheet.
 *
 * The renderer styles a slide as a card in a grey scrolling document. The page draws its
 * own stage around the slide, so the card's shadow and width cap go. `all: initial` keeps
 * the site's inherited font, colour and line height out of the slide's text.
 */
export const SLIDE_FRAME_CSS = `
:host { all: initial; display: block; color-scheme: light }
.pxh-slide { display: block; width: 100%; box-shadow: none; background: #fff }
.pxh-slide > svg { display: block; width: 100%; height: auto }
`

/**
 * Cut a rendered deck document into one piece per slide.
 *
 * Each slide goes into its own shadow root, not one shared document. The renderer numbers
 * gradient ids per slide, so in one document slide 11's `url(#pxh-p0)` paints slide 1's
 * gradient. A shadow root scopes ids, so each slide gets its own.
 *
 * Four edits on the way out, and none of them touches the model:
 *
 * - **Line heights become unitless.** See {@link unitlessLineHeights}.
 * - **Groups lose their transform.** The model places a group's children in slide
 *   coordinates, and `pptx-html` 0.2 says as much in its group renderer, but it still
 *   wraps them in the group's own `translate`. Every grouped shape lands twice as far from
 *   the origin, and the showcase's KPI cards partly leave the slide. Upstream's fix is to
 *   drop that transform, and this does the same. Groups are named by id from the model,
 *   because the markup cannot tell them apart: a table's cells are node-addressed `<g>`s
 *   too, and they are placed relative to the table, so its transform has to stay. A group
 *   inherited from a layout or master carries no id and keeps its offset.
 * - **Editing is off.** The renderer marks every run `contenteditable`, because its document
 *   is meant to be edited and parsed back. Nothing parses this one, so a stray click
 *   should not put a caret in the slide.
 * - **Pictures are inlined as data URIs.** The document resolves them with a script, and
 *   `innerHTML` in a shadow root runs no scripts.
 */
export function splitDeck(
	doc: Document,
	groups: ReadonlySet<string> = new Set()
): Omit<DeckPreview, 'warnings' | 'fidelity'> {
	const styles = [...doc.querySelectorAll('style')].map((style) => style.textContent ?? '').join('\n')
	const assetUrl = assetUrls(doc)

	const slides: SlideView[] = []
	for (const section of doc.querySelectorAll('section[data-pxh-slide]')) {
		const slide = section.cloneNode(true) as Element
		for (const node of slide.querySelectorAll('[contenteditable]')) node.removeAttribute('contenteditable')
		for (const node of slide.querySelectorAll('[style*="line-height"]')) {
			node.setAttribute('style', unitlessLineHeights(node.getAttribute('style') ?? ''))
		}
		for (const node of slide.querySelectorAll('g[data-pxh-node]')) {
			if (groups.has(node.getAttribute('data-pxh-node') ?? '')) node.removeAttribute('transform')
		}
		for (const node of slide.querySelectorAll('[data-pxh-asset]')) {
			const url = assetUrl.get(node.getAttribute('data-pxh-asset') ?? '')
			if (url) node.setAttribute(node.tagName.toLowerCase() === 'image' ? 'href' : 'src', url)
		}

		const aside = section.nextElementSibling
		const notes = aside?.classList.contains('pxh-aside')
			? [...aside.querySelectorAll('p')].map((p) => p.textContent?.trim() ?? '').filter(Boolean)
			: []

		slides.push({ number: Number(section.getAttribute('data-pxh-slide')), markup: slide.outerHTML, notes })
	}

	const viewBox = doc.querySelector('section[data-pxh-slide] > svg')?.getAttribute('viewBox')?.split(/\s+/).map(Number)
	const [, , width = 16, height = 9] = viewBox ?? []
	return { slides, styles, aspectRatio: width > 0 && height > 0 ? width / height : 16 / 9 }
}

/** The id of every group on every slide, nested ones included. */
function groupIds(slides: ReadonlyArray<{ nodes: Parameters<typeof eachNode>[0] }>): Set<string> {
	const ids = new Set<string>()
	for (const slide of slides) {
		eachNode(slide.nodes, (node) => {
			if (node.kind === 'group') ids.add(node.id)
		})
	}
	return ids
}

/** The document's inline asset block, as data URIs keyed by manifest name. */
function assetUrls(doc: Document): Map<string, string> {
	const urls = new Map<string, string>()
	const block = doc.getElementById('pxh-assets')?.textContent
	if (!block) return urls
	const data = JSON.parse(block) as Record<string, string>
	const manifest = (
		JSON.parse(doc.getElementById('pxh-ir')?.textContent || '{}') as {
			assets?: Array<{ name: string; contentType?: string }>
		}
	).assets
	const types = new Map((manifest ?? []).map((entry) => [entry.name, entry.contentType]))
	for (const [name, base64] of Object.entries(data)) {
		urls.set(name, `data:${types.get(name) ?? 'application/octet-stream'};base64,${base64}`)
	}
	return urls
}

/**
 * Rewrite `line-height: 130%` as `line-height: 1.3`.
 *
 * PowerPoint's percentage line spacing is a multiple of each line's own font size, which in
 * CSS is a unitless number. The renderer writes a percentage on the paragraph instead, and
 * a percentage resolves once, against the paragraph's font size, before the runs inherit
 * it. The runs carry the real size, so a 44pt title gets 16.8px lines and its two lines
 * paint on top of each other. Upstream already writes single spacing as a unitless number;
 * the fix is to do the same for every percentage.
 */
export function unitlessLineHeights(style: string): string {
	return style.replace(
		/line-height:\s*(\d+(?:\.\d+)?)%/g,
		(_, percent: string) => `line-height:${Number(percent) / 100}`
	)
}

/** "1 difference", "4 differences". */
export function counted(count: number, noun: string): string {
	return `${count} ${noun}${count === 1 ? '' : 's'}`
}

/**
 * Group rows by their `construct × disposition` pair, keeping one detail and the slides.
 *
 * A deck of eleven slides raises the same "this shape holds an automatic field" note on
 * every slide carrying a footer. Listing it eleven times reads as eleven problems.
 */
export function summarizeNotes(notes: readonly FidelityRow[]): Array<{
	key: string
	construct: string
	disposition: string
	cause: string
	detail: string
	slides: number[]
}> {
	const groups = new Map<
		string,
		{ construct: string; disposition: string; cause: string; detail: string; slides: number[] }
	>()
	for (const note of notes) {
		const key = `${note.construct}\0${note.disposition}\0${note.cause}`
		const group = groups.get(key)
		if (group) {
			if (!group.slides.includes(note.slide)) group.slides.push(note.slide)
		} else {
			groups.set(key, {
				construct: note.construct,
				disposition: note.disposition,
				cause: note.cause,
				detail: note.detail,
				slides: [note.slide],
			})
		}
	}
	return [...groups.entries()].map(([key, group]) => ({ key, ...group }))
}

/**
 * "3, 5 and 7" — the slides one grouped note applies to, in prose rather than as a list.
 *
 * Three or more consecutive slides read as a range, "2 to 10", because the notes that
 * group are the ones that fire on every slide with a footer, and nine numbers in a row
 * say less than their two ends.
 */
export function slideList(slides: readonly number[]): string {
	const sorted = [...new Set(slides)].sort((a, b) => a - b)
	const parts: string[] = []
	let run: number[] = []
	// A run of three or more becomes its two ends; anything shorter is listed. Written as
	// an accumulator rather than index arithmetic so no element is read as possibly absent.
	const close = (): void => {
		if (run.length >= 3) parts.push(`${run[0]} to ${run[run.length - 1]}`)
		else parts.push(...run.map(String))
		run = []
	}
	for (const slide of sorted) {
		const previous = run[run.length - 1]
		if (previous !== undefined && slide !== previous + 1) close()
		run.push(slide)
	}
	close()
	if (parts.length <= 1) return parts[0] ?? ''
	return `${parts.slice(0, -1).join(', ')} and ${parts[parts.length - 1]}`
}

/** A message for the status region. Errors arrive as `unknown` from a `catch`. */
export function failureMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}
