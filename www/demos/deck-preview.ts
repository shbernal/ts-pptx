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
import { importDeck, renderDeck } from 'pptx-html'
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

export interface DeckPreview {
	/** A complete HTML document, for an `<iframe srcdoc>`. */
	html: string
	slideCount: number
	/** `renderDeck`'s own shortfalls — missing bytes, an unplaceable node. Normally empty. */
	warnings: string[]
	notes: FidelityRow[]
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
export async function previewDeck(bytes: Uint8Array): Promise<DeckPreview> {
	const { render, assets } = await importDeck(bytes)
	const { html, warnings } = await renderDeck(render, {
		bytes: (name: string) => assets.bytesFor({ $asset: name }),
	})

	return {
		html,
		slideCount: render.slides.length,
		warnings,
		notes: render.slides.flatMap((slide) =>
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

/** "3, 5 and 7" — the slides one grouped note applies to, in prose rather than as a list. */
export function slideList(slides: readonly number[]): string {
	const sorted = [...slides].sort((a, b) => a - b).map(String)
	if (sorted.length <= 1) return sorted[0] ?? ''
	return `${sorted.slice(0, -1).join(', ')} and ${sorted[sorted.length - 1]}`
}

/** A message for the status region. Errors arrive as `unknown` from a `catch`. */
export function failureMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}
