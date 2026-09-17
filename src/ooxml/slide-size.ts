/**
 * The slide size a presentation has when `presentation.xml` declares none.
 *
 * `p:sldSz` is optional in `CT_Presentation` (`minOccurs="0"`), and the schema says nothing about
 * what its absence means: `cx` and `cy` are required only when the element is there. So the answer
 * is PowerPoint's. A deck with `p:sldSz` removed opens at 720 × 540 pt, the 10in × 7.5in on-screen
 * 4:3 size (`ppSlideSizeOnScreen`), measured over COM against a control copy that kept its
 * 960 × 540 pt declaration. That is `LAYOUT_4x3`, not the write path's `LAYOUT_16x9` default and
 * not `LAYOUT_WIDE`.
 *
 * Anything that has to name a size for such a deck takes this one. The script converter and
 * `inspect` each had their own, and they disagreed.
 */
export const ABSENT_SLIDE_SIZE_EMU: Readonly<{ widthEmu: number; heightEmu: number }> = Object.freeze({
	widthEmu: 9144000,
	heightEmu: 6858000,
})

/**
 * The label PowerPoint shows for a slide size, as `docProps/app.xml`'s `<PresentationFormat>`
 * spells it.
 *
 * These five strings are PowerPoint's own, read back from decks it re-saved. A deck whose
 * `p:sldSz` carries `type="screen4x3"`, `"screen16x9"` or `"screen16x10"` gets the matching
 * on-screen label; `LAYOUT_WIDE`'s 13.333in × 7.5in is recognised as `Widescreen` from its
 * dimensions alone, with no `type` at all; anything else is `Custom`.
 *
 * Matched on exact EMU rather than on aspect ratio: the ratio of 10in × 7.5in and of
 * 13.333in × 10in is the same, and PowerPoint labels only the first of them 4:3.
 */
const PRESENTATION_FORMATS: ReadonlyArray<readonly [number, number, string]> = Object.freeze([
	[9144000, 6858000, 'On-screen Show (4:3)'],
	[9144000, 5143500, 'On-screen Show (16:9)'],
	[9144000, 5715000, 'On-screen Show (16:10)'],
	[12192000, 6858000, 'Widescreen'],
])

/**
 * `<PresentationFormat>` for a deck of this size, or `Custom` for one that matches no standard
 * size. `app.xml` used to say `On-screen Show (16:9)` whatever the deck.
 * @param widthEmu - the deck's slide width in EMU
 * @param heightEmu - the deck's slide height in EMU
 */
export function presentationFormatLabel(widthEmu: number, heightEmu: number): string {
	return PRESENTATION_FORMATS.find(([cx, cy]) => cx === widthEmu && cy === heightEmu)?.[2] ?? 'Custom'
}
