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

/** The `ST_SlideSizeType` values a written deck can carry; `custom` is the schema default. */
export type SlideSizeType = 'screen4x3' | 'screen16x9' | 'screen16x10'

/**
 * The standard slide sizes, each with the label PowerPoint shows for it, as `docProps/app.xml`'s
 * `<PresentationFormat>` spells it, and the `p:sldSz/@type` PowerPoint writes for it.
 *
 * The labels are PowerPoint's own, read back from decks it re-saved, and three of them follow
 * `@type` rather than the dimensions: a deck re-saved from `type="screen16x9"` comes back
 * `On-screen Show (16:9)`, while the same dimensions set through `PageSetup`, which writes no type,
 * come back `Custom`. `LAYOUT_WIDE`'s 13.333in × 7.5in is recognised as `Widescreen` from its
 * dimensions alone, and PowerPoint writes it with no `type` at all. Anything else is `Custom`.
 *
 * One table for both, so the label `app.xml` states and the type `presentation.xml` declares
 * cannot disagree about a deck.
 *
 * Matched on exact EMU rather than on aspect ratio: the ratio of 10in × 7.5in and of
 * 13.333in × 10in is the same, and PowerPoint labels only the first of them 4:3.
 */
const STANDARD_SLIDE_SIZES: ReadonlyArray<
	Readonly<{ cx: number; cy: number; label: string; type: SlideSizeType | undefined }>
> = Object.freeze([
	{ cx: 9144000, cy: 6858000, label: 'On-screen Show (4:3)', type: 'screen4x3' },
	{ cx: 9144000, cy: 5143500, label: 'On-screen Show (16:9)', type: 'screen16x9' },
	{ cx: 9144000, cy: 5715000, label: 'On-screen Show (16:10)', type: 'screen16x10' },
	{ cx: 12192000, cy: 6858000, label: 'Widescreen', type: undefined },
])

function standardSlideSize(widthEmu: number, heightEmu: number) {
	return STANDARD_SLIDE_SIZES.find(({ cx, cy }) => cx === widthEmu && cy === heightEmu)
}

/**
 * `<PresentationFormat>` for a deck of this size, or `Custom` for one that matches no standard
 * size. `app.xml` used to say `On-screen Show (16:9)` whatever the deck.
 * @param widthEmu - the deck's slide width in EMU
 * @param heightEmu - the deck's slide height in EMU
 */
export function presentationFormatLabel(widthEmu: number, heightEmu: number): string {
	return standardSlideSize(widthEmu, heightEmu)?.label ?? 'Custom'
}

/**
 * `p:sldSz/@type` for a deck of this size, or `undefined` where PowerPoint writes none: at the
 * widescreen size, and at any size that is not standard.
 *
 * Taken from the dimensions, not from the layout's name, so a `defineLayout` at exactly
 * 10in × 5.625in is typed as the on-screen 16:9 size it is, as its `app.xml` label already says.
 * @param widthEmu - the deck's slide width in EMU
 * @param heightEmu - the deck's slide height in EMU
 */
export function slideSizeType(widthEmu: number, heightEmu: number): SlideSizeType | undefined {
	return standardSlideSize(widthEmu, heightEmu)?.type
}
