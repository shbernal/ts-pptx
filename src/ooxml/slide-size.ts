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
