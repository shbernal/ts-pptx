/**
 * OPC relationship type URIs used across the read model.
 *
 * These are long, near-identical strings that differ in one path segment, which makes a typo
 * both easy to write and invisible on review — a wrong URI does not throw, it just silently
 * matches nothing. Naming them once removes that class of mistake.
 *
 * Only the types used by more than one module live here. A rel type that exactly one module
 * cares about (`chart`, `hyperlink`, `tags`, the two comments parts, …) stays declared next to
 * the code that reads it — moving it here would trade a definition you can see for one you have
 * to go and find, and buy nothing.
 *
 * `src/gen/` writes four of these same URIs as inline literals. Sharing across the gen/read
 * boundary needs a home neither side owns, so that is a separate decision, not an oversight.
 */

const ECMA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** `p:sldIdLst` → a slide part. */
export const SLIDE_REL = `${ECMA}/slide`
/** A slide → its layout; a master → each of its layouts. */
export const SLIDE_LAYOUT_REL = `${ECMA}/slideLayout`
/** A layout → its master; the presentation → each master. */
export const SLIDE_MASTER_REL = `${ECMA}/slideMaster`
/** A slide → its notes slide. */
export const NOTES_SLIDE_REL = `${ECMA}/notesSlide`
/** A notes slide → the notes master; the presentation → the notes master. */
export const NOTES_MASTER_REL = `${ECMA}/notesMaster`
/** A master or notes master → its theme. Layouts and slides inherit it rather than holding one. */
export const THEME_REL = `${ECMA}/theme`
/** The package root → `ppt/presentation.xml`. The entry point to everything above. */
export const OFFICE_DOCUMENT_REL = `${ECMA}/officeDocument`
/** Any part → an image in `ppt/media/`. */
export const IMAGE_REL = `${ECMA}/image`
