/**
 * OPC relationship type URIs used across the read model.
 *
 * These are long, near-identical strings that differ in one path segment, which makes a typo
 * both easy to write and invisible on review — a wrong URI does not throw, it just silently
 * matches nothing. Naming them once removes that class of mistake.
 *
 * Only the types shared by more than one module live here. Several read modules still declare
 * their own private copies of these same URIs; consolidating those is a separate sweep.
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
