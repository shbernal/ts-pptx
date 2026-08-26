/**
 * OPC relationship type URIs and part content types, and the schema-URI bases they are built from.
 *
 * These are long, near-identical strings that differ in one path segment, which makes a typo
 * both easy to write and invisible on review — a wrong URI does not throw, it silently matches
 * nothing on the read side and produces a part PowerPoint ignores or rejects on the write side.
 *
 * **Why this module is neither `gen/` nor `read/`.** Both halves need the same URIs, and for a
 * while each kept its own copy (`gen/oxml/schema-uris.ts`, `read/api/rel-types.ts`), with a note
 * in both saying a shared home was "a separate decision". It is this module: a rel type is a fact
 * about the OOXML package format, not about writing or reading one, so neither side should own it.
 * A divergence between the two copies would mean the writer emitting a rel the reader cannot find
 * — a round-trip bug with no compile-time signal.
 *
 * Only types used by more than one module live here. A rel type exactly one module cares about
 * (`chart`, `hyperlink`, `tags`, the two comments parts, `oleObject`, the MS media rels, …) stays
 * declared next to the code that reads or emits it — hoisting those would trade a definition you
 * can see for one you have to go and find, and buy nothing.
 *
 * The same reasoning covers **part content types**, which are the same kind of fact and fail the
 * same silent way. The write side is the exception at the bottom of this file: `[Content_Types].xml`
 * spells its entries out locally so each one stays greppable by its suffix next to the part it
 * declares.
 */

/** Root of every ECMA-376 schema URI. Private: callers want one of the prefixes below. */
const SCHEMA_BASE = 'http://schemas.openxmlformats.org/'

/**
 * Prefix for `.rels` `Type` URIs — append the type name (`slide`, `image`, `theme`, …).
 * Exported for the write side, which builds one-off rel types at their call sites.
 */
export const OFFICE_REL = SCHEMA_BASE + 'officeDocument/2006/relationships/'

/** `xmlns` of every `<Relationships>` part. Also the prefix of the package-scoped rel types. */
export const PACKAGE_REL_NS = SCHEMA_BASE + 'package/2006/relationships'

/** `p:sldIdLst` → a slide part. */
export const SLIDE_REL = OFFICE_REL + 'slide'
/** A slide → its layout; a master → each of its layouts. */
export const SLIDE_LAYOUT_REL = OFFICE_REL + 'slideLayout'
/** A layout → its master; the presentation → each master. */
export const SLIDE_MASTER_REL = OFFICE_REL + 'slideMaster'
/** A slide → its notes slide. */
export const NOTES_SLIDE_REL = OFFICE_REL + 'notesSlide'
/** A notes slide → the notes master; the presentation → the notes master. */
export const NOTES_MASTER_REL = OFFICE_REL + 'notesMaster'
/** A master or notes master → its theme. Layouts and slides inherit it rather than holding one. */
export const THEME_REL = OFFICE_REL + 'theme'
/** The package root → `ppt/presentation.xml`. The entry point to everything above. */
export const OFFICE_DOCUMENT_REL = OFFICE_REL + 'officeDocument'
/** Any part → an image in `ppt/media/`. */
export const IMAGE_REL = OFFICE_REL + 'image'

// --- Part content types ----------------------------------------------------

/** Root of the ECMA-376 content types. Private: callers want one of the constants below. */
const OD_CONTENT = 'application/vnd.openxmlformats-officedocument.'

/** `ppt/slideMasters/slideMasterN.xml`. */
export const SLIDE_MASTER_CONTENT_TYPE = OD_CONTENT + 'presentationml.slideMaster+xml'
/** `ppt/slideLayouts/slideLayoutN.xml`. */
export const SLIDE_LAYOUT_CONTENT_TYPE = OD_CONTENT + 'presentationml.slideLayout+xml'
/** `ppt/notesSlides/notesSlideN.xml`. */
export const NOTES_SLIDE_CONTENT_TYPE = OD_CONTENT + 'presentationml.notesSlide+xml'
/** `ppt/notesMasters/notesMasterN.xml`. */
export const NOTES_MASTER_CONTENT_TYPE = OD_CONTENT + 'presentationml.notesMaster+xml'
/** `ppt/theme/themeN.xml`. Not `presentationml.`: a theme is a DrawingML part. */
export const THEME_CONTENT_TYPE = OD_CONTENT + 'theme+xml'
