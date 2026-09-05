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
 *
 * That exception has been re-examined and kept, because the failure it risks is not silent after
 * all: the nine part kinds it restates are in every deck this library writes, so changing one of
 * them on the write side moves the bytes of `[Content_Types].xml` and fails the byte-identity gate
 * along with the suites that read a written deck back; changing one here fails the read tests that
 * find a layout or a master by content type. Both halves are pinned by the corpus rather than by
 * agreement between two literals, which is what makes the local spelling affordable.
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

/**
 * Prefix for the Microsoft rel types this library emits and reads. Exported for the same reason
 * as {@link OFFICE_REL}: the write side builds one-off MS rel types at their call sites.
 */
const MS_REL = 'http://schemas.microsoft.com/office/'

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
/** Any part → an external or internal jump target. External ones carry `TargetMode="External"`. */
export const HYPERLINK_REL = OFFICE_REL + 'hyperlink'
/** The presentation → `ppt/tableStyles.xml`. */
export const TABLE_STYLES_REL = OFFICE_REL + 'tableStyles'
/** The package root → `docProps/core.xml`. Package-scoped, so not under {@link OFFICE_REL}. */
export const CORE_PROPS_REL = PACKAGE_REL_NS + '/metadata/core-properties'
/** The package root → `docProps/app.xml`. */
export const EXTENDED_PROPS_REL = OFFICE_REL + 'extended-properties'
/** The package root → `docProps/custom.xml`. */
export const CUSTOM_PROPS_REL = OFFICE_REL + 'custom-properties'

// --- Media -----------------------------------------------------------------

/** A slide → an audio blob in `ppt/media/`. Paired with {@link MS_MEDIA_REL} on the same target. */
export const AUDIO_REL = OFFICE_REL + 'audio'
/** A slide → a video blob in `ppt/media/`, or an online-video placeholder. Paired with {@link MS_MEDIA_REL}. */
export const VIDEO_REL = OFFICE_REL + 'video'
/**
 * The MS-2007 `media` rel every audio/video shape carries alongside the ECMA one, pointing at the
 * same target. It is what `<p14:media r:embed>` resolves, so a media shape needs both.
 */
export const MS_MEDIA_REL = MS_REL + '2007/relationships/media'

// --- Charts and embeddings -------------------------------------------------

/** A slide → a classic (`c:`) chart part. */
export const CHART_REL = OFFICE_REL + 'chart'
/** A slide → a chartEx (`cx:`, Office 2016) chart part — an MS rel, not the ECMA `chart` one. */
export const CHARTEX_REL = MS_REL + '2014/relationships/chartEx'
/** A chart part → its mandatory style sidecar. */
export const CHART_STYLE_REL = MS_REL + '2011/relationships/chartStyle'
/** A chart part → its mandatory colour-style sidecar. */
export const CHART_COLOR_STYLE_REL = MS_REL + '2011/relationships/chartColorStyle'
/** Any part → an embedded OPC package (an Office file, itself a zip): a chart's workbook, an OLE payload. */
export const PACKAGE_REL = OFFICE_REL + 'package'

// --- Part content types ----------------------------------------------------

/** Root of the ECMA-376 content types. Private: callers want one of the constants below. */
const OD_CONTENT = 'application/vnd.openxmlformats-officedocument.'

/** `ppt/slides/slideN.xml`. */
export const SLIDE_CONTENT_TYPE = OD_CONTENT + 'presentationml.slide+xml'

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
/** `ppt/tableStyles.xml`. */
export const TABLE_STYLES_CONTENT_TYPE = OD_CONTENT + 'presentationml.tableStyles+xml'
/** `ppt/presentation.xml` in an editable `.pptx`. A `.potx` template's main part is a different one. */
export const PRESENTATION_MAIN_CONTENT_TYPE = OD_CONTENT + 'presentationml.presentation.main+xml'
/** Every `.rels` part. Declared once as a `Default` for the `rels` extension, never as an Override. */
export const RELATIONSHIPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml'
