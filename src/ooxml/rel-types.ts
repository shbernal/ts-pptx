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
 * (`tags`, `oleObject`, the modern comments part, …) stays declared next to the code that reads or
 * emits it — hoisting those would trade a definition you can see for one you have to go and find,
 * and buy nothing. A writer and a reader of the same part count as two modules even when each
 * names the string once: that pair is exactly the divergence this module exists to prevent.
 *
 * The same reasoning covers **part content types**, which are the same kind of fact and fail the
 * same silent way, including every entry the writer's `[Content_Types].xml` declares.
 */

import { OOXML_NS } from './namespaces.js'

/** Root of every ECMA-376 schema URI. Private: callers want one of the prefixes below. */
const SCHEMA_BASE = 'http://schemas.openxmlformats.org/'

/**
 * Prefix for `.rels` `Type` URIs — append the type name (`slide`, `image`, `theme`, …).
 * Exported for the write side, which builds one-off rel types at their call sites.
 */
export const OFFICE_REL = SCHEMA_BASE + 'officeDocument/2006/relationships/'

/** `xmlns` of every `<Relationships>` part. Also the prefix of the package-scoped rel types. */
export const PACKAGE_REL_NS = OOXML_NS.pr

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
/** A slide → its legacy comments part (`ppt/comments/commentN.xml`). */
export const COMMENTS_REL = OFFICE_REL + 'comments'
/** The presentation → the deck-wide legacy comment-author registry (`ppt/commentAuthors.xml`). */
export const COMMENT_AUTHORS_REL = OFFICE_REL + 'commentAuthors'
/** The presentation → the 2018 comment-author registry (`ppt/authors.xml`), a Microsoft rel type. */
export const MODERN_COMMENT_AUTHORS_REL = MS_REL + '2018/10/relationships/authors'
/** The presentation → `ppt/presProps.xml`. */
export const PRES_PROPS_REL = OFFICE_REL + 'presProps'
/** The presentation → `ppt/viewProps.xml`. */
export const VIEW_PROPS_REL = OFFICE_REL + 'viewProps'
/** The presentation → the handout master. */
export const HANDOUT_MASTER_REL = OFFICE_REL + 'handoutMaster'
/** A part → a theme override, which replaces the deck theme for that part alone. */
export const THEME_OVERRIDE_REL = OFFICE_REL + 'themeOverride'
/** The presentation → an embedded font part, one per face. */
export const FONT_REL = OFFICE_REL + 'font'

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
/** A slide → an embedded 3D model (`.glb`). Note `2017/06`, not the `2017` of the model3d namespace. */
export const MODEL3D_REL = MS_REL + '2017/06/relationships/model3d'

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
/** Root of the OPC package-level content types. Private, like {@link OD_CONTENT}. */
const PACKAGE_CONTENT = 'application/vnd.openxmlformats-package.'
/** Root of the Microsoft Office content types the chartEx family uses. Private. */
const MS_OFFICE_CONTENT = 'application/vnd.ms-office.'
/** Root of the PowerPoint-specific Microsoft content types the 2018 comment parts use. Private. */
const MS_POWERPOINT_CONTENT = 'application/vnd.ms-powerpoint.'

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
export const RELATIONSHIPS_CONTENT_TYPE = PACKAGE_CONTENT + 'relationships+xml'
/** `ppt/presProps.xml`. */
export const PRES_PROPS_CONTENT_TYPE = OD_CONTENT + 'presentationml.presProps+xml'
/** `ppt/viewProps.xml`. */
export const VIEW_PROPS_CONTENT_TYPE = OD_CONTENT + 'presentationml.viewProps+xml'
/** `ppt/comments/commentN.xml`, a legacy comments part. */
export const COMMENTS_CONTENT_TYPE = OD_CONTENT + 'presentationml.comments+xml'
/** `ppt/commentAuthors.xml`. */
export const COMMENT_AUTHORS_CONTENT_TYPE = OD_CONTENT + 'presentationml.commentAuthors+xml'
/** `docProps/core.xml`. Package-scoped, so not under the officedocument prefix. */
export const CORE_PROPS_CONTENT_TYPE = PACKAGE_CONTENT + 'core-properties+xml'
/** `docProps/app.xml`. */
export const EXTENDED_PROPS_CONTENT_TYPE = OD_CONTENT + 'extended-properties+xml'
/** `docProps/custom.xml`. */
export const CUSTOM_PROPS_CONTENT_TYPE = OD_CONTENT + 'custom-properties+xml'
/** `ppt/charts/chartN.xml`, a classic (`c:`) chart. */
export const CHART_CONTENT_TYPE = OD_CONTENT + 'drawingml.chart+xml'
/** `ppt/charts/chartExN.xml`. A chartEx part carries a Microsoft content type, not an `openxmlformats` one. */
export const CHARTEX_CONTENT_TYPE = MS_OFFICE_CONTENT + 'chartex+xml'
/** A chart's mandatory `style{N}.xml` sidecar. */
export const CHART_STYLE_CONTENT_TYPE = MS_OFFICE_CONTENT + 'chartstyle+xml'
/** A chart's mandatory `colors{N}.xml` sidecar. */
export const CHART_COLOR_STYLE_CONTENT_TYPE = MS_OFFICE_CONTENT + 'chartcolorstyle+xml'
/** An embedded `.xlsx` workbook: a chart's data, or an OLE payload. */
export const XLSX_CONTENT_TYPE = OD_CONTENT + 'spreadsheetml.sheet'
/** A whole `.pptx` package: the MIME type of a written deck, and of one embedded as an OLE payload. */
export const PPTX_CONTENT_TYPE = OD_CONTENT + 'presentationml.presentation'
/** The handout master part. */
export const HANDOUT_MASTER_CONTENT_TYPE = OD_CONTENT + 'presentationml.handoutMaster+xml'
/** A theme override part. Not `presentationml.`: like a theme, it is a DrawingML part. */
export const THEME_OVERRIDE_CONTENT_TYPE = OD_CONTENT + 'themeOverride+xml'
/** The 2018 comment-author registry. A PowerPoint content type, not an `openxmlformats` one. */
export const MODERN_COMMENT_AUTHORS_CONTENT_TYPE = MS_POWERPOINT_CONTENT + 'authors+xml'

// --- Parts a page points at without owning ---------------------------------

/**
 * How far a shared part's ownership reaches.
 *
 * - `deck`: chrome the whole deck owns — masters, layouts, themes, the property parts, the author
 *   registries. A page points at it and never owns it, and it is not an orphan just because the
 *   slide that last reached it is gone.
 * - `media`: a blob PowerPoint stores once and points every shape that shows it at. Page copies
 *   share it, but one that nothing references any more is an orphan.
 * - `page`: another page, or an external link. A page copy points at whichever copy of the target
 *   the import chose.
 */
export type SharedPartScope = 'deck' | 'media' | 'page'

/** One kind of part a page may point at without owning it. */
export interface SharedPartKind {
	/** The relationship type a page reaches it by. */
	readonly relType: string
	/** Its content type, or `null` where nothing matches a part of this kind by content type. */
	readonly contentType: string | null
	readonly scope: SharedPartScope
}

/**
 * Every kind of part a page points at without owning it.
 *
 * Two questions read this table, under different keys. Copying a page asks whether a relationship's
 * target may be shared by the copies rather than copied (every row, by `relType`). Pruning after a
 * slide is removed asks whether a part is deck chrome that stays even while nothing reaches it (the
 * `deck` rows, by `contentType`). They were two hand-kept lists, and the two author registries were
 * on the first and missing from the second.
 */
export const SHARED_PARTS: readonly SharedPartKind[] = [
	{ relType: SLIDE_LAYOUT_REL, contentType: SLIDE_LAYOUT_CONTENT_TYPE, scope: 'deck' },
	{ relType: SLIDE_MASTER_REL, contentType: SLIDE_MASTER_CONTENT_TYPE, scope: 'deck' },
	{ relType: NOTES_MASTER_REL, contentType: NOTES_MASTER_CONTENT_TYPE, scope: 'deck' },
	{ relType: THEME_REL, contentType: THEME_CONTENT_TYPE, scope: 'deck' },
	{ relType: THEME_OVERRIDE_REL, contentType: THEME_OVERRIDE_CONTENT_TYPE, scope: 'deck' },
	{ relType: HANDOUT_MASTER_REL, contentType: HANDOUT_MASTER_CONTENT_TYPE, scope: 'deck' },
	{ relType: PRES_PROPS_REL, contentType: PRES_PROPS_CONTENT_TYPE, scope: 'deck' },
	{ relType: VIEW_PROPS_REL, contentType: VIEW_PROPS_CONTENT_TYPE, scope: 'deck' },
	{ relType: TABLE_STYLES_REL, contentType: TABLE_STYLES_CONTENT_TYPE, scope: 'deck' },
	{ relType: OFFICE_DOCUMENT_REL, contentType: PRESENTATION_MAIN_CONTENT_TYPE, scope: 'deck' },
	{ relType: COMMENT_AUTHORS_REL, contentType: COMMENT_AUTHORS_CONTENT_TYPE, scope: 'deck' },
	{ relType: MODERN_COMMENT_AUTHORS_REL, contentType: MODERN_COMMENT_AUTHORS_CONTENT_TYPE, scope: 'deck' },
	{ relType: IMAGE_REL, contentType: null, scope: 'media' },
	{ relType: AUDIO_REL, contentType: null, scope: 'media' },
	{ relType: VIDEO_REL, contentType: null, scope: 'media' },
	{ relType: MS_MEDIA_REL, contentType: null, scope: 'media' },
	{ relType: MODEL3D_REL, contentType: null, scope: 'media' },
	{ relType: FONT_REL, contentType: null, scope: 'media' },
	{ relType: SLIDE_REL, contentType: null, scope: 'page' },
	{ relType: HYPERLINK_REL, contentType: null, scope: 'page' },
]
