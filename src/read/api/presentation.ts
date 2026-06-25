/**
 * Read-model entry point: `Presentation` wraps an `OpcPackage` and exposes a
 * navigable, typed view of the deck (slides → shapes → text), backed by the
 * live DOM so the same nodes can later be mutated.
 */
import { emuToInches } from '../../units.js'
import { OpcPackage, type OpcInput } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { relativePartName, relsPartNameFor } from '../opc/partnames.js'
import { ELEMENT_NODE, OOXML_NS, attr, createElement, firstChild, getElements, getOrAddChild, insertInOrder, intValue, removeChildrenByQName, setAttr, type Element } from '../oxml/dom.js'
import { EMBEDDED_FONT_SLOTS, FONT_DATA_CONTENT_TYPE, FONT_DATA_EXTENSION, FONT_REL_TYPE, type EmbeddedFontSlot } from '../../embedded-fonts.js'
import { flattenShape, flattenSlide, remapLiteralColors, restyleSlide, type FlattenContext } from '../oxml/theme.js'
import { resolveSlideThemeParts } from './theme-context.js'
import { Slide } from './slide.js'
import { wrapShapeElement, type Shape } from './shapes.js'

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const NOTES_SLIDE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
// Microsoft 2007 `media` rel: paired with the ECMA audio/video rel (same Target),
// referenced by the slide body's <p14:media r:embed>.
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'
const TABLE_STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
const TABLE_STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'
// The well-known "no style / table grid" GUID PowerPoint uses as a tblStyleLst @def.
const TABLE_STYLES_DEFAULT_GUID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

const SLIDE_MASTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
const SLIDE_LAYOUT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/** Content type of the main part in an editable `.pptx` package. */
const PRESENTATION_MAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
/** Content type of the main part in a `.potx` template package — flipped to {@link PRESENTATION_MAIN_CONTENT_TYPE} by {@link Presentation.fromTemplate}. */
const PRESENTATION_TEMPLATE_MAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'

const textEncoder = new TextEncoder()

/**
 * Content types that are shared deck chrome: reachable through the
 * presentation → master → layout → theme graph, not owned by any one slide.
 * {@link Presentation.removeSlide} never prunes these as a removed slide's
 * orphan, even while momentarily unreferenced.
 */
const SHARED_CHROME_CONTENT_TYPES = new Set([
	SLIDE_MASTER_CONTENT_TYPE,
	SLIDE_LAYOUT_CONTENT_TYPE,
	'application/vnd.openxmlformats-officedocument.theme+xml',
	'application/vnd.openxmlformats-officedocument.themeOverride+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
	'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
	PRESENTATION_MAIN_CONTENT_TYPE,
])

/**
 * `p:sldIdLst`'s document-order successors in `CT_Presentation` (ECMA-376):
 * everything that may legally follow it, so an inserted `p:sldIdLst` lands in the
 * right position when a template omitted it (zero slides).
 */
const PRESENTATION_SLD_ID_LST_SUCCESSORS = [
	'p:sldSz',
	'p:notesSz',
	'p:smartTags',
	'p:embeddedFontLst',
	'p:custShowLst',
	'p:photoAlbum',
	'p:custDataLst',
	'p:kinsoku',
	'p:defaultTextStyle',
	'p:modifyVerifier',
	'p:extLst',
]

/**
 * `p:embeddedFontLst`'s document-order successors in `CT_Presentation` (index 7,
 * after `smartTags`): everything that may legally follow it, so a created list
 * lands in the right slot when the deck has none yet.
 */
const PRESENTATION_EMBEDDED_FONT_LST_SUCCESSORS = ['p:custShowLst', 'p:photoAlbum', 'p:custDataLst', 'p:kinsoku', 'p:defaultTextStyle', 'p:modifyVerifier', 'p:extLst']

/**
 * A face slot's document-order successors in `CT_EmbeddedFontListEntry`
 * (`font`, `regular`, `bold`, `italic`, `boldItalic`), so a newly-inserted face
 * keeps the schema's child order regardless of which slots already exist.
 */
const EMBEDDED_FONT_FACE_SUCCESSORS: Record<EmbeddedFontSlot, string[]> = {
	regular: ['p:bold', 'p:italic', 'p:boldItalic'],
	bold: ['p:italic', 'p:boldItalic'],
	italic: ['p:boldItalic'],
	boldItalic: [],
}

/** ST_SlideId minimum (ECMA-376): slide ids live in [256, 2147483647]. */
const MIN_SLIDE_ID = 256

/** ST_SlideLayoutId minimum (ECMA-376): slide-layout ids start at 2147483648. */
const MIN_SLIDE_LAYOUT_ID = 2147483648

/** ST_SlideMasterId minimum (ECMA-376): slide-master ids start at 2147483648. */
const MIN_SLIDE_MASTER_ID = 2147483648

/** Slide dimensions, in both EMU (the OOXML unit) and inches. */
export interface SlideSize {
	widthEmu: number
	heightEmu: number
	widthIn: number
	heightIn: number
}

/** Options for {@link Presentation.importSlide}. */
export interface ImportSlideOptions {
	/**
	 * How the imported slide relates to themes.
	 *
	 * - `'copy'` (default): bring the slide's own `slideLayout → slideMaster →
	 *   theme` subgraph across, so the slide renders byte-for-byte as authored.
	 *   A deck stitched from N sources then carries N themes/masters.
	 * - `'preserve'`: *flatten then attach* — bake the source theme's colours and
	 *   style-matrix fills into the slide XML (so its pixels do not change), then
	 *   bind it to this deck's existing master/layout instead of importing the
	 *   source theme. The result is a single-theme file whose imported slides are
	 *   theme-independent: it fixes renderers that mis-resolve per-slide scheme
	 *   colours against the wrong (first) theme, and tidies the deck for handoff.
	 *
	 *   To stay faithful across the rebind, `preserve` also bakes the inheritance
	 *   the rebind would otherwise break explicitly onto the slide: the slide's
	 *   effective background; each placeholder's *inherited* geometry (`a:xfrm`
	 *   from the matching source layout/master placeholder) so it cannot shift or
	 *   clip; and each placeholder run's *inherited* colour and size/weight
	 *   (`sz`/`b`/`i`) from the source layout/master text styles. Typeface
	 *   (`a:latin`) is left to re-bind to the destination theme, like `fontRef`.
	 *   By default it does **not** carry decorative graphics that live on the source
	 *   master/layout shape tree (logos, accent shapes): those belong to the
	 *   master `preserve` deliberately drops. Set {@link carryMasterGraphics} to
	 *   bake them onto the slide instead.
	 * - `'restyle'`: re-brand to *this* deck. Rebind the slide to this deck's
	 *   master/layout exactly like `preserve` but **skip the flatten** — leave every
	 *   `a:schemeClr`, style-matrix ref (`fillRef`/`lnRef`/`effectRef`/`fontRef`),
	 *   and `p:bg` `bgRef` symbolic so they re-resolve against the *destination*
	 *   theme. The slide keeps its geometry, text, and structure but adopts this
	 *   deck's colours/fonts: `preserve` makes it "look the same everywhere",
	 *   `restyle` makes it "look like mine". The slide's own `p:clrMapOvr` is dropped
	 *   so the destination master's `clrMap` governs the re-brand.
	 *
	 *   **Load-bearing limitation:** `restyle` can only recolour what is *symbolic*.
	 *   Anything the source authored as a literal `a:srgbClr` has no theme reference
	 *   to re-resolve and stays exactly that colour, so a slide with a baked literal
	 *   palette re-brands little or nothing. Use `restyle` for slides built against
	 *   theme colours/style matrices, not hardcoded RGB — or set {@link remapLiterals}
	 *   to force-remap literals and copy table styles. Re-brand is inherently a
	 *   visual change (a source `accent1` light-on-dark can invert against a dark
	 *   destination `accent1`), so its output needs visual QA. A restyled table
	 *   resolves its `@tableStyleId` against the *destination* `tableStyles`; if the
	 *   destination lacks that id the table falls back (which {@link remapLiterals}
	 *   also addresses by copying the source style across).
	 */
	theme?: 'copy' | 'preserve' | 'restyle'

	/**
	 * `restyle` mode only. Push the re-brand past what symbolic references reach, for
	 * slides whose palette is partly hardcoded. When `true` it does two things the
	 * plain re-brand cannot:
	 *
	 * - **Literal colours** — every literal `a:srgbClr` equal to a *source*-theme
	 *   `clrScheme` slot is rewritten back to a symbolic `a:schemeClr` (routed through
	 *   the source `clrMap`), so it re-resolves against this deck's theme instead of
	 *   staying its authored RGB. A literal matching no source slot is left untouched.
	 * - **Table styles** — any `@tableStyleId` the slide references is copied from the
	 *   source `tableStyles.xml` into this deck's (same id, idempotent, leaving an id
	 *   the destination already defines alone), so a restyled table keeps its style
	 *   instead of falling back. The copied definition is itself symbolic, so it
	 *   re-brands to this deck's theme.
	 *
	 * Off by default: it deliberately reinterprets authored literals as theme colours,
	 * which is a visual change that needs QA. Ignored unless `theme` is `'restyle'`.
	 */
	remapLiterals?: boolean

	/**
	 * `preserve`/`restyle` modes only. When `true`, bake the source
	 * `slideLayout`/`slideMaster` shape-tree decorations (everything on those shape
	 * trees *except* placeholders — logos, accent curves, footers drawn as shapes)
	 * onto the imported slide, behind its own content, so master/layout branding
	 * survives the rebind to this deck's master. Their media are copied across.
	 *
	 * Under `preserve` the carried decorations' theme references are flattened like
	 * the slide's own content; under `restyle` they are left symbolic and so
	 * re-brand to the destination palette along with the slide — note a carried
	 * source logo could recolour unexpectedly under a different palette.
	 *
	 * Off by default: it raises fidelity for cover/closer/divider slides at the
	 * cost of duplicating shapes that would otherwise live once on the shared
	 * master, so opt in only when that branding actually needs to travel with the
	 * slide. Ignored unless `theme` is `'preserve'` or `'restyle'`.
	 */
	carryMasterGraphics?: boolean

	/**
	 * Zero-based insert position in `p:sldIdLst` (deck order). `0` makes the
	 * imported slide first; an `at` past the current slide count — or omitting it —
	 * appends. Use it to place brand bookends (cover at `0`, closer appended)
	 * around generator-authored interior slides regardless of import order.
	 */
	at?: number

	/**
	 * Carry the *source deck's* embedded fonts (`p:embeddedFontLst` in its
	 * `presentation.xml`) into this deck, so an imported slide that renders with an
	 * embedded face keeps it on machines that lack the font. Off by default —
	 * fonts live on the presentation, not the slide, so they are only worth copying
	 * when you want the embed to travel.
	 *
	 * The font binary parts are copied under fresh `/ppt/fonts/` names (deduped via
	 * the per-source copy registry, so repeated imports copy each face once), the
	 * `application/x-fontdata` Default is added, and entries are merged into this
	 * deck's `p:embeddedFontLst` — de-duplicated by `typeface` + face slot, so a
	 * face this deck already embeds is reused rather than duplicated. The carry is a
	 * whole-deck operation (it copies *all* the source's embedded fonts, not just the
	 * faces this one slide uses), since the source list does not record which slide
	 * uses which face.
	 * @default false
	 */
	embedFonts?: boolean

	/**
	 * What to do when the source slide size differs from this deck's. By default
	 * (`false`) `importSlide` throws on any mismatch. Set it to rescale the imported
	 * geometry onto this deck's canvas instead:
	 *
	 * - `'fit'` (or `true`): uniform scale by `min(sx, sy)` then center, so aspect
	 *   ratio is preserved (circles stay circles, rotations hold) and the slack on
	 *   the longer axis becomes a centered margin. Matches PowerPoint's "Ensure Fit".
	 * - `'stretch'`: independent per-axis scale so the content fills the canvas;
	 *   distorts shapes and cannot faithfully reposition rotated shapes. Matches
	 *   "Maximize".
	 *
	 * Only *geometry* is rescaled — every top-level shape/group/graphicFrame
	 * transform (`a:off`/`a:ext`) and table grid (`a:gridCol@w`, `a:tr@h`). Font
	 * sizes, line widths, and other absolute sizes are left as authored, so heavy
	 * down-scaling can leave text overflowing its (now smaller) box. In `copy` mode
	 * the imported layout and master shape trees are rescaled too, so inherited
	 * placeholder and background geometry stays aligned; in `preserve`/`restyle` the
	 * slide rebinds to this deck's master/layout (already the right size), so only
	 * the slide is touched.
	 * @default false
	 */
	rescale?: boolean | 'fit' | 'stretch'
}

/** Options for {@link Presentation.importShape} / {@link Presentation.importShapes}. */
export interface ImportShapeOptions {
	/**
	 * How the lifted shape relates to themes, mirroring {@link ImportSlideOptions}
	 * but scoped to one shape subtree:
	 *
	 * - `'preserve'` (default): bake the shape's scheme/style-matrix colours (and,
	 *   for a lifted placeholder, its inherited geometry/colour/size) to literals
	 *   using the *source* slide's theme, so it keeps its look on a host slide whose
	 *   theme differs. The safe default for composing across decks.
	 * - `'restyle'`: leave the shape's theme references symbolic so it re-brands to
	 *   the host theme. Only *symbolic* colours re-brand — literal `a:srgbClr` the
	 *   source baked in stays put (same limitation as `importSlide` restyle).
	 * - `'copy'`: bring the shape's XML across untouched; only sane when the host
	 *   already shares the source theme.
	 *
	 * Unlike a slide import this never runs the slide-scoped background passes — a
	 * background belongs to a slide, not to a composed shape.
	 */
	theme?: 'preserve' | 'restyle' | 'copy'
	/** Override left edge (EMU). Omitted axes keep the shape's source `a:off`/`a:ext`. */
	left?: number
	/** Override top edge (EMU). */
	top?: number
	/** Override width (EMU); must be positive. */
	width?: number
	/** Override height (EMU); must be positive. */
	height?: number
	/**
	 * Insert position among the host shape tree's shape children (z-order, 0 =
	 * backmost). Out-of-range or omitted appends on top. A batch inserts in the
	 * given order starting at this position.
	 */
	at?: number
}

/** A master brought across by {@link Presentation.importSlideMasters}. */
export interface ImportedSlideMaster {
	/** Partname of the copied master in this (destination) package. */
	partName: string
	/** Partnames of the layouts copied under it, in source `p:sldLayoutIdLst` order. */
	layoutPartNames: string[]
}

/** Options for {@link Presentation.importSlideMasters}. */
export interface ImportSlideMastersOptions {
	/**
	 * Pick which of the source's masters to graft. Receives the master's `p:cSld`
	 * `name` (`''` when unnamed) and its zero-based index in the source's
	 * `p:sldMasterIdLst`. Default: every master.
	 */
	masters?: (name: string, index: number) => boolean
	/**
	 * Pick which layouts under each grafted master to bring. Receives the layout's
	 * `p:cSld` `name` and its zero-based index within that master's
	 * `p:sldLayoutIdLst`. Default: the whole family.
	 */
	layouts?: (name: string, index: number) => boolean
	/**
	 * Require the source and destination slide sizes to match (default `true`).
	 * A grafted master is shipped into the layout gallery, not applied to existing
	 * slides — but a layout authored at a different canvas size shows up mis-scaled
	 * in that gallery, so the guard is on by default. Pass `false` to graft anyway.
	 */
	requireEqualSize?: boolean
}

/** A layout in this deck's gallery, addressable as an {@link AppendSlidesOptions} target. */
export interface LayoutHandle {
	/** Partname of the layout in this package (e.g. `/ppt/slideLayouts/slideLayout2.xml`). */
	partName: string
	/** The layout's `p:cSld@name` (`''` when unnamed). */
	name: string
	/** Partname of the master this layout belongs to. */
	masterPartName: string
	/** Zero-based index of the master in `p:sldMasterIdLst`. */
	masterIndex: number
	/** Zero-based index of the layout within its master's `p:sldLayoutIdLst`. */
	layoutIndex: number
}

/**
 * One authored slide, extracted from a generator for injection into an existing
 * package. The slide body XML references its media/hyperlinks by relationship id
 * only, so {@link Presentation.appendSlides} preserves each `rId` and only
 * repoints its target — see {@link SlideSource}.
 */
export interface ExtractedSlide {
	/** Standalone `<p:sld>` part body (XML declaration + namespaces included). */
	xml: string
	/** Image media the body references, keyed by the `rId` used in {@link xml}. */
	media: Array<{ rId: number; bytes: Uint8Array; extn: string; contentType: string }>
	/** External hyperlink rels, keyed by the `rId` used in {@link xml}. */
	hyperlinks: Array<{ rId: number; target: string }>
	/**
	 * Charts the body references, keyed by the `rId` used in {@link xml}. Each carries
	 * the chart part XML and its embedded workbook bytes; the chart part's own `.rels`
	 * (workbook reference) is rebuilt by {@link Presentation.appendSlides}.
	 */
	charts: Array<{ rId: number; chartXml: string; embeddingBytes: Uint8Array }>
	/** Internal slide-to-slide links: the `rId` used in {@link xml} → 1-based source slide number. */
	slideLinks: Array<{ rId: number; sourceSlideNumber: number }>
	/**
	 * Embedded audio/video the body references. Each item is one media binary backed
	 * by two rels sharing a Target — the ECMA `audio`/`video` rel (`mediaRid`) and the
	 * MS-2007 `media` rel (`msMediaRid`) — plus a separate preview image rel
	 * (`previewRid`). {@link Presentation.appendSlides} reproduces this rel graph.
	 */
	avMedia: AvMediaItem[]
	/**
	 * Online (external-link) video the body references. Each item is two *external*
	 * rels sharing one link Target — the ECMA `video` rel (`mediaRid`) and the MS-2007
	 * `media` rel (`msMediaRid`) — with **no** media binary part. The poster frame is a
	 * normal image rel carried by {@link media}. {@link Presentation.appendSlides}
	 * wires the two external rels without reserving a media part or content type.
	 */
	onlineMedia: OnlineMediaItem[]
}

/** One embedded audio/video item extracted for {@link Presentation.appendSlides}. */
export interface AvMediaItem {
	/** Whether the item is `audio` (`<a:audioFile>`) or `video` (`<a:videoFile>`). */
	mtype: 'audio' | 'video'
	/** Body `rId` of the ECMA `audio`/`video` rel (`r:link`); points at the media part. */
	mediaRid: number
	/** Body `rId` of the MS-2007 `media` rel (`p14:media r:embed`); shares the media part Target. */
	msMediaRid: number
	/** Body `rId` of the preview image rel (`a:blip r:embed` in the blipFill). */
	previewRid: number
	/** The audio/video binary the media part will hold. */
	mediaBytes: Uint8Array
	/** Media file extension (no dot), e.g. `mp4`, `mp3`. */
	mediaExtn: string
	/** OPC content type for the media part (PowerPoint-authored, e.g. `audio/mpeg`). */
	mediaContentType: string
	/** The preview/poster image bytes. */
	previewBytes: Uint8Array
	/** Preview image extension (no dot), e.g. `png`. */
	previewExtn: string
	/** OPC content type for the preview image part. */
	previewContentType: string
}

/** One online (external-link) video item extracted for {@link Presentation.appendSlides}. */
export interface OnlineMediaItem {
	/** Body `rId` of the ECMA `video` rel (`a:videoFile r:link`); External, no part. */
	mediaRid: number
	/** Body `rId` of the MS-2007 `media` rel (`p14:media r:link`); External, shares the link Target. */
	msMediaRid: number
	/** The external video link both rels point at (`TargetMode="External"`). */
	link: string
}

/** A generator's authored slides + canvas size, the input to {@link Presentation.appendSlides}. */
export interface ExtractedSlides {
	widthEmu: number
	heightEmu: number
	slides: ExtractedSlide[]
}

/**
 * Structural view of a slide producer (a `PptxGenJS` instance satisfies this).
 * Kept structural so the read subsystem never imports the generator at runtime.
 */
export interface SlideSource {
	extractSlides(opts?: { onMediaError?: 'throw' | 'placeholder' }): Promise<ExtractedSlides>
}

/** Options for {@link Presentation.appendSlides}. */
export interface AppendSlidesOptions {
	/** Target layout to bind every appended slide to: by `p:cSld@name` or a {@link LayoutHandle}. */
	layout: string | LayoutHandle
	/**
	 * Zero-based `p:sldIdLst` position for the first appended slide; subsequent
	 * slides follow it in order. Omitted/out-of-range appends at the end.
	 */
	at?: number
	/** How `addImage` media errors surface during extraction (default `'throw'`). */
	onMediaError?: 'throw' | 'placeholder'
}

/** Options for {@link Presentation.fromTemplate}. */
export interface FromTemplateOptions {
	/**
	 * Keep a `.potx` main part's `…template.main+xml` content type instead of
	 * normalizing it to the editable `…presentation.main+xml` type. Off by default
	 * (the saved package opens as an editable deck, not a template). No effect on a
	 * `.pptx` input, whose main part is already editable.
	 */
	keepTemplateContentType?: boolean
}

export class Presentation {
	#presentationPart: Part | undefined
	/**
	 * Per-source copy registry for {@link importSlide}: source `OpcPackage` →
	 * (source partname → partname allocated in this package). Lets parts shared
	 * across imports from the same source deck (layout, master, theme, media) be
	 * copied once and reused on later calls.
	 */
	#importRegistry = new Map<OpcPackage, Map<string, string>>()
	/**
	 * Parts whose geometry {@link importSlide}'s `rescale` has already rewritten, so a
	 * layout/master shared across repeated imports from one source is not scaled twice.
	 */
	#rescaledParts = new Set<string>()

	private constructor(readonly opc: OpcPackage) {}

	/** Open a `.pptx` from bytes and wrap it as a navigable `Presentation`. */
	static async load(input: OpcInput): Promise<Presentation> {
		return new Presentation(await OpcPackage.load(input))
	}

	/** Wrap an already-loaded OPC package (e.g. from the lower-level API). */
	static fromPackage(opc: OpcPackage): Presentation {
		return new Presentation(opc)
	}

	/**
	 * Open a PowerPoint template (`.pptx` or `.potx`) and return it as an empty
	 * deck shell ready to author onto: its slide masters, layouts, and theme are
	 * kept **byte-identical**, while any sample slides the template carried are
	 * stripped so only the shared chrome remains.
	 *
	 * Use it to build a fresh deck on a corporate template without rebuilding the
	 * masters in code: discover the bindable layouts with {@link layouts}, then
	 * author slides with a generator sized to match the template and graft them in
	 * with {@link appendSlides} (which enforces an equal slide size). Saving yields
	 * an editable `.pptx` that reuses the template's authored chrome verbatim.
	 *
	 * ```ts
	 * const deck = await Presentation.fromTemplate(templateBytes) // .pptx or .potx
	 * deck.layouts().map(l => l.name)                             // discover layouts
	 * await deck.appendSlides(pptx, { layout: 'Title and Content' })
	 * const out = await deck.save()                               // editable .pptx
	 * ```
	 *
	 * A `.potx` package declares its main part with the template content type; by
	 * default that override is flipped to the editable presentation type so the
	 * saved output opens as a normal deck. Pass `keepTemplateContentType: true` to
	 * preserve the template type. A `.pptx` input needs no flip, and a template
	 * that already carries zero slides makes the strip a no-op.
	 */
	static async fromTemplate(input: OpcInput, options: FromTemplateOptions = {}): Promise<Presentation> {
		const pres = new Presentation(await OpcPackage.load(input))

		// Normalize a .potx main part to the editable presentation content type so
		// the saved package opens as a deck, not a template. The officeDocument
		// relationship resolves the main part regardless of its content type, so a
		// .potx already loads; only the [Content_Types].xml override needs flipping.
		if (!options.keepTemplateContentType) {
			const mainPart = pres.presentationPart
			if (mainPart.contentType === PRESENTATION_TEMPLATE_MAIN_CONTENT_TYPE) {
				pres.opc.contentTypes.ensureRegistered(mainPart.partName, PRESENTATION_MAIN_CONTENT_TYPE)
			}
		}

		// Strip any sample slides to a master/layout-only shell. removeSlide never
		// prunes shared chrome, so masters/layouts/theme stay byte-identical.
		while (pres.slides.length > 0) pres.removeSlide(0)

		return pres
	}

	/** The main presentation part (`/ppt/presentation.xml`), resolved via the package `officeDocument` relationship. */
	get presentationPart(): Part {
		if (this.#presentationPart) return this.#presentationPart
		const packageRels = this.opc.relationshipsFor('/')
		const officeDocument = packageRels.byType(OFFICE_DOCUMENT_REL)
		if (officeDocument.length !== 1) {
			throw new Error(`Expected exactly one officeDocument relationship, found ${officeDocument.length}`)
		}
		const partName = packageRels.resolveTarget(officeDocument[0].id)
		const part = this.opc.part(partName)
		if (!part) throw new Error(`officeDocument relationship targets a missing part: ${partName}`)
		this.#presentationPart = part
		return part
	}

	/** The slides in presentation order (resolved from `p:sldIdLst` + the presentation's relationships). */
	get slides(): Slide[] {
		const root = this.presentationPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (!sldIdLst) return []
		const rels = this.opc.relationshipsFor(this.presentationPart.partName)
		const slides: Slide[] = []
		let index = 0
		for (const sldId of getElements(sldIdLst, 'p:sldId')) {
			const relId = attr(sldId, 'r:id')
			if (!relId) continue
			const partName = rels.resolveTarget(relId)
			const part = this.opc.part(partName)
			if (!part) throw new Error(`Slide relationship ${relId} targets a missing part: ${partName}`)
			slides.push(new Slide(this, part, intValue(attr(sldId, 'id')) ?? 0, index++))
		}
		return slides
	}

	/** Slide dimensions (`p:sldSz`), or `null` if the presentation declares none. */
	get slideSize(): SlideSize | null {
		const root = this.presentationPart.dom.documentElement
		const sldSz = root && firstChild(root, 'p:sldSz')
		if (!sldSz) return null
		const widthEmu = intValue(attr(sldSz, 'cx'))
		const heightEmu = intValue(attr(sldSz, 'cy'))
		if (widthEmu === null || heightEmu === null) return null
		return { widthEmu, heightEmu, widthIn: emuToInches(widthEmu), heightIn: emuToInches(heightEmu) }
	}

	/**
	 * Duplicate the slide at `index` and insert the copy at `options.at` (deck
	 * order; `0` = first), defaulting to appending at the end when `at` is omitted
	 * or out of range. Returns the new slide. The new slide part copies the source
	 * bytes verbatim and shares the source's relationship targets (layout, images,
	 * …) by copying its `.rels`; a new presentation→slide relationship and a
	 * `p:sldId` entry are wired up. Marks the presentation part dirty.
	 *
	 * Note: relationships are copied as-is, so a source slide that owns a
	 * one-to-one part (e.g. a notes slide) would end up shared with the clone.
	 */
	cloneSlide(index: number, options: { at?: number } = {}): Slide {
		const source = this.slides[index]
		if (!source) throw new Error(`No slide at index ${index} to clone`)
		const opc = this.opc
		const sourcePart = source.part

		// 1. Copy the slide part bytes verbatim into a fresh slide partname.
		const newPartName = opc.reservePartNameLike(sourcePart.partName)
		const newPart = opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)

		// 2. Copy the slide's relationships (targets resolve identically — same dir).
		const sourceRels = opc.part(relsPartNameFor(sourcePart.partName))
		if (sourceRels) opc.addPart(relsPartNameFor(newPartName), sourceRels.contentType, sourceRels.bytes)

		// 3. Wire the new slide into the presentation (rel + p:sldId entry) at `at`.
		return this.#insertSlidePart(newPart, options.at)
	}

	/**
	 * Remove the slide at `index` (deck order) and return its former partname. The
	 * `p:sldId` entry and the presentation→slide relationship are dropped, the slide
	 * part and its `.rels` are deleted, and any part the slide *privately* owned
	 * (its notes slide, slide-only media, charts/embeddings) that no remaining part
	 * references is pruned too — recursively. Shared deck chrome (layout, master,
	 * theme, …) is never pruned, so the deck stays renderable; removing every slide
	 * leaves a valid master/layout-only package (a template shell).
	 *
	 * Untouched parts stay byte-identical, matching the package fidelity contract.
	 * Throws when there is no slide at `index`.
	 */
	removeSlide(index: number): string {
		const slide = this.slides[index]
		if (!slide) throw new Error(`No slide at index ${index} to remove`)
		const partName = slide.partName

		// The slide's internal targets, captured before its rels are dropped, so the
		// parts it privately owned can be pruned afterwards.
		const slideRels = this.opc.relationshipsFor(partName)
		const formerTargets = [...slideRels].filter((rel) => rel.targetMode !== 'External').map((rel) => slideRels.resolveTarget(rel.id))

		// Unwire from presentation.xml: remove the matching p:sldId and the rel.
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const root = presPart.dom.documentElement
		const sldIdLst = root && firstChild(root, 'p:sldIdLst')
		if (sldIdLst) {
			for (const sldId of getElements(sldIdLst, 'p:sldId')) {
				const relId = attr(sldId, 'r:id')
				if (relId && presRels.get(relId) && presRels.resolveTarget(relId) === partName) {
					sldIdLst.removeChild(sldId)
					presRels.remove(relId)
					break
				}
			}
		}
		presPart.markDirty()

		// Drop the slide part and its .rels, then prune the parts it privately owned.
		this.opc.removePart(relsPartNameFor(partName))
		this.opc.removePart(partName)
		for (const target of formerTargets) this.#pruneIfOrphan(target)

		return partName
	}

	/**
	 * Remove `partName` if it is neither shared chrome nor still referenced by any
	 * remaining part, then recurse into the parts it referenced. The pruning a
	 * removed slide triggers (notes/media/charts the slide alone used).
	 */
	#pruneIfOrphan(partName: string): void {
		const part = this.opc.part(partName)
		if (!part || SHARED_CHROME_CONTENT_TYPES.has(part.contentType)) return
		if (this.#isReferenced(partName)) return
		const rels = this.opc.relationshipsFor(partName)
		const childTargets = [...rels].filter((rel) => rel.targetMode !== 'External').map((rel) => rels.resolveTarget(rel.id))
		this.opc.removePart(relsPartNameFor(partName))
		this.opc.removePart(partName)
		for (const child of childTargets) this.#pruneIfOrphan(child)
	}

	/** Whether any remaining part (or the package root) resolves an internal relationship to `partName`. */
	#isReferenced(partName: string): boolean {
		for (const owner of [...this.opc.parts.keys(), '/']) {
			if (owner.endsWith('.rels')) continue
			const rels = this.opc.relationshipsFor(owner)
			for (const rel of rels) {
				if (rel.targetMode === 'External') continue
				if (rels.resolveTarget(rel.id) === partName) return true
			}
		}
		return false
	}

	/**
	 * Append a copy of `source.slides[index]` to this presentation and return it.
	 *
	 * Unlike {@link cloneSlide} (same-deck duplicate), this copies a slide across
	 * a package boundary: it brings the connected sub-graph the slide depends on —
	 * its `slideLayout` → `slideMaster` → `theme`, plus any media, charts, and
	 * embeddings — into this package under fresh partnames, rewriting every
	 * partname, relationship id, and content-type registration so the result is a
	 * self-consistent OPC package. Parts of this (target) package that are not
	 * touched stay byte-identical, matching `cloneSlide`'s fidelity contract.
	 *
	 * Only the layout(s) actually used by imported slides are copied; the imported
	 * master's `p:sldLayoutIdLst` is pruned to exactly those, mirroring how
	 * PowerPoint's "Reuse Slides" brings a slide across. Parts shared by repeated
	 * imports from the same source deck are copied once and reused.
	 *
	 * With `{ theme: 'preserve' }` the slide's source theme is instead *flattened*
	 * into the slide XML and the slide is bound to this deck's existing
	 * master/layout; with `{ theme: 'restyle' }` the slide is bound to this deck's
	 * master/layout with its theme references left symbolic, so it re-brands to the
	 * destination palette — see {@link ImportSlideOptions}.
	 *
	 * v1 limitations: by default the source slide size must equal this
	 * presentation's — pass `{ rescale: 'fit' | 'stretch' }` to rescale the imported
	 * geometry onto this deck's canvas instead (geometry only, not fonts/line
	 * widths). Source notes are dropped.
	 */
	importSlide(source: Presentation, index: number, options: ImportSlideOptions = {}): Slide {
		const sourceSlide = source.slides[index]
		if (!sourceSlide) throw new Error(`No slide at index ${index} to import`)

		// 1. Pre-flight: slide sizes must match unless the caller opts into a rescale.
		const target = this.slideSize
		const incoming = source.slideSize
		const sizesDiffer = !target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu
		if (sizesDiffer && !options.rescale) {
			const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
			throw new Error(`importSlide requires equal slide sizes (pass { rescale: 'fit' | 'stretch' } to rescale); target is ${fmt(target)}, source is ${fmt(incoming)}`)
		}
		if (sizesDiffer && options.rescale && (!target || !incoming)) {
			throw new Error('importSlide rescale requires both decks to declare a slide size (p:sldSz)')
		}

		// 2. Copy the slide and its dependencies. 'preserve' flattens the theme into
		//    the slide and attaches it to this deck's master; 'restyle' attaches it
		//    to this deck's master with theme refs left symbolic (re-brand); 'copy'
		//    brings the source theme subgraph across wholesale.
		const newPartName =
			options.theme === 'preserve'
				? this.#importSlidePreserve(source, sourceSlide, options.carryMasterGraphics === true)
				: options.theme === 'restyle'
					? this.#importSlideRestyle(source, sourceSlide, options.carryMasterGraphics === true, options.remapLiterals === true)
					: this.#copyPart(source.opc, sourceSlide.partName)
		const newPart = this.opc.part(newPartName)
		if (!newPart) throw new Error(`Imported slide part went missing: ${newPartName}`)

		// 2b. Rescale the imported geometry to this deck's canvas when sizes differ.
		if (sizesDiffer && options.rescale && target && incoming) {
			this.#rescaleImportedGeometry(newPartName, options.theme, incoming, target, options.rescale === true ? 'fit' : options.rescale)
		}

		// 3. Wire the new slide into the presentation (rel + p:sldId entry) at `at`.
		const slide = this.#insertSlidePart(newPart, options.at)

		// 4. Optionally carry the source deck's embedded fonts (presentation-level, so
		//    a separate traversal from the slide-part copy chain above).
		if (options.embedFonts) this.#carryEmbeddedFonts(source)

		return slide
	}

	/**
	 * Copy `source`'s embedded fonts into this deck and merge them into our
	 * `p:embeddedFontLst`. Font binaries come across via {@link #copyPart} (so the
	 * per-source registry dedupes faces shared across repeated imports); entries are
	 * merged by `typeface` + face slot, so a face this deck already embeds is reused
	 * rather than duplicated. No-op when the source embeds no fonts. See
	 * {@link ImportSlideOptions.embedFonts}.
	 */
	#carryEmbeddedFonts(source: Presentation): void {
		const sourceRoot = source.presentationPart.dom.documentElement
		const sourceLst = sourceRoot && firstChild(sourceRoot, 'p:embeddedFontLst')
		const sourceEntries = sourceLst ? getElements(sourceLst, 'p:embeddedFont') : []
		if (sourceEntries.length === 0) return

		const presPart = this.presentationPart
		const presRoot = presPart.dom.documentElement
		if (!presRoot) throw new Error('presentation.xml has no document element to carry embedded fonts into')
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const sourcePresRels = source.opc.relationshipsFor(source.presentationPart.partName)

		// Target list, created at CT_Presentation index 7 when this deck has none yet.
		const targetLst = getOrAddChild(presRoot, 'p:embeddedFontLst', PRESENTATION_EMBEDDED_FONT_LST_SUCCESSORS)
		const targetByTypeface = new Map<string, Element>()
		for (const entry of getElements(targetLst, 'p:embeddedFont')) {
			const font = firstChild(entry, 'p:font')
			const typeface = font && attr(font, 'typeface')
			if (typeface) targetByTypeface.set(typeface, entry)
		}

		let copiedAny = false
		for (const srcEntry of sourceEntries) {
			const srcFont = firstChild(srcEntry, 'p:font')
			const typeface = srcFont ? attr(srcFont, 'typeface') : null
			if (!srcFont || !typeface) continue

			// Find or create the target entry for this typeface, cloning the source
			// p:font identity attributes (typeface + optional panose/pitchFamily/charset).
			let targetEntry = targetByTypeface.get(typeface)
			if (!targetEntry) {
				targetEntry = createElement(presPart.dom, 'p:embeddedFont')
				const targetFont = createElement(presPart.dom, 'p:font')
				for (const name of ['typeface', 'panose', 'pitchFamily', 'charset']) {
					const value = attr(srcFont, name)
					if (value !== null) setAttr(targetFont, name, value)
				}
				targetEntry.appendChild(targetFont)
				targetLst.appendChild(targetEntry)
				targetByTypeface.set(typeface, targetEntry)
			}

			for (const slot of EMBEDDED_FONT_SLOTS) {
				const srcFace = firstChild(srcEntry, `p:${slot}`)
				if (!srcFace) continue
				if (firstChild(targetEntry, `p:${slot}`)) continue // de-dupe: face already present
				const srcRid = attr(srcFace, 'r:id')
				if (!srcRid) continue

				// Ensure the fntdata Default exists *before* copying the part, so #copyPart's
				// addPart resolves the content type via the Default (no per-part Override).
				this.opc.contentTypes.ensureDefault(FONT_DATA_EXTENSION, FONT_DATA_CONTENT_TYPE)
				const newFontPart = this.#copyPart(source.opc, sourcePresRels.resolveTarget(srcRid))
				const relId = presRels.add(FONT_REL_TYPE, relativePartName(presPart.partName, newFontPart)).id

				const targetFace = createElement(presPart.dom, `p:${slot}`)
				setAttr(targetFace, 'r:id', relId)
				insertInOrder(targetEntry, targetFace, EMBEDDED_FONT_FACE_SUCCESSORS[slot])
				copiedAny = true
			}
		}

		if (copiedAny) presPart.markDirty()
	}

	/**
	 * Rescale an imported slide's geometry onto this deck's canvas (the `rescale`
	 * option of {@link importSlide}). Rewrites every top-level shape/group/
	 * graphicFrame transform and table grid on the slide; in `copy` mode also
	 * rescales the imported layout and master shape trees (resolved via the
	 * slide → layout → master rel chain) so inherited placeholder/background geometry
	 * stays aligned. `preserve`/`restyle` rebind to this deck's own master/layout —
	 * already the destination size — so only the slide is touched. Geometry only:
	 * font sizes and line widths are left as authored.
	 */
	#rescaleImportedGeometry(slidePartName: string, theme: ImportSlideOptions['theme'], source: SlideSize, target: SlideSize, mode: 'fit' | 'stretch'): void {
		const transform = computeRescale(source, target, mode)
		this.#rescalePartGeometry(slidePartName, transform)
		if (theme === undefined || theme === 'copy') {
			const layout = this.#resolveSingleRel(this.opc, slidePartName, SLIDE_LAYOUT_REL)
			const master = layout ? this.#resolveSingleRel(this.opc, layout, SLIDE_MASTER_REL) : null
			if (layout) this.#rescalePartGeometry(layout, transform)
			if (master) this.#rescalePartGeometry(master, transform)
		}
	}

	/**
	 * Rescale one part's `p:spTree` geometry in place. Idempotent per part
	 * (#rescaledParts), so a layout/master shared across repeated imports from one
	 * source is rescaled exactly once.
	 */
	#rescalePartGeometry(partName: string, transform: RescaleTransform): void {
		if (this.#rescaledParts.has(partName)) return
		this.#rescaledParts.add(partName)
		const part = this.opc.part(partName)
		const root = part?.dom.documentElement
		const cSld = root && firstChild(root, 'p:cSld')
		const spTree = cSld && firstChild(cSld, 'p:spTree')
		if (!part || !spTree) return
		rescaleSpTree(spTree, transform)
		part.markDirty()
	}

	/**
	 * Graft slide master(s) from another open package into this one and return what
	 * was copied. Unlike {@link importSlide} — which brings a master across only as
	 * the dependency of an imported *slide* and prunes it to the one layout that
	 * slide uses — this copies a master together with its **whole** layout family
	 * and attaches it to no slide: the master and its layouts land in this deck's
	 * layout gallery (PowerPoint's *Insert ▸ New Slide* / *Layout* picker) without
	 * changing any existing slide.
	 *
	 * It is the "ship a brand template's layouts into a generated deck" capability,
	 * kept brand-agnostic here: the caller supplies the source `.pptx`. Each grafted
	 * master is wired into `p:sldMasterIdLst` (so renderers treat it as active) and
	 * its `p:sldLayoutIdLst` is rebuilt to list exactly the copied layouts; the
	 * connected theme/media/tag parts come across under fresh partnames, and parts
	 * shared with earlier imports from the same source are reused (the copy
	 * registry), so a re-call is idempotent. Untouched parts of this package stay
	 * byte-identical, matching {@link importSlide}'s fidelity contract.
	 *
	 * `options.masters` / `options.layouts` narrow what is grafted; by default every
	 * master and every layout comes across. The source and destination slide sizes
	 * must match unless `options.requireEqualSize` is `false` (see
	 * {@link ImportSlideMastersOptions}).
	 *
	 * v1 limitations mirror {@link importSlide}: no geometry rescaling, and
	 * presentation-level embedded fonts are not copied.
	 */
	importSlideMasters(source: Presentation, options: ImportSlideMastersOptions = {}): ImportedSlideMaster[] {
		if (options.requireEqualSize !== false) {
			const target = this.slideSize
			const incoming = source.slideSize
			if (!target || !incoming || target.widthEmu !== incoming.widthEmu || target.heightEmu !== incoming.heightEmu) {
				const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
				throw new Error(`importSlideMasters requires equal slide sizes (pass { requireEqualSize: false } to override); target is ${fmt(target)}, source is ${fmt(incoming)}`)
			}
		}

		const pickMaster = options.masters ?? (() => true)
		const pickLayout = options.layouts ?? (() => true)

		const imported: ImportedSlideMaster[] = []
		source.#slideMasterPartNames().forEach((masterPartName, masterIndex) => {
			if (!pickMaster(cSldName(source.opc.part(masterPartName)), masterIndex)) return

			// Copy the (lean) master first: #copyPart registers it in p:sldMasterIdLst
			// and clears its layout list, then each copied layout re-links itself in.
			const newMasterPartName = this.#copyPart(source.opc, masterPartName)

			const layoutPartNames: string[] = []
			source.#layoutPartNamesOf(masterPartName).forEach((layoutPartName, layoutIndex) => {
				if (!pickLayout(cSldName(source.opc.part(layoutPartName)), layoutIndex)) return
				layoutPartNames.push(this.#copyPart(source.opc, layoutPartName))
			})

			imported.push({ partName: newMasterPartName, layoutPartNames })
		})
		return imported
	}

	/**
	 * The deck's slide layouts, in master then layout order — the gallery a new
	 * slide can bind to. Each {@link LayoutHandle} addresses one layout for
	 * {@link appendSlides}; the `name` is its `p:cSld@name`. Read-only enumeration:
	 * it copies nothing and leaves the package byte-identical.
	 */
	layouts(): LayoutHandle[] {
		const out: LayoutHandle[] = []
		this.#slideMasterPartNames().forEach((masterPartName, masterIndex) => {
			this.#layoutPartNamesOf(masterPartName).forEach((layoutPartName, layoutIndex) => {
				out.push({
					partName: layoutPartName,
					name: cSldName(this.opc.part(layoutPartName)),
					masterPartName,
					masterIndex,
					layoutIndex,
				})
			})
		})
		return out
	}

	/**
	 * Append generator-produced slides onto this deck, binding each to an existing
	 * layout, and return the new {@link Slide}s. This is the hybrid
	 * "generate-onto-existing" path: the deck's masters, layouts, theme — and every
	 * other untouched part — stay **byte-identical** (only `presentation.xml`, its
	 * `.rels`, `[Content_Types].xml`, and the freshly-added slide/media parts
	 * change), because the existing chrome is never regenerated.
	 *
	 * `source` is any slide producer (a `PptxGenJS` instance); its authored slides
	 * are serialized via {@link SlideSource.extractSlides} and spliced in under
	 * fresh partnames, with each slide's `slideLayout` relationship pointed at the
	 * layout named by `options.layout` and its image/hyperlink relationships rebuilt
	 * (preserving the body's relationship ids). Insert position follows
	 * `options.at` (see {@link AppendSlidesOptions}).
	 *
	 * Charts and internal slide-to-slide hyperlinks are carried across: chart parts
	 * (chart XML + `.rels` + embedded workbook) are injected under fresh names, and a
	 * `slide:N` link is repointed at the Nth appended slide's new partname.
	 *
	 * Limitations:
	 * - Audio/video media in an appended slide throws (fixture-gated; backlog
	 *   `dn-append-av-media`).
	 * - An internal link to a source slide outside the appended batch throws (its
	 *   target has no counterpart in the destination).
	 * - Appended slides are concrete absolute-positioned content with no placeholder
	 *   inheritance from the bound layout; the binding governs theme/`clrMap`
	 *   resolution and the "based on" link, not placeholder geometry. Author with
	 *   concrete colours — any `schemeClr` re-resolves against the destination theme.
	 * - Source and destination slide sizes must match (no geometry rescale).
	 * - Notes are not generated.
	 */
	async appendSlides(source: SlideSource, options: AppendSlidesOptions): Promise<Slide[]> {
		// 1. Resolve the target layout partname (explicit; no silent fallback).
		const gallery = this.layouts()
		let target: LayoutHandle
		if (typeof options.layout === 'string') {
			const matches = gallery.filter(l => l.name === options.layout)
			if (matches.length === 0) {
				const names = gallery.map(l => JSON.stringify(l.name)).join(', ')
				throw new Error(`appendSlides: no layout named ${JSON.stringify(options.layout)}; available: ${names || '(none)'}`)
			}
			if (matches.length > 1) {
				throw new Error(`appendSlides: layout name ${JSON.stringify(options.layout)} is ambiguous (${matches.length} layouts share it); pass a LayoutHandle from layouts() instead`)
			}
			target = matches[0]
		} else {
			const handle = options.layout
			if (!gallery.some(l => l.partName === handle.partName)) {
				throw new Error(`appendSlides: layout ${handle.partName} does not belong to this presentation`)
			}
			target = handle
		}

		// 2. Author + extract; enforce equal slide size (no geometry rescale in v1).
		const extracted = await source.extractSlides({ onMediaError: options.onMediaError })
		const size = this.slideSize
		if (!size || size.widthEmu !== extracted.widthEmu || size.heightEmu !== extracted.heightEmu) {
			const fmt = (w: number, h: number): string => `${w}×${h} EMU`
			throw new Error(`appendSlides requires equal slide sizes; target is ${size ? fmt(size.widthEmu, size.heightEmu) : 'unknown'}, source is ${fmt(extracted.widthEmu, extracted.heightEmu)}`)
		}

		// Any existing slide partname seeds the fresh-partname family; fall back to a
		// literal seed for a slide-less template shell (reservePartNameLike parses the
		// string, it does not require the part to exist).
		const slideTemplate = this.slides[0]?.partName ?? '/ppt/slides/slide1.xml'

		// Pass 1: reserve + add every slide body first, so internal slide-to-slide
		// links (which may point forward) can resolve to any appended slide. Adding
		// each part immediately claims its name — reservePartNameLike returns max+1
		// from the existing parts, so the next reservation sees it. (addPart registers
		// the slide's Override content type.)
		const placed = extracted.slides.map(slide => {
			const partName = this.opc.reservePartNameLike(slideTemplate)
			const part = this.opc.addPart(partName, SLIDE_CONTENT_TYPE, textEncoder.encode(slide.xml))
			return { slide, part, partName }
		})

		// 1-based source slide number -> the appended slide's new partname.
		const partBySourceNumber = new Map<number, string>(placed.map((p, i) => [i + 1, p.partName]))

		// Pass 2: build each slide's .rels and wire it into presentation.xml. Media,
		// hyperlinks, charts, and slide-links keep the body's rId (addWithId); the
		// layout rel is added last via add() so its auto-id cannot collide.
		const added: Slide[] = []
		placed.forEach(({ slide, part, partName }, i) => {
			const rels = this.opc.relationshipsFor(partName)
			for (const m of slide.media) {
				const mediaPartName = this.opc.reserveMediaPartName(m.extn)
				this.opc.addPart(mediaPartName, m.contentType, m.bytes)
				rels.addWithId(`rId${m.rId}`, IMAGE_REL, relativePartName(partName, mediaPartName))
			}
			for (const av of slide.avMedia) {
				// One media part backs two rels (ECMA audio/video + MS-2007 media) sharing
				// its Target; the preview poster is a separate image part. ensureDefault
				// runs before addPart so the content type resolves via a Default extension
				// entry (what PowerPoint authors) rather than a per-part Override.
				const mediaPartName = this.opc.reserveMediaPartName(av.mediaExtn, 'media')
				this.opc.contentTypes.ensureDefault(av.mediaExtn, av.mediaContentType)
				this.opc.addPart(mediaPartName, av.mediaContentType, av.mediaBytes)
				const mediaTarget = relativePartName(partName, mediaPartName)
				rels.addWithId(`rId${av.mediaRid}`, av.mtype === 'audio' ? AUDIO_REL : VIDEO_REL, mediaTarget)
				rels.addWithId(`rId${av.msMediaRid}`, MS_MEDIA_REL, mediaTarget)

				const previewPartName = this.opc.reserveMediaPartName(av.previewExtn)
				this.opc.contentTypes.ensureDefault(av.previewExtn, av.previewContentType)
				this.opc.addPart(previewPartName, av.previewContentType, av.previewBytes)
				rels.addWithId(`rId${av.previewRid}`, IMAGE_REL, relativePartName(partName, previewPartName))
			}
			for (const ov of slide.onlineMedia) {
				// Online (external-link) video: two External rels share the link Target — the
				// ECMA video rel and the MS-2007 media rel — with no media binary part and no
				// content-type entry. The poster image is wired by the `slide.media` loop above.
				rels.addWithId(`rId${ov.mediaRid}`, VIDEO_REL, ov.link, 'External')
				rels.addWithId(`rId${ov.msMediaRid}`, MS_MEDIA_REL, ov.link, 'External')
			}
			for (const h of slide.hyperlinks) {
				rels.addWithId(`rId${h.rId}`, HYPERLINK_REL, h.target, 'External')
			}
			for (const c of slide.charts) {
				// Chart part + its embedded workbook, each under a fresh name. The chart
				// XML references the workbook through the chart part's own rId1, so the
				// chart .rels is rebuilt here against the reserved workbook partname.
				const chartPartName = this.opc.reservePartNameLike('/ppt/charts/chart1.xml')
				this.opc.addPart(chartPartName, CHART_CONTENT_TYPE, textEncoder.encode(c.chartXml))
				const embeddingPartName = this.opc.reservePartNameLike('/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
				this.opc.contentTypes.ensureDefault('xlsx', XLSX_CONTENT_TYPE)
				this.opc.addPart(embeddingPartName, XLSX_CONTENT_TYPE, c.embeddingBytes)
				this.opc.relationshipsFor(chartPartName).addWithId('rId1', PACKAGE_REL, relativePartName(chartPartName, embeddingPartName))
				rels.addWithId(`rId${c.rId}`, CHART_REL, relativePartName(partName, chartPartName))
			}
			for (const link of slide.slideLinks) {
				const targetPartName = partBySourceNumber.get(link.sourceSlideNumber)
				if (!targetPartName) {
					throw new Error(`appendSlides: slide ${i} links to source slide ${link.sourceSlideNumber}, which is not among the appended slides`)
				}
				rels.addWithId(`rId${link.rId}`, SLIDE_REL, relativePartName(partName, targetPartName))
			}
			rels.add(SLIDE_LAYOUT_REL, relativePartName(partName, target.partName))

			// Wire into presentation.xml (rel + p:sldId) at the requested position.
			const at = options.at === undefined ? undefined : options.at + i
			added.push(this.#insertSlidePart(part, at))
		})

		return added
	}

	/** Source-side helper: master partnames in `p:sldMasterIdLst` order. */
	#slideMasterPartNames(): string[] {
		const root = this.presentationPart.dom.documentElement
		const lst = root && firstChild(root, 'p:sldMasterIdLst')
		if (!lst) return []
		const rels = this.opc.relationshipsFor(this.presentationPart.partName)
		const out: string[] = []
		for (const entry of getElements(lst, 'p:sldMasterId')) {
			const relId = attr(entry, 'r:id')
			if (relId) out.push(rels.resolveTarget(relId))
		}
		return out
	}

	/** Source-side helper: a master's layout partnames in `p:sldLayoutIdLst` order. */
	#layoutPartNamesOf(masterPartName: string): string[] {
		const root = this.opc.part(masterPartName)?.dom.documentElement
		const lst = root && firstChild(root, 'p:sldLayoutIdLst')
		if (!lst) return []
		const rels = this.opc.relationshipsFor(masterPartName)
		const out: string[] = []
		for (const entry of getElements(lst, 'p:sldLayoutId')) {
			const relId = attr(entry, 'r:id')
			if (relId) out.push(rels.resolveTarget(relId))
		}
		return out
	}

	/**
	 * Copy one shape — an autoshape, picture, table/chart graphic frame, connector,
	 * or group — from `source.shapes[shapeIndex]` onto `target`, returning the new
	 * {@link Shape}. `target` must be a slide of *this* presentation; `source` may
	 * belong to any open presentation.
	 *
	 * The lifted subtree is copied self-consistently: every media/chart/embedding it
	 * depends on is dragged into this package (deduped against earlier imports from
	 * the same source), its `r:embed`/`r:id`/… are rewritten to fresh host-slide
	 * relationships, and its drawing ids (including a group's children) are reassigned
	 * so they cannot collide with the host. With `theme: 'preserve'` (default) the
	 * shape's theme references are baked to literals against the *source* theme so it
	 * renders the same on a foreign host; `restyle` leaves them symbolic to re-brand;
	 * `copy` brings the XML across untouched — see {@link ImportShapeOptions}.
	 *
	 * v1 limitations: source and target slide sizes must match (no geometry rescale,
	 * as with `importSlide`); the source slide's build animation/timing for the shape
	 * is dropped; and lifting a *placeholder* is best-effort — `preserve` bakes its
	 * inherited geometry/colour/size, but for clean results prefer lifting concrete
	 * content shapes/tables/charts over placeholders.
	 */
	importShape(target: Slide, source: Slide, shapeIndex: number, options: ImportShapeOptions = {}): Shape {
		return this.importShapes(target, source, [shapeIndex], options)[0]
	}

	/**
	 * Batch form of {@link importShape}: copy several shapes from one source slide
	 * onto `target` in the given order. Media/chart/embedding parts shared by the
	 * lifted shapes (and by earlier imports from the same source deck) are copied
	 * once via the copy registry, and shared images resolve to a single host-slide
	 * relationship. Returns the new {@link Shape}s in `shapeIndices` order.
	 */
	importShapes(target: Slide, source: Slide, shapeIndices: number[], options: ImportShapeOptions = {}): Shape[] {
		if (target.presentation !== this) throw new Error('importShape: target slide must belong to this presentation')

		// Pre-flight: v1 does not rescale geometry, so slide sizes must match.
		const targetSize = this.slideSize
		const sourceSize = source.presentation.slideSize
		if (!targetSize || !sourceSize || targetSize.widthEmu !== sourceSize.widthEmu || targetSize.heightEmu !== sourceSize.heightEmu) {
			const fmt = (s: SlideSize | null): string => (s ? `${s.widthEmu}×${s.heightEmu} EMU` : 'unknown')
			throw new Error(`importShape requires equal slide sizes; target is ${fmt(targetSize)}, source is ${fmt(sourceSize)}`)
		}

		// Resolve + validate every index up front so a bad batch throws before mutating.
		const sourceShapes = source.shapes
		const sourceElements = shapeIndices.map((i) => {
			const shape = sourceShapes[i]
			if (!shape) throw new Error(`No shape at index ${i} on the source slide (it has ${sourceShapes.length})`)
			return shape.element_
		})

		const spTree = target.shapeTree()
		if (!spTree) throw new Error(`importShape: target slide ${target.partName} has no shape tree`)
		const targetDoc = spTree.ownerDocument
		if (!targetDoc) throw new Error('importShape: target slide DOM has no owner document')

		const theme = options.theme ?? 'preserve'
		const sourceOpc = source.presentation.opc
		const sourceRels = sourceOpc.relationshipsFor(source.partName)
		const targetRels = this.opc.relationshipsFor(target.partName)
		// One rel-id map across the batch so shapes sharing a source image share a rel.
		const relIdMap = new Map<string, string>()
		// preserve: build the source theme context once; copy/restyle need none.
		const ctx = theme === 'preserve' ? this.#sourceFlattenContext(sourceOpc, source.partName) : null

		// Anchor for z-order: the existing shape currently at `at` (insert before it,
		// preserving batch order), else append before any trailing p:extLst.
		const extLst = firstChild(spTree, 'p:extLst')
		const anchor = options.at == null ? extLst : (nthShapeChild(spTree, options.at) ?? extLst)

		const result: Shape[] = []
		for (const shapeEl of sourceElements) {
			const imported = targetDoc.importNode(shapeEl, true) as Element

			// Drag media/charts/embeddings across and rewrite refs to fresh host rels.
			this.#rewriteCarriedRels(imported, sourceOpc, sourceRels, target.partName, targetRels, relIdMap)

			// preserve: bake the source theme onto the subtree. The flatten passes match
			// descendants (not the root), so wrap the shape in a throwaway container.
			if (ctx) {
				const holder = createElement(targetDoc, 'p:spTree')
				holder.appendChild(imported)
				flattenShape(holder, ctx)
			}

			// Give the shape and any group children collision-free host ids.
			let nextId = target.nextShapeId()
			const cNvPrs = imported.getElementsByTagNameNS(OOXML_NS.p, 'cNvPr')
			for (let i = 0; i < cNvPrs.length; i++) setAttr(cNvPrs[i], 'id', String(nextId++))

			// Insert into the host tree (this reparents it out of any holder).
			spTree.insertBefore(imported, anchor)

			const shape = wrapShapeElement(imported, target)
			if (!shape) throw new Error(`importShape: unsupported shape element <${imported.localName}>`)
			if (options.left != null) shape.left = options.left
			if (options.top != null) shape.top = options.top
			if (options.width != null) shape.width = options.width
			if (options.height != null) shape.height = options.height
			result.push(shape)
		}

		target.part.markDirty()
		return result
	}

	/**
	 * Import a slide in `preserve` mode: rebind it to this deck's master/layout
	 * (see {@link #importSlideRebind}), then flatten its source theme into the slide
	 * XML (scheme colours + style-matrix fills baked to literals). Returns the new
	 * partname.
	 *
	 * The flatten context is gathered from the *source* subgraph, so it can be read
	 * before or after the rebind; the rebind injects any carried decorations before
	 * we flatten, so a single sweep resolves the theme references on the slide's own
	 * content and on the carried decorations together.
	 */
	#importSlidePreserve(source: Presentation, sourceSlide: Slide, carryGraphics: boolean): string {
		const ctx = this.#sourceFlattenContext(source.opc, sourceSlide.partName)
		const { newPartName, slideRoot, newPart } = this.#importSlideRebind(source, sourceSlide, carryGraphics)
		flattenSlide(slideRoot, ctx)
		newPart.markDirty()
		return newPartName
	}

	/**
	 * Import a slide in `restyle` mode: rebind it to this deck's master/layout (see
	 * {@link #importSlideRebind}) and then {@link restyleSlide} it — drop its colour
	 * map override but bake *nothing*, so its symbolic theme references re-resolve
	 * against the destination theme and the slide re-brands. Returns the new
	 * partname.
	 *
	 * The deliberate inverse of `preserve`: no flatten, no inherited-background
	 * bake, no placeholder colour/size/geometry bake — every one of those would pin
	 * the slide to its source look, the opposite of re-branding. Carried
	 * decorations are left symbolic too, so they re-brand along with the slide.
	 *
	 * With `remapLiterals` it additionally force-remaps the slide's source-theme
	 * literal colours back to symbolic scheme colours and copies any referenced
	 * source table style into this deck — the two things plain `restyle` cannot
	 * re-brand (see {@link ImportSlideOptions.remapLiterals}).
	 */
	#importSlideRestyle(source: Presentation, sourceSlide: Slide, carryGraphics: boolean, remapLiterals: boolean): string {
		const { newPartName, slideRoot, newPart } = this.#importSlideRebind(source, sourceSlide, carryGraphics)
		restyleSlide(slideRoot)
		if (remapLiterals) {
			// The source colour context (slot ↔ RGB ↔ token) the literals are matched against.
			const parts = resolveSlideThemeParts(source.opc, sourceSlide.partName)
			remapLiteralColors(slideRoot, { clrMap: parts.clrMap, clrScheme: parts.clrScheme })
			this.#copySourceTableStyles(source.opc, slideRoot)
		}
		newPart.markDirty()
		return newPartName
	}

	/**
	 * Copy any table style the restyled slide references from the source
	 * `tableStyles.xml` into this deck's (the `remapLiterals` table leg). A restyled
	 * table resolves its `@tableStyleId` against the *destination* `tableStyles` and
	 * silently falls back when that id is absent; the source `<a:tblStyle>` is itself
	 * symbolic (scheme colours), so copying it verbatim under the **same** id lets the
	 * table keep its structure while re-branding to this deck's theme — no table-XML
	 * rewrite needed. Idempotent: an id this deck already defines is left as-is, and a
	 * referenced id with no source definition is skipped. Creates and wires a
	 * `tableStyles.xml` part if this deck has none.
	 */
	#copySourceTableStyles(sourceOpc: OpcPackage, slideRoot: Element): void {
		const ids = new Set<string>()
		const idEls = slideRoot.getElementsByTagNameNS(OOXML_NS.a, 'tableStyleId')
		for (let i = 0; i < idEls.length; i++) {
			const id = idEls[i].textContent?.trim()
			if (id) ids.add(id)
		}
		if (ids.size === 0) return

		const sourceList = this.#tableStyleList(sourceOpc)
		if (!sourceList) return
		const sourceStyles = new Map(getElements(sourceList, 'a:tblStyle').map((st) => [attr(st, 'styleId'), st] as const))

		const destPart = this.#ensureTableStylesPart()
		const destList = destPart.dom.documentElement
		if (!destList) return
		const present = new Set(getElements(destList, 'a:tblStyle').map((st) => attr(st, 'styleId')))

		let added = false
		for (const id of ids) {
			if (present.has(id)) continue
			const src = sourceStyles.get(id)
			if (!src) continue
			destList.appendChild(destList.ownerDocument!.importNode(src, true))
			present.add(id)
			added = true
		}
		if (added) destPart.markDirty()
	}

	/** The `a:tblStyleLst` root of a package's `tableStyles.xml`, or `null` when it has none. */
	#tableStyleList(opc: OpcPackage): Element | null {
		const part = opc.partsByContentType(TABLE_STYLES_CONTENT_TYPE)[0]
		return part ? part.dom.documentElement : null
	}

	/**
	 * This deck's `tableStyles.xml` part, creating an empty one (and wiring its
	 * `presentation.xml` relationship + content type) when the deck has none.
	 */
	#ensureTableStylesPart(): Part {
		const existing = this.opc.partsByContentType(TABLE_STYLES_CONTENT_TYPE)[0]
		if (existing) return existing
		const partName = this.opc.reservePartNameLike('/ppt/tableStyles.xml')
		const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:tblStyleLst xmlns:a="${OOXML_NS.a}" def="${TABLE_STYLES_DEFAULT_GUID}"/>`
		const part = this.opc.addPart(partName, TABLE_STYLES_CONTENT_TYPE, textEncoder.encode(xml))
		const presRels = this.opc.relationshipsFor(this.presentationPart.partName)
		presRels.add(TABLE_STYLES_REL, relativePartName(this.presentationPart.partName, partName))
		return part
	}

	/**
	 * The rebind shared by `preserve` and `restyle`: copy the slide bytes into a
	 * fresh part, rebuild its relationships (drop notes, repoint the `slideLayout`
	 * rel at this deck's existing layout, copy every other internal target —
	 * media/charts — and pass externals through), and optionally bake the source
	 * master/layout decorations onto the slide. Returns the new part, its name, and
	 * its live root element for the caller's mode-specific pass (flatten vs restyle).
	 *
	 * This carries *no* theme baking of its own — not even the inherited background.
	 * `preserve` adds that via {@link flattenSlide}'s context; `restyle` must not,
	 * so the background stays symbolic and re-brands.
	 */
	#importSlideRebind(source: Presentation, sourceSlide: Slide, carryGraphics: boolean): { newPartName: string; slideRoot: Element; newPart: Part } {
		const destLayout = this.#destinationLayoutPartName()

		// Copy the slide bytes into a fresh partname; we then mutate that copy's DOM
		// (a distinct document, so the source package is never touched).
		const sourcePart = source.opc.part(sourceSlide.partName)
		if (!sourcePart) throw new Error(`importSlide: source package has no part ${sourceSlide.partName}`)
		const newPartName = this.opc.reservePartNameLike(sourceSlide.partName)
		const newPart = this.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
		const slideRoot = newPart.dom.documentElement
		if (!slideRoot) throw new Error(`Imported slide ${newPartName} has no root element`)

		// Rebuild the slide's relationships: drop notes, repoint slideLayout at the
		// destination layout, and copy every other internal target (media/charts).
		const sourceRels = source.opc.relationshipsFor(sourceSlide.partName)
		const targetRels = this.opc.relationshipsFor(newPartName)
		for (const rel of sourceRels) {
			if (rel.type === NOTES_SLIDE_REL) continue
			if (rel.type === SLIDE_LAYOUT_REL) {
				targetRels.addWithId(rel.id, SLIDE_LAYOUT_REL, relativePartName(newPartName, destLayout))
				continue
			}
			if (rel.targetMode === 'External') {
				targetRels.addWithId(rel.id, rel.type, rel.target, 'External')
				continue
			}
			const newTarget = this.#copyPart(source.opc, sourceRels.resolveTarget(rel.id))
			targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTarget))
		}

		// Optionally bake the source master/layout decorations (logos, accent shapes)
		// onto the slide behind its own content. Done after the slide's own rels are
		// in place (so carried media get fresh, non-colliding ids) but before the
		// caller's flatten/restyle pass acts on the carried shapes.
		if (carryGraphics) this.#carryMasterGraphics(source.opc, slideRoot, newPartName, sourceSlide.partName)

		return { newPartName, slideRoot, newPart }
	}

	/**
	 * Bake the source `slideLayout`/`slideMaster` shape-tree decorations onto the
	 * imported slide (the `carryMasterGraphics` path). Every shape on those trees
	 * *except* placeholders is deep-copied into the slide's `p:spTree` ahead of its
	 * own content — master decorations first, then layout, then the slide's shapes —
	 * so document (z-)order keeps the master furthest back. Each decoration's media
	 * and other relationship targets are copied into this package and its
	 * `r:embed`/`r:id`/… references rewritten to fresh slide-local ids. The injected
	 * shapes are left for the caller's {@link flattenSlide} pass to resolve any
	 * theme references they carry.
	 */
	#carryMasterGraphics(sourceOpc: OpcPackage, slideRoot: Element, newPartName: string, slidePartName: string): void {
		const layoutPartName = this.#resolveSingleRel(sourceOpc, slidePartName, SLIDE_LAYOUT_REL)
		const masterPartName = layoutPartName ? this.#resolveSingleRel(sourceOpc, layoutPartName, SLIDE_MASTER_REL) : null
		const cSld = firstChild(slideRoot, 'p:cSld')
		const spTree = cSld && firstChild(cSld, 'p:spTree')
		if (!spTree) return

		const doc = slideRoot.ownerDocument!
		const slideRels = this.opc.relationshipsFor(newPartName)
		const relIdMap = new Map<string, string>()
		// Insert ahead of the slide's own first shape so decorations render behind it.
		const anchor = firstShapeChild(spTree)
		// Master behind layout behind the slide (document order == z-order).
		for (const partName of [masterPartName, layoutPartName]) {
			if (!partName) continue
			const decorations = carriedDecorations(sourceOpc.part(partName)?.dom.documentElement ?? null)
			if (decorations.length === 0) continue
			const sourceRels = sourceOpc.relationshipsFor(partName)
			for (const deco of decorations) {
				const imported = doc.importNode(deco, true) as Element
				this.#rewriteCarriedRels(imported, sourceOpc, sourceRels, newPartName, slideRels, relIdMap)
				spTree.insertBefore(imported, anchor)
			}
		}
	}

	/**
	 * Rewrite every relationship reference (`r:embed`, `r:id`, `r:link`, …) inside a
	 * carried decoration so it points at a fresh slide-local relationship, copying
	 * the referenced part into this package on first sight. `relIdMap` (keyed by
	 * source part + source rel id) dedupes references shared within one import call.
	 */
	#rewriteCarriedRels(node: Element, sourceOpc: OpcPackage, sourceRels: Relationships, newPartName: string, slideRels: Relationships, relIdMap: Map<string, string>): void {
		const elements: Element[] = []
		collectElements(node, elements)
		for (const el of elements) {
			const refs: { local: string; id: string }[] = []
			const attrs = el.attributes
			for (let i = 0; i < attrs.length; i++) {
				const a = attrs.item(i)
				if (!a || a.namespaceURI !== OOXML_NS.r || !a.value) continue
				if (!sourceRels.get(a.value)) continue // an r-namespaced attribute that isn't a relationship id
				refs.push({ local: a.localName, id: a.value })
			}
			for (const { local, id } of refs) {
				setAttr(el, `r:${local}`, this.#carryRel(sourceOpc, sourceRels, id, newPartName, slideRels, relIdMap))
			}
		}
	}

	/** Resolve a carried decoration's source relationship to a fresh slide-local id, copying its internal target. */
	#carryRel(sourceOpc: OpcPackage, sourceRels: Relationships, id: string, newPartName: string, slideRels: Relationships, relIdMap: Map<string, string>): string {
		const key = `${sourceRels.sourcePartName}|${id}`
		const cached = relIdMap.get(key)
		if (cached) return cached
		const rel = sourceRels.get(id)!
		const newId =
			rel.targetMode === 'External'
				? slideRels.add(rel.type, rel.target, 'External').id
				: slideRels.add(rel.type, relativePartName(newPartName, this.#copyPart(sourceOpc, sourceRels.resolveTarget(id)))).id
		relIdMap.set(key, newId)
		return newId
	}

	/**
	 * The partname of the layout this deck's slides should attach to in `preserve`
	 * mode: the first layout of the first slide master. Throws when the deck has no
	 * master/layout to attach to (a deck pptxgenjs always provides).
	 */
	#destinationLayoutPartName(): string {
		const presRels = this.opc.relationshipsFor(this.presentationPart.partName)
		const masterRel = presRels.byType(SLIDE_MASTER_REL)[0]
		if (!masterRel) throw new Error('importSlide preserve mode requires a slide master in the destination deck')
		const masterPartName = presRels.resolveTarget(masterRel.id)
		const masterRels = this.opc.relationshipsFor(masterPartName)
		const layoutRel = masterRels.byType(SLIDE_LAYOUT_REL)[0]
		if (!layoutRel) throw new Error('importSlide preserve mode requires a slide layout in the destination deck')
		return masterRels.resolveTarget(layoutRel.id)
	}

	/**
	 * Gather the flatten context for a source slide: walk slide → layout → master →
	 * theme, reading the effective colour map (the slide's `clrMapOvr` override, or
	 * the master `clrMap`), the theme `clrScheme`, and the theme `fmtScheme`.
	 */
	#sourceFlattenContext(sourceOpc: OpcPackage, slidePartName: string): FlattenContext {
		// Reuse the shared slide → layout → master → theme walk (also backing the
		// read-model colour getters), then layer the flatten-only needs on top.
		const parts = resolveSlideThemeParts(sourceOpc, slidePartName)
		const themeElements = parts.themeElements
		return {
			clrMap: parts.clrMap,
			clrScheme: parts.clrScheme,
			fmtScheme: themeElements ? firstChild(themeElements, 'a:fmtScheme') : null,
			inheritedBackground: this.#effectiveBackground(sourceOpc, parts.slideRoot, parts.layoutPartName, parts.masterPartName),
			layoutRoot: parts.layoutRoot,
			masterRoot: parts.masterRoot,
		}
	}

	/**
	 * The background the slide effectively inherits from its source subgraph: the
	 * layout's `p:bg`, else the master's. Returns `null` when the slide carries its
	 * own `p:bg` (it stays on the slide and is flattened directly) or none exists.
	 */
	#effectiveBackground(sourceOpc: OpcPackage, slideRoot: Element | null, layoutPartName: string | null, masterPartName: string | null): Element | null {
		if (slideRoot && this.#backgroundOf(slideRoot)) return null
		const layoutRoot = layoutPartName ? (sourceOpc.part(layoutPartName)?.dom.documentElement ?? null) : null
		const masterRoot = masterPartName ? (sourceOpc.part(masterPartName)?.dom.documentElement ?? null) : null
		return (layoutRoot && this.#backgroundOf(layoutRoot)) ?? (masterRoot && this.#backgroundOf(masterRoot)) ?? null
	}

	/** The `p:cSld/p:bg` element of a slide/layout/master root, or `null`. */
	#backgroundOf(root: Element): Element | null {
		const cSld = firstChild(root, 'p:cSld')
		return cSld ? firstChild(cSld, 'p:bg') : null
	}

	/** Resolve the single relationship of `type` owned by `partName`, or `null`. */
	#resolveSingleRel(sourceOpc: OpcPackage, partName: string, type: string): string | null {
		const rels = sourceOpc.relationshipsFor(partName)
		const rel = rels.byType(type)[0]
		return rel ? rels.resolveTarget(rel.id) : null
	}

	/** The copy registry for one source package (created on first use). */
	#registryFor(sourceOpc: OpcPackage): Map<string, string> {
		let registry = this.#importRegistry.get(sourceOpc)
		if (!registry) {
			registry = new Map()
			this.#importRegistry.set(sourceOpc, registry)
		}
		return registry
	}

	/**
	 * Copy `sourcePartName` (and, recursively, every internal part it references)
	 * from `sourceOpc` into this package, returning the new partname. Idempotent
	 * per source package via the copy registry. Relationship ids are preserved so
	 * the copied part body's `r:id`/`r:embed` references stay valid; targets are
	 * rewritten to the freshly-allocated partnames. Notes relationships are
	 * dropped. A copied `slideMaster` does not drag in all its sibling layouts —
	 * each imported `slideLayout` wires itself into the master instead (see
	 * {@link #linkLayoutIntoMaster}).
	 */
	#copyPart(sourceOpc: OpcPackage, sourcePartName: string): string {
		const registry = this.#registryFor(sourceOpc)
		const existing = registry.get(sourcePartName)
		if (existing) return existing

		const sourcePart = sourceOpc.part(sourcePartName)
		if (!sourcePart) throw new Error(`importSlide: source package has no part ${sourcePartName}`)

		const newPartName = this.opc.reservePartNameLike(sourcePartName)
		this.opc.addPart(newPartName, sourcePart.contentType, sourcePart.bytes)
		// Record before recursing so the master↔layout cycle terminates.
		registry.set(sourcePartName, newPartName)

		const isMaster = sourcePart.contentType === SLIDE_MASTER_CONTENT_TYPE
		const sourceRels = sourceOpc.relationshipsFor(sourcePartName)
		const targetRels = this.opc.relationshipsFor(newPartName)
		for (const rel of sourceRels) {
			// Notes pull in a notesMaster + its own theme; an imported slide does not need them.
			if (rel.type === NOTES_SLIDE_REL) continue
			// Lean master: skip its layout rels; copied layouts re-link themselves.
			if (isMaster && rel.type === SLIDE_LAYOUT_REL) continue
			if (rel.targetMode === 'External') {
				targetRels.addWithId(rel.id, rel.type, rel.target, 'External')
				continue
			}
			const newTargetPartName = this.#copyPart(sourceOpc, sourceRels.resolveTarget(rel.id))
			targetRels.addWithId(rel.id, rel.type, relativePartName(newPartName, newTargetPartName))
		}

		if (isMaster) {
			this.#clearLayoutIdList(newPartName)
			// Register the copied master in presentation.xml. Without a
			// `p:sldMasterId` entry (and a presentation→master relationship) the
			// master is inert: PowerPoint/LibreOffice ignore its background and shape
			// tree, so a `copy`-imported slide whose look lives on its master (a
			// cover/closer) renders blank. Idempotent, so masters shared across
			// repeated imports are registered exactly once.
			this.#registerMaster(newPartName)
		}
		if (sourcePart.contentType === SLIDE_LAYOUT_CONTENT_TYPE) {
			this.#linkLayoutIntoMaster(sourceOpc, sourceRels, newPartName)
		}

		return newPartName
	}

	/**
	 * Wire a freshly-copied slide master into `presentation.xml`: add a
	 * presentation→master relationship and a `p:sldMasterId` entry in
	 * `p:sldMasterIdLst`. A master that is reachable only through the
	 * slide→layout→master rel chain but absent from `p:sldMasterIdLst` is treated
	 * as inactive by renderers, so its background/graphics never paint. No-op when
	 * the master is already registered (shared across imports from one source).
	 */
	#registerMaster(masterPartName: string): void {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		for (const rel of presRels.byType(SLIDE_MASTER_REL)) {
			if (presRels.resolveTarget(rel.id) === masterPartName) return
		}
		const relId = presRels.add(SLIDE_MASTER_REL, relativePartName(presPart.partName, masterPartName)).id

		const root = presPart.dom.documentElement
		if (!root) throw new Error('presentation.xml has no document element to register a master in')
		// `p:sldMasterIdLst` is the first child of CT_Presentation; create it before
		// any later sibling if a (degenerate) deck lacks one.
		const lst = getOrAddChild(root, 'p:sldMasterIdLst', [
			'p:notesMasterIdLst',
			'p:handoutMasterIdLst',
			'p:sldIdLst',
			'p:sldSz',
			'p:notesSz',
			'p:embeddedFontLst',
			'p:custShowLst',
			'p:photoAlbum',
			'p:custDataLst',
			'p:kinsoku',
			'p:defaultTextStyle',
			'p:modifyVerifier',
			'p:extLst',
		])
		const entry = createElement(presPart.dom, 'p:sldMasterId')
		setAttr(entry, 'id', String(this.#nextSlideMasterId(lst)))
		setAttr(entry, 'r:id', relId)
		lst.appendChild(entry)
		presPart.markDirty()
	}

	/** A slide-master id one past the highest in `sldMasterIdLst`, floored at ST_SlideMasterId's minimum. */
	#nextSlideMasterId(sldMasterIdLst: Element): number {
		let max = MIN_SLIDE_MASTER_ID - 1
		for (const entry of getElements(sldMasterIdLst, 'p:sldMasterId')) {
			const id = intValue(attr(entry, 'id'))
			if (id !== null && id > max) max = id
		}
		return max + 1
	}

	/** Empty a freshly-copied master's `p:sldLayoutIdLst`; copied layouts repopulate it. */
	#clearLayoutIdList(masterPartName: string): void {
		const masterPart = this.opc.part(masterPartName)
		const root = masterPart?.dom.documentElement
		const lst = root && firstChild(root, 'p:sldLayoutIdLst')
		if (!masterPart || !lst) return
		removeChildrenByQName(lst, ['p:sldLayoutId'])
		masterPart.markDirty()
	}

	/**
	 * Wire a just-copied layout into its (already-copied) master: add a
	 * master→layout relationship and append a `p:sldLayoutId` entry. Called once
	 * per copied layout, so the master accumulates exactly the imported layouts.
	 */
	#linkLayoutIntoMaster(sourceOpc: OpcPackage, layoutSourceRels: ReturnType<OpcPackage['relationshipsFor']>, layoutPartName: string): void {
		const masterRel = layoutSourceRels.byType(SLIDE_MASTER_REL)[0]
		if (!masterRel) return
		const masterPartName = this.#registryFor(sourceOpc).get(layoutSourceRels.resolveTarget(masterRel.id))
		if (!masterPartName) return
		const masterPart = this.opc.part(masterPartName)
		const root = masterPart?.dom.documentElement
		if (!masterPart || !root) return

		const masterRels = this.opc.relationshipsFor(masterPartName)
		const relId = masterRels.add(SLIDE_LAYOUT_REL, relativePartName(masterPartName, layoutPartName)).id
		const lst = getOrAddChild(root, 'p:sldLayoutIdLst', ['p:transition', 'p:timing', 'p:hf', 'p:txStyles', 'p:extLst'])
		const entry = createElement(masterPart.dom, 'p:sldLayoutId')
		setAttr(entry, 'id', String(this.#nextSlideLayoutId(lst)))
		setAttr(entry, 'r:id', relId)
		lst.appendChild(entry)
		masterPart.markDirty()
	}

	/**
	 * Wire a new slide part into `p:sldIdLst` (rel + `p:sldId`) at zero-based
	 * position `at` and return it. `p:sldIdLst` order *is* deck order, so the
	 * insertion point is the only bookkeeping needed. An `at` that is omitted,
	 * negative, or `>=` the current slide count appends (the prior behaviour).
	 */
	#insertSlidePart(newPart: Part, at?: number): Slide {
		const presPart = this.presentationPart
		const presRels = this.opc.relationshipsFor(presPart.partName)
		const relId = presRels.add(SLIDE_REL, relativePartName(presPart.partName, newPart.partName)).id

		const root = presPart.dom.documentElement
		if (!root) throw new Error('presentation.xml has no document element to append a slide to')
		// A template with zero slides omits p:sldIdLst entirely; create it in
		// CT_Presentation document order (after the *IdLst children, before p:sldSz).
		const sldIdLst = getOrAddChild(root, 'p:sldIdLst', PRESENTATION_SLD_ID_LST_SUCCESSORS)
		const existing = getElements(sldIdLst, 'p:sldId')
		const newSlideId = this.#nextSlideId(existing)
		const sldId = createElement(presPart.dom, 'p:sldId')
		setAttr(sldId, 'id', String(newSlideId))
		setAttr(sldId, 'r:id', relId)

		const inRange = at !== undefined && at >= 0 && at < existing.length
		const newIndex = inRange ? at : existing.length
		if (inRange) sldIdLst.insertBefore(sldId, existing[at])
		else sldIdLst.appendChild(sldId)
		presPart.markDirty()

		return new Slide(this, newPart, newSlideId, newIndex)
	}

	/** A slide id one past the highest existing, but at least ST_SlideId's minimum. */
	#nextSlideId(sldIds: ReturnType<typeof getElements>): number {
		let max = MIN_SLIDE_ID - 1
		for (const sldId of sldIds) {
			const id = intValue(attr(sldId, 'id'))
			if (id !== null && id > max) max = id
		}
		return max + 1
	}

	/** A slide-layout id one past the highest in `sldLayoutIdLst`, floored at ST_SlideLayoutId's minimum. */
	#nextSlideLayoutId(sldLayoutIdLst: Element): number {
		let max = MIN_SLIDE_LAYOUT_ID - 1
		for (const entry of getElements(sldLayoutIdLst, 'p:sldLayoutId')) {
			const id = intValue(attr(entry, 'id'))
			if (id !== null && id > max) max = id
		}
		return max + 1
	}

	/** Re-emit the package; untouched parts stay byte-identical (see `OpcPackage.save`). */
	async save(): Promise<Uint8Array> {
		return this.opc.save()
	}
}

/** The `p:cSld@name` of a slide/layout/master part (`''` when absent). */
function cSldName(part: Part | undefined): string {
	const root = part?.dom.documentElement
	const cSld = root && firstChild(root, 'p:cSld')
	return (cSld && attr(cSld, 'name')) ?? ''
}

/** Whether `el` is a `p:spTree`'s own group properties (`p:nvGrpSpPr`/`p:grpSpPr`), not a shape. */
function isSpTreeProperty(el: Element): boolean {
	return el.namespaceURI === OOXML_NS.p && (el.localName === 'nvGrpSpPr' || el.localName === 'grpSpPr')
}

/** First direct child *element* of `parent` (skipping text/comment nodes), or `null`. */
function firstChildElement(parent: Element): Element | null {
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) return node as Element
	}
	return null
}

/** Whether a `p:spTree` child is a placeholder shape (its `*nvPr` carries a `p:ph`). */
function isPlaceholderShape(shape: Element): boolean {
	const nv = firstChildElement(shape)
	const nvPr = nv && firstChild(nv, 'p:nvPr')
	return !!(nvPr && firstChild(nvPr, 'p:ph'))
}

/** The decorative shapes on a layout/master `p:spTree`: every shape child except placeholders. */
function carriedDecorations(root: Element | null): Element[] {
	if (!root) return []
	const cSld = firstChild(root, 'p:cSld')
	const spTree = cSld && firstChild(cSld, 'p:spTree')
	if (!spTree) return []
	const out: Element[] = []
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (isSpTreeProperty(el) || isPlaceholderShape(el)) continue
		out.push(el)
	}
	return out
}

/**
 * The `n`-th shape child of a `p:spTree` in document (z-)order, skipping the
 * tree's own `nvGrpSpPr`/`grpSpPr` and any trailing `p:extLst`. Returns `null`
 * when `n` is past the last shape (the caller then appends).
 */
function nthShapeChild(spTree: Element, n: number): Element | null {
	let i = 0
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (isSpTreeProperty(el)) continue
		if (el.namespaceURI === OOXML_NS.p && el.localName === 'extLst') continue
		if (i === n) return el
		i++
	}
	return null
}

/** The first shape child of a `p:spTree` (skipping `nvGrpSpPr`/`grpSpPr`), or `null`. */
function firstShapeChild(spTree: Element): Element | null {
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (!isSpTreeProperty(el)) return el
	}
	return null
}

/** Collect `node` and all its descendant elements (document order) into `out`. */
function collectElements(node: Element, out: Element[]): void {
	out.push(node)
	for (let child = node.firstChild; child; child = child.nextSibling) {
		if (child.nodeType === ELEMENT_NODE) collectElements(child as Element, out)
	}
}

/** An EMU coordinate remap: `newX = x*sx + dx`, `newY = y*sy + dy`; sizes scale by `sx`/`sy` only. */
interface RescaleTransform {
	sx: number
	sy: number
	dx: number
	dy: number
}

/** Build the EMU transform mapping source-canvas coordinates onto the target canvas (see {@link ImportSlideOptions.rescale}). */
function computeRescale(source: SlideSize, target: SlideSize, mode: 'fit' | 'stretch'): RescaleTransform {
	if (mode === 'stretch') {
		return { sx: target.widthEmu / source.widthEmu, sy: target.heightEmu / source.heightEmu, dx: 0, dy: 0 }
	}
	// 'fit': uniform scale (no distortion), centering the slack on the longer axis.
	const scale = Math.min(target.widthEmu / source.widthEmu, target.heightEmu / source.heightEmu)
	return {
		sx: scale,
		sy: scale,
		dx: (target.widthEmu - source.widthEmu * scale) / 2,
		dy: (target.heightEmu - source.heightEmu * scale) / 2,
	}
}

/**
 * Rescale every top-level transform on a `p:spTree`: shapes/pictures/connectors
 * (`p:spPr/a:xfrm`), groups (`p:grpSpPr/a:xfrm` — only the group's own off/ext, so
 * its children remap via the unchanged `chOff`/`chExt` rather than being recursed
 * into), and graphic frames (`p:graphicFrame/p:xfrm`, plus any inner table grid).
 * Placeholders that inherit geometry carry no `a:xfrm` and are left to inherit from
 * the (separately rescaled) layout/master.
 */
function rescaleSpTree(spTree: Element, t: RescaleTransform): void {
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (el.namespaceURI !== OOXML_NS.p) continue
		switch (el.localName) {
			case 'sp':
			case 'pic':
			case 'cxnSp': {
				const spPr = firstChild(el, 'p:spPr')
				rescaleXfrm(spPr && firstChild(spPr, 'a:xfrm'), t)
				break
			}
			case 'grpSp': {
				const grpSpPr = firstChild(el, 'p:grpSpPr')
				rescaleXfrm(grpSpPr && firstChild(grpSpPr, 'a:xfrm'), t)
				break
			}
			case 'graphicFrame': {
				rescaleXfrm(firstChild(el, 'p:xfrm'), t)
				rescaleTableGrid(el, t)
				break
			}
		}
	}
}

/** Rewrite an `a:xfrm`/`p:xfrm`: `a:off` is repositioned (scale + translate); `a:ext` is resized (scale only). */
function rescaleXfrm(xfrm: Element | null, t: RescaleTransform): void {
	if (!xfrm) return
	const off = firstChild(xfrm, 'a:off')
	if (off) {
		const x = intValue(attr(off, 'x'))
		const y = intValue(attr(off, 'y'))
		if (x !== null) setAttr(off, 'x', String(Math.round(x * t.sx + t.dx)))
		if (y !== null) setAttr(off, 'y', String(Math.round(y * t.sy + t.dy)))
	}
	const ext = firstChild(xfrm, 'a:ext')
	if (ext) {
		const cx = intValue(attr(ext, 'cx'))
		const cy = intValue(attr(ext, 'cy'))
		if (cx !== null) setAttr(ext, 'cx', String(Math.max(0, Math.round(cx * t.sx))))
		if (cy !== null) setAttr(ext, 'cy', String(Math.max(0, Math.round(cy * t.sy))))
	}
}

/**
 * Scale a graphic-frame table's intrinsic dimensions — column widths by `sx`, row
 * heights by `sy` — so the table resizes with the slide (PowerPoint derives table
 * geometry from the grid, not the frame `a:ext`). No-op for non-table frames.
 */
function rescaleTableGrid(graphicFrame: Element, t: RescaleTransform): void {
	const graphic = firstChild(graphicFrame, 'a:graphic')
	const graphicData = graphic && firstChild(graphic, 'a:graphicData')
	const tbl = graphicData && firstChild(graphicData, 'a:tbl')
	if (!tbl) return
	const grid = firstChild(tbl, 'a:tblGrid')
	if (grid) {
		for (const col of getElements(grid, 'a:gridCol')) {
			const w = intValue(attr(col, 'w'))
			if (w !== null) setAttr(col, 'w', String(Math.max(0, Math.round(w * t.sx))))
		}
	}
	for (const tr of getElements(tbl, 'a:tr')) {
		const h = intValue(attr(tr, 'h'))
		if (h !== null) setAttr(tr, 'h', String(Math.max(0, Math.round(h * t.sy))))
	}
}
