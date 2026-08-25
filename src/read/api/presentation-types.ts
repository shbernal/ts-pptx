/**
 * Public option and result shapes for the read-model {@link ./presentation} API.
 *
 * These live apart from `Presentation` itself because they are the documented contract
 * callers program against -- `src/read.ts` re-exports them -- while the class is the
 * implementation. Keeping them here also keeps the class file to the behaviour.
 */

import type { EmbeddedFont } from '../../embedded-fonts.js'
// Type-only: keeps this contract module free of any runtime dependency on the
// class that implements it (the import is erased, so no cycle survives to ESM).
import type { Presentation } from './presentation.js'

/** Slide dimensions, in both EMU (the OOXML unit) and inches. */
export interface SlideSize {
	widthEmu: number
	heightEmu: number
	widthIn: number
	heightIn: number
}

/**
 * One `p:embeddedFont` entry, read from `presentation.xml`'s `p:embeddedFontLst`
 * (see {@link Presentation.embeddedFonts}). A read-only view: it names the family
 * and resolves each embedded face's `r:id` to the absolute partname of its binary
 * `/ppt/fonts/*.fntdata` part, so a consumer can enumerate the embedded faces and
 * pull their bytes without hand-parsing the presentation part or its rels.
 */
export interface EmbeddedFontInfo {
	/** `p:font/@typeface` — the family name PowerPoint binds the embed to. */
	typeface: string
	/** `p:font/@panose`, or `null` when the entry declares none. */
	panose: string | null
	/** The embedded faces, in `regular, bold, italic, boldItalic` schema order. */
	faces: EmbeddedFontFaceInfo[]
}

/** One face of an {@link EmbeddedFontInfo}: which slot it fills and the part its `r:id` resolves to. */
export interface EmbeddedFontFaceInfo {
	/** Which weight/style slot this face fills. */
	slot: 'regular' | 'bold' | 'italic' | 'boldItalic'
	/** Absolute partname of the face's binary `.fntdata` part (e.g. `/ppt/fonts/font1.fntdata`). */
	partName: string
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

	/**
	 * Carry the source slide's speaker notes across. By default (`false`) the
	 * `notesSlide` relationship is dropped, so the imported slide has no notes.
	 *
	 * When set, the source `notesSlide` is copied and wired to the imported slide.
	 * Its `slide` back-relationship is repointed at the new slide (the source slide
	 * is *not* copied). A presentation has at most one `notesMaster`: if this deck
	 * already has one, the imported notes reuse it (the source notesMaster and its
	 * theme are not copied); if it has none, the source notesMaster (and its theme)
	 * are copied and registered. The destination's notes styling therefore wins
	 * when both decks define one.
	 * @default false
	 */
	importNotes?: boolean
}

/** Options for {@link Presentation.importShape} / {@link Presentation.importShapes}. */
export interface ImportShapeOptions {
	/**
	 * How the lifted shape relates to themes, mirroring {@link ImportSlideOptions}
	 * but scoped to one shape subtree:
	 *
	 * - `'preserve'` (default): bake the shape's scheme/style-matrix colours (and,
	 *   for a lifted placeholder, its inherited geometry/colour/size/anchor/list
	 *   style) to literals using the *source* slide's theme, so it keeps its look on
	 *   a host slide whose theme differs. A lifted placeholder is also *demoted* to a
	 *   plain shape (its `p:ph` stripped) once everything it inherited is baked, so it
	 *   neither re-resolves against the host's placeholder of the same type/idx nor
	 *   collides with it. The safe default for composing across decks.
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
	/**
	 * What to do when the source slide size differs from this deck's. Default
	 * (`false`/omitted) throws on any mismatch. Set it to rescale the lifted
	 * shape's geometry onto this deck's canvas instead (mirrors
	 * {@link ImportSlideOptions.rescale}, scoped to one shape):
	 *
	 * - `'fit'` (or `true`): uniform scale by `min(sx, sy)`, slack centered — the
	 *   shape keeps its aspect ratio (circles stay circles, rotations hold) and
	 *   lands at the position it would occupy on a centered copy of the source
	 *   canvas. Matches PowerPoint's "Ensure Fit".
	 * - `'stretch'`: independent per-axis scale; distorts. Matches "Maximize".
	 *
	 * Only *geometry* is rescaled — the shape/group/graphicFrame transform
	 * (`a:off`/`a:ext`) and any table grid (`a:gridCol@w`, `a:tr@h`). Font sizes,
	 * line widths, and other absolute sizes are left as authored, so heavy
	 * down-scaling can leave text overflowing its (now smaller) box. Explicit
	 * `left`/`top`/`width`/`height` overrides are applied *after* rescale and win.
	 * @default false
	 */
	rescale?: boolean | 'fit' | 'stretch'
	/**
	 * Carry the lifted shape's build animation across (default `false`). A shape's
	 * entrance/emphasis/exit build lives in the slide-scoped `p:timing`, not in the
	 * shape subtree, so it is dropped unless opted in. When set, the shape's effect
	 * click-group(s) and `<p:bldP>` are copied into the destination `p:timing`
	 * (created if absent), their `spid` references remapped to the shape's new id and
	 * their `<p:cTn>` ids renumbered to stay collision-free, and appended after any
	 * existing build — mirroring how PowerPoint merges a pasted shape's animation
	 * (see the `import-animation-merge` oracle). @default false
	 */
	carryAnimation?: boolean
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

	/**
	 * Carry the *source deck's* embedded fonts (`p:embeddedFontLst` in its
	 * `presentation.xml`) into this deck, so a slide later bound to a grafted layout
	 * keeps that layout's embedded faces on machines that lack the font locally.
	 * Off by default — fonts live on the presentation, not the master, and copying
	 * them can add megabytes — so they are only worth carrying when you want the
	 * embed to travel with the gallery.
	 *
	 * Semantics are identical to {@link ImportSlideOptions.embedFonts}: binaries are
	 * copied under fresh `/ppt/fonts/` names (deduped via the per-source copy
	 * registry, so a re-call carries each face once), the `application/x-fontdata`
	 * Default is added, and entries merge into this deck's `p:embeddedFontLst`
	 * de-duplicated by `typeface` + face slot. The carry is a whole-deck operation —
	 * it copies *all* the source's embedded fonts, not only the faces the grafted
	 * masters use, since the source list does not record which face belongs to which
	 * master.
	 * @default false
	 */
	embedFonts?: boolean

	/**
	 * Carry the *source deck's* table styles (`ppt/tableStyles.xml`) into this deck,
	 * so a table inserted on a grafted layout picks up the source's table styling
	 * instead of this deck's. Off by default: it rewrites an existing part rather than
	 * only adding new ones, so it is opt-in like {@link embedFonts}.
	 *
	 * Without it a grafted master's layouts arrive but its table styling does not, and
	 * a new table falls back to whatever default this deck carries — for a generated
	 * deck, a stub whose default is the standard *Medium Style 2 - Accent 1*. The
	 * mismatch is visible: the same table renders in a different accent than it would
	 * in the source deck.
	 *
	 * Styles union by `styleId` (a style this deck already defines wins, so a re-call
	 * is idempotent), while `a:tblStyleLst@def` — the default table style — is
	 * **source-wins**. Carrying the styles without the `def` would not fix the
	 * mismatch: the standard default GUID is one most templates *also* define, so the
	 * default would still resolve to the wrong style. The carry is whole-deck — it
	 * copies *all* the source's table styles, not only those the grafted masters use,
	 * since `tableStyles.xml` does not record which style belongs to which master.
	 * @default false
	 */
	tableStyles?: boolean

	/**
	 * Move the grafted masters to the front of `p:sldMasterIdLst`, ahead of the ones
	 * this deck already had, keeping their import order. Off by default: which master
	 * leads is a statement about what the deck *is*, so it is the caller's call rather
	 * than a side effect of grafting.
	 *
	 * `p:sldMasterIdLst` order is not part of theme resolution — a slide resolves
	 * through its own layout's master, so this changes no existing slide's appearance.
	 * What it changes is the deck's identity in PowerPoint's UI: the list's first entry
	 * becomes `Designs(1)`, which is the theme the Design tab shows and the one
	 * Design ▸ Variants applies. Graft a brand master into a generated deck without
	 * this and the deck still presents as the generator's stock theme, because the
	 * stub master was there first.
	 *
	 * Reordering rewrites nothing but the list: relationships, ids, and every part
	 * outside `presentation.xml` are untouched, and a re-call is a no-op once the
	 * grafted masters already lead.
	 * @default false
	 */
	primary?: boolean
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
	charts: ExtractedChart[]
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
	/**
	 * The slide's speaker notes, or `undefined` when it has none. Carried as a ready
	 * `<p:notes>` part body plus its external hyperlink rels;
	 * {@link Presentation.appendSlides} creates the notes part, wires it to the appended
	 * slide, and binds it to a notes master (see {@link ExtractedSlides.notesMaster}).
	 */
	notes?: ExtractedNotes
}

/** One slide's speaker notes, extracted for {@link Presentation.appendSlides}. */
interface ExtractedNotes {
	/** Standalone `<p:notes>` part body (XML declaration + namespaces included). */
	xml: string
	/**
	 * External hyperlink rels the notes body references, keyed by the `rId` used in
	 * {@link xml}. Notes rels reserve `rId1` for the notes master and `rId2` for the
	 * slide, so these start at `rId3`. Notes support external `url` links only.
	 */
	hyperlinks: Array<{ rId: number; target: string }>
}

/**
 * One chart extracted for {@link Presentation.appendSlides}, keyed by the `rId` the slide
 * body uses to reference it.
 *
 * A classic (2007, `c:chartSpace`) chart is one part plus its embedded workbook. A chartEx
 * (Office 2016 — waterfall, funnel, treemap, ...) chart is a different part in a different
 * namespace, reached through a different relationship type, and PowerPoint reports it as
 * corrupt without its two style sidecars — so {@link chartEx} carries them and its presence
 * is what tells the two shapes apart.
 */
interface ExtractedChart {
	/** Body `rId` of the slide's rel pointing at the chart part. */
	rId: number
	/** The chart part body: `<c:chartSpace>` for a classic chart, `<cx:chartSpace>` for a chartEx one. */
	chartXml: string
	/** The chart's embedded `.xlsx` workbook, the data source PowerPoint opens on edit. */
	embeddingBytes: Uint8Array
	/**
	 * Present only for a chartEx chart, and then always with both sidecars. A chartEx part
	 * *requires* a chart-style and a color-style part or PowerPoint reports the deck as
	 * corrupt (`0x80070570`) — schema-valid but unopenable — so these are not optional and
	 * {@link Presentation.appendSlides} injects them alongside the chart part.
	 */
	chartEx?: {
		/** `cs:chartStyle` body for the chart's `style{N}.xml` part. */
		styleXml: string
		/** `cs:colorStyle` body for the chart's `colors{N}.xml` part. */
		colorsXml: string
	}
}

/** One embedded audio/video item extracted for {@link Presentation.appendSlides}. */
interface AvMediaItem {
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
interface OnlineMediaItem {
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
	/**
	 * The generator's presentation-level embedded font faces (`pptx.embedFont`),
	 * carried into the destination deck by {@link Presentation.appendSlides}. Each
	 * face carries its raw `bytes`. Empty when the generator embeds no fonts.
	 */
	embeddedFonts: EmbeddedFont[]
	/**
	 * A notes master to fall back on, present only when at least one slide carries notes.
	 *
	 * A notes slide must bind to a notes master, and a destination template often has
	 * none — a deck authored without speaker notes carries no `notesMaster` part. So the
	 * generator's own notes master rides along, and {@link Presentation.appendSlides}
	 * installs it (with its theme) **only** when the destination has none. When the
	 * destination already has one, it wins and this is discarded — matching the
	 * `importNotes` policy, where the destination's notes styling also wins.
	 */
	notesMaster?: { xml: string; themeXml: string }
}

/**
 * Structural view of a slide producer (a `TsPptx` instance satisfies this).
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

/** One selected native source page and its final zero-based destination position. */
export interface ImportSlidesRequest {
	/** Already-loaded source presentation whose immutable package bytes own the page. */
	source: Presentation
	/**
	 * Zero-based source slide index. Naming one source page in several requests is
	 * how a batch asks for several independent copies of it.
	 */
	sourceIndex: number
	/** Zero-based position in the complete destination slide list after this batch. */
	outputIndex: number
	/**
	 * Carry this page's speaker notes (`notesSlideN.xml`) across, as
	 * `importSlide`'s {@link ImportSlideOptions.importNotes} does. Default false:
	 * the slide copy drops the `notesSlide` relationship, so the imported page
	 * arrives without notes.
	 *
	 * Per request rather than per batch because a stitch mixes sources — the
	 * notes of a library's cover page are worth carrying where a scratch deck's
	 * are not. The deck-wide half of the policy is not per request: a
	 * presentation holds at most one `notesMaster`, so the destination's own is
	 * reused when it has one and the first carried master is installed when it
	 * has none, exactly as `importSlide` and `appendSlides` do.
	 */
	importNotes?: boolean
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
